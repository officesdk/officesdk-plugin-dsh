import test from 'node:test';
import assert from 'node:assert/strict';
import { createHash, randomBytes } from 'node:crypto';
import { EventEmitter } from 'node:events';
import { mkdtempSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { Readable } from 'node:stream';
import { Downloader } from '../src/host/fetch/downloader.mjs';
import { recordHighest } from '../src/host/fetch/office-kit.mjs';
import { OfficeKitRuntime } from '../src/host/fetch/runtime.mjs';
import { KitDocuments, KitProcess } from '../../kit-host/src/index.mjs';
import { selectProfile, sourcesFor } from '../src/host/fetch/download-sources.mjs';
import { makeTar } from './fixtures/tar-writer.mjs';

const sri = (algorithm, data) => `${algorithm}-${createHash(algorithm).update(data).digest('base64')}`;
const runtimeBytes = randomBytes(30_000);
const kitTar = makeTar([{ name: 'offline-kit-runtime', data: runtimeBytes, mode: 0o755 }], { gzip: true });
const editorTar = makeTar([{ name: 'package/dist/resources/catalog.json', data: '{}' }], { gzip: true });
const manifest = {
  v: 1, officeKit: { version: '0.1.2', protocol: '5', indexVersion: 3, readsIndexVersions: [1, 2, 3], downgradeFloor: '0.1.0' },
  sources: [{ id: 'release', kind: 'url', base: 'https://release.example/0.1.2/' }],
  kit: { 'darwin-arm64': { status: 'stable', file: 'kit.tgz', size: kitTar.length, sha256: 'x', integrity: sri('sha512', kitTar), binary: 'offline-kit-runtime',
    files: { 'offline-kit-runtime': { sha256: sri('sha256', runtimeBytes), executable: true } } } },
  editor: { file: 'editor.tgz', size: editorTar.length, sha256: 'x', integrity: sri('sha512', editorTar), root: 'package/dist/resources', files: { 'catalog.json': { sha256: sri('sha256', '{}') } },
    npm: { registry: 'https://registry.example/', package: '@example/core-editor', version: '0.1.2' } },
};
const host = { platform: 'darwin', arch: 'arm64' };
// A valid profile that lists another editor source after npm (a drill): npm with short limits, then a mirror (MIRROR).
const npmDrill = () => { const doc = sourcesFor('public'); doc.profiles.public.editor = [{ id: 'npm', connectTimeoutMs: 3000, idleMs: 8000, attempts: 1 }, 'mirror']; return selectProfile(doc, {}); };
const MIRROR = 'https://mirror.example/office-kit';

function runtime({ open, prepare = async () => {} } = {}) {
  const dataRoot = mkdtempSync(join(tmpdir(), 'okrt-'));
  const downloader = new Downloader({ root: join(dataRoot, 'office-kit'), delaysMs: [1], open: open ?? (async (_source, item) => ({ body: Readable.from([item.archive.file === 'kit.tgz' ? kitTar : editorTar]), resumed: false, format: 'tgz' })) });
  const rt = new OfficeKitRuntime({ manifest, dataRoot, packageRoot: join(dataRoot, 'pkg'), host, downloader });
  rt.kitItem.prepare = prepare;
  return rt;
}

test('the engine downloads on first need: the Kit path comes back, the editor resources become the static root, status says ready', async () => {
  const rt = runtime();
  assert.equal(rt.ready(), false);
  assert.equal(rt.status().state, 'idle');
  assert.equal(rt.editorDir, undefined);
  const binary = await rt.binary();
  assert.equal(binary, join(rt.root, 'kit', '0.1.2', 'darwin-arm64', 'offline-kit-runtime'));
  assert.equal(rt.editorDir, join(rt.root, 'editor', '0.1.2'));
  assert.equal(rt.ready(), true);
  assert.deepEqual(rt.status(), { state: 'ready', received: kitTar.length + editorTar.length, total: kitTar.length + editorTar.length });
});

test('a packaged editor (261112-bundle-editor) is used in place: only the Kit downloads, nothing of the editor is written to the data directory, the sizes count the Kit alone', async () => {
  const { mkdirSync, writeFileSync, existsSync } = await import('node:fs');
  const opened = [];
  const dataRoot = mkdtempSync(join(tmpdir(), 'okrt-'));
  const packageRoot = join(dataRoot, 'pkg');
  mkdirSync(join(packageRoot, 'resources', 'office-kit'), { recursive: true });
  writeFileSync(join(packageRoot, 'resources', 'office-kit', 'catalog.json'), '{}');
  const downloader = new Downloader({ root: join(dataRoot, 'office-kit'), delaysMs: [1], open: async (source, item) => { opened.push(item.archive.file); return { body: Readable.from([item.archive.file === 'kit.tgz' ? kitTar : editorTar]), resumed: false, format: 'tgz' }; } });
  const rt = new OfficeKitRuntime({ manifest, dataRoot, packageRoot, host, downloader });
  rt.kitItem.prepare = async () => {};
  assert.equal(rt.ready(), false, 'the Kit is still missing');
  assert.deepEqual(rt.status(), { state: 'idle', received: 0, total: kitTar.length });
  await rt.binary();
  assert.deepEqual(opened, ['kit.tgz'], 'the editor is never fetched');
  assert.equal(rt.editorDir, join(packageRoot, 'resources', 'office-kit'), 'the static root serves the packaged copy');
  assert.equal(existsSync(join(rt.root, 'editor')), false, 'no editor directory in the data directory');
  assert.equal(rt.downloader.readState().items?.['editor/0.1.2'], undefined, 'and no use recorded for one');
  assert.equal(rt.ready(), true);
  assert.deepEqual(rt.status(), { state: 'ready', received: kitTar.length, total: kitTar.length });
});

test('with a packaged editor, cleanup lets every downloaded editor age out; while the packaged copy is incomplete the download in use stays (261112-bundle-editor)', async () => {
  const { mkdirSync, writeFileSync, existsSync, rmSync } = await import('node:fs');
  const rt = runtime();
  await rt.binary();
  const packaged = join(rt.root, '..', 'pkg', 'resources', 'office-kit');
  const editorDir = version => join(rt.root, 'editor', version);
  mkdirSync(editorDir('0.1.1'), { recursive: true });
  const old = Date.now() - 90 * 24 * 3600_000;
  const age = () => {
    const state = rt.downloader.readState();
    state.items['editor/0.1.2'] = { ...state.items['editor/0.1.2'], lastUsed: old + 1000 };
    state.items['editor/0.1.1'] = { dir: editorDir('0.1.1'), lastUsed: old };
    writeFileSync(join(rt.root, 'state.json'), JSON.stringify(state));
  };
  age();
  // Not packaged (this package lacks resources/office-kit): this version and the previous one used stay.
  assert.deepEqual(await rt.prune({ kitWork: mkdtempSync(join(tmpdir(), 'kitwork-')) }), []);
  mkdirSync(packaged, { recursive: true });
  writeFileSync(join(packaged, 'catalog.json'), '{}');
  age();
  assert.deepEqual((await rt.prune({ kitWork: mkdtempSync(join(tmpdir(), 'kitwork-')) })).sort(), [editorDir('0.1.1'), editorDir('0.1.2')].sort());
  assert.ok(existsSync(rt.kitItem.dir), 'the Kit stays');
  // A packaged copy that lost a file (security software) is not packaged: the release download it falls back to is kept.
  rmSync(join(packaged, 'catalog.json'));
  await rt.binary();
  age();
  assert.deepEqual(await rt.prune({ kitWork: mkdtempSync(join(tmpdir(), 'kitwork-')) }), [], 'the fallback download in use stays');
  assert.ok(existsSync(editorDir('0.1.2')));
});

test('a failed download says so in Chinese and English with the reason; status reports it; a later call tries again', async () => {
  let fail = true;
  const rt = runtime({ open: async (_source, item) => { if (fail) throw Object.assign(new Error('HTTP 503'), { retry: false }); return { body: Readable.from([item.archive.file === 'kit.tgz' ? kitTar : editorTar]), resumed: false, format: 'tgz' }; } });
  await assert.rejects(rt.binary(), error => error.code === 'download-failed' && /文档引擎下载失败/.test(error.message) && /HTTP 503/.test(error.message) && /retry when online/.test(error.message) && !/offline package/.test(error.message));
  assert.equal(rt.status().state, 'failed');
  fail = false;
  await rt.binary();
  assert.equal(rt.status().state, 'ready');
});

test('a machine that ran a newer engine refuses an older one below its floor, without downloading', async () => {
  let opened = 0;
  const rt = runtime({ open: async () => { opened++; throw new Error('no'); } });
  await recordHighest(rt.root, { officeKit: { version: '0.3.0', indexVersion: 5, downgradeFloor: '0.2.0' } });
  await assert.rejects(rt.binary(), error => error.code === 'kit-downgrade' && /0\.3\.0/.test(error.message));
  assert.equal(opened, 0);
  // markStarted records this engine (a lower one never lowers the record).
  const fresh = runtime();
  await fresh.markStarted();
  assert.match(JSON.stringify(await import('../src/host/fetch/office-kit.mjs').then(m => m.readHighest(fresh.root))), /"0\.1\.2"/);
});

test('a platform the manifest has no Kit for is unsupported, and binary() says so', async () => {
  const dataRoot = mkdtempSync(join(tmpdir(), 'okrt-'));
  const rt = new OfficeKitRuntime({ manifest, dataRoot, host: { platform: 'linux', arch: 'x64' } });
  assert.deepEqual(rt.unsupported, { platform: 'linux-x64', reason: 'platform', supported: [{ platform: 'darwin-arm64', status: 'stable' }] });
  assert.equal(rt.status().state, 'unsupported');
  await assert.rejects(rt.binary(), error => error.code === 'unsupported-platform');
});

test('KitProcess resolves a binary function before spawning; its failure reaches the open as an explained unavailable error', async () => {
  const spawned = [];
  const kit = new KitProcess({ binary: async () => { throw Object.assign(new Error('文档引擎下载失败：HTTP 503'), { code: 'download-failed' }); }, work: mkdtempSync(join(tmpdir(), 'kitwork-')), spawn: (...args) => { spawned.push(args); return new EventEmitter(); } });
  await assert.rejects(kit.start(), /文档引擎下载失败/);
  assert.equal(spawned.length, 0);
  const documents = new KitDocuments({ kit, authorize: async () => ({}), stat: async () => ({ absolutePath: tmpdir(), version: '1' }), writeAccess: async () => 'allowed', exportRoot: tmpdir() });
  await assert.rejects(documents.open({ workspaceRoot: '/w' }, 'a.docx', { id: 'p', principal: 'owner' }, { mode: 'view' }),
    error => error.name === 'KitError' && error.code === 'unavailable' && /文档引擎下载失败/.test(error.message) && error.details.engine === 'download-failed');
  await kit.stop?.();
});

test('a Kit stopped while its binary is still downloading is never spawned (R1-F2)', async () => {
  let resolveBinary;
  const spawned = [];
  const kit = new KitProcess({ binary: () => new Promise(resolve => { resolveBinary = resolve; }), work: mkdtempSync(join(tmpdir(), 'kitwork-')), spawn: (...args) => { spawned.push(args); return new EventEmitter(); } });
  const starting = kit.start();
  await new Promise(r => setTimeout(r, 10));
  await kit.stop();
  resolveBinary('/somewhere/offline-kit-runtime');
  await assert.rejects(starting, error => error.code === 'kit-suspended');
  assert.equal(spawned.length, 0);
});

test('a Kit that will not start is reported as such, once, without downloading again (R1-F7)', async () => {
  let opened = 0;
  const rt = runtime({ open: async (_source, item) => { if (item.archive.file === 'kit.tgz') opened++; return { body: Readable.from([item.archive.file === 'kit.tgz' ? kitTar : editorTar]), resumed: false, format: 'tgz' }; },
    prepare: async () => { throw Object.assign(new Error('timed out after 120 s'), { code: 'kit-selftest', retry: false, final: true }); } });
  rt.kitItem.sources = [{ id: 'a', kind: 'url', base: 'x' }, { id: 'b', kind: 'url', base: 'y' }];
  await assert.rejects(rt.binary(), error => error.code === 'kit-selftest' && /无法启动/.test(error.message) && /timed out/.test(error.message));
  assert.equal(opened, 1, 'the Kit archive once: no second try, no second source');
});

test('cleanup keeps this version, the previous one used, and a running Kit (R1-F9)', async () => {
  const rt = runtime();
  await rt.binary();
  const { mkdirSync, writeFileSync, existsSync } = await import('node:fs');
  const dir = version => join(rt.root, 'kit', version, 'darwin-arm64');
  for (const version of ['0.0.8', '0.0.9', '0.1.0', '0.1.1']) mkdirSync(dir(version), { recursive: true });
  const old = Date.now() - 90 * 24 * 3600_000;
  const state = rt.downloader.readState();
  state.items['kit/0.0.8'] = { dir: dir('0.0.8'), lastUsed: old };
  state.items['kit/0.0.9'] = { dir: dir('0.0.9'), lastUsed: old + 1000 };
  state.items['kit/0.1.0'] = { dir: dir('0.1.0'), lastUsed: old + 2000 };
  writeFileSync(join(rt.root, 'state.json'), JSON.stringify(state));
  const kitWork = mkdtempSync(join(tmpdir(), 'kitwork-'));
  writeFileSync(join(kitWork, '.officesdk-kit.running'), JSON.stringify({ pid: process.pid, binary: join(dir('0.0.8'), 'offline-kit-runtime') }));
  const removed = await rt.prune({ kitWork });
  assert.deepEqual(removed.sort(), [dir('0.0.9'), dir('0.1.1')].sort());
  assert.ok(existsSync(dir('0.0.8')), 'running');
  assert.ok(existsSync(dir('0.1.0')), 'the previous version used');
  assert.ok(existsSync(rt.kitItem.dir), 'this version');
});

test('the failure shown is the source that was tried, not one skipped for lack of a token (smoke finding)', async () => {
  const rt = runtime({ open: async source => { if (source.kind === 'token-only') throw Object.assign(new Error('no token for this source'), { code: 'no-token', retry: false }); throw Object.assign(new Error('https://release.example/kit.tgz: ECONNREFUSED'), { code: 'network', retry: false }); } });
  rt.kitItem.sources = [{ id: 'release', kind: 'url', base: 'x' }, { id: 'other', kind: 'token-only' }];
  await assert.rejects(rt.binary(), error => /ECONNREFUSED/.test(error.message) && !/no token for this source/.test(error.message));
});

test('while the editor comes from a mirror after npm failed, status names the fallback in its text; a test platform says so (spec 261111-kit-0-1-3)', async () => {
  let release;
  const gate = new Promise(resolve => { release = resolve; });
  const testManifest = { ...manifest, kit: { 'linux-x64': { ...manifest.kit['darwin-arm64'], status: 'test' } },
    editor: { ...manifest.editor, npm: { registry: 'https://npm.example.invalid/', package: '@example/offline-kit-editor', version: '0.1.2' } } };
  const dataRoot = mkdtempSync(join(tmpdir(), 'okrt-'));
  const logs = [];
  const open = async (source, item) => {
    if (source.kind === 'npm') throw Object.assign(new Error('https://npm.example.invalid/: ENOTFOUND'), { code: 'network' });
    if (item.archive.file === 'editor.tgz') {
      async function* slow() { yield editorTar.subarray(0, 10); await gate; yield editorTar.subarray(10); }
      return { body: slow(), resumed: false, format: 'tgz' };
    }
    return { body: Readable.from([kitTar]), resumed: false, format: 'tgz' };
  };
  const downloader = new Downloader({ root: join(dataRoot, 'office-kit'), delaysMs: [1], open, log: (level, text) => logs.push(text) });
  const rt = new OfficeKitRuntime({ manifest: testManifest, dataRoot, packageRoot: join(dataRoot, 'pkg'), host: { platform: 'linux', arch: 'x64' }, downloader, mirror: MIRROR, profile: npmDrill(), log: (level, text) => logs.push(text) });
  rt.kitItem.prepare = async () => {};
  assert.deepEqual(rt.editorItem.sources.map(source => source.id), ['npm', 'mirror']);
  const running = rt.binary();
  // The editor waits for the Kit's lease (polled every second) before it starts.
  for (let i = 0; i < 400 && !(rt.status().fallback); i++) await new Promise(resolve => setTimeout(resolve, 10));
  const status = rt.status();
  if (!status.fallback) release();
  assert.equal(status.state, 'downloading');
  assert.deepEqual(status.fallback, ['npm']);
  assert.equal(status.platformStatus, 'test');
  assert.match(status.text, /npm 源连不上，已改用其他来源下载/);
  assert.match(status.text, /此平台测试中/);
  release();
  assert.equal(await running, join(rt.root, 'kit', '0.1.2', 'linux-x64', 'offline-kit-runtime'));
  assert.ok(logs.some(text => /the document engine for linux-x64 is in test/.test(text)));
  assert.ok(logs.some(text => /falling back to mirror/.test(text)));
  assert.deepEqual(rt.status(), { state: 'ready', received: kitTar.length + editorTar.length, total: kitTar.length + editorTar.length });
});

test('a fallback recorded by an earlier failed download is not shown when the retry comes from npm (R1-F1)', async () => {
  const fallbackManifest = { ...manifest, editor: { ...manifest.editor, npm: { registry: 'https://npm.example.invalid/', package: '@example/offline-kit-editor', version: '0.1.2' } } };
  let online = false;
  let release;
  const gate = new Promise(resolve => { release = resolve; });
  const open = async (source, item) => {
    if (!online) throw Object.assign(new Error(`${source.id}: ENOTFOUND`), { code: 'network', retry: false });
    if (item.archive.file === 'editor.tgz') {
      assert.equal(source.id, 'npm', 'online, the editor comes from npm');
      async function* slow() { yield editorTar.subarray(0, 10); await gate; yield editorTar.subarray(10); }
      return { body: slow(), resumed: false, format: 'tgz' };
    }
    return { body: Readable.from([kitTar]), resumed: false, format: 'tgz' };
  };
  const dataRoot = mkdtempSync(join(tmpdir(), 'okrt-'));
  const downloader = new Downloader({ root: join(dataRoot, 'office-kit'), delaysMs: [1], open });
  const rt = new OfficeKitRuntime({ manifest: fallbackManifest, dataRoot, packageRoot: join(dataRoot, 'pkg'), host, downloader, mirror: MIRROR, profile: npmDrill() });
  rt.kitItem.prepare = async () => {};
  await assert.rejects(rt.binary());
  assert.equal(rt.status().state, 'failed');
  online = true;
  const running = rt.binary();
  for (let i = 0; i < 400 && rt.status().state !== 'downloading'; i++) await new Promise(resolve => setTimeout(resolve, 10));
  for (let i = 0; i < 400 && rt.status().source !== 'npm'; i++) await new Promise(resolve => setTimeout(resolve, 10));
  const status = rt.status();
  release();
  assert.equal(status.source, 'npm');
  assert.equal(status.fallback, undefined);
  assert.doesNotMatch(status.text, /npm 源连不上/);
  await running;
});

test('by default the editor comes from npm alone: when npm fails, the error names the registry, not a backup source', async () => {
  const dataRoot = mkdtempSync(join(tmpdir(), 'okrt-'));
  const open = async (source, item) => {
    if (source.kind === 'npm') throw Object.assign(new Error('https://registry.example/: ENOTFOUND'), { code: 'network' });
    return { body: Readable.from([item.archive.file === 'kit.tgz' ? kitTar : editorTar]), resumed: false, format: 'tgz' };
  };
  const downloader = new Downloader({ root: join(dataRoot, 'office-kit'), delaysMs: [1], open });
  const rt = new OfficeKitRuntime({ manifest, dataRoot, packageRoot: join(dataRoot, 'pkg'), host, downloader });
  rt.kitItem.prepare = async () => {};
  assert.deepEqual(rt.editorItem.sources.map(source => [source.id, source.registry, source.attempts]), [['npm', 'https://registry.example/', 3]]);
  assert.deepEqual(rt.kitItem.sources.map(source => source.base), ['https://release.example/0.1.2/'], 'the Kit from the manifest\'s release sources');
  const error = await rt.binary().catch(caught => caught);
  assert.equal(error.code, 'download-failed');
  assert.match(error.message, /registry\.example/);
  assert.doesNotMatch(error.message, /备用|backup/);
  assert.equal(rt.status().state, 'failed');
  assert.doesNotMatch(String(rt.status().text ?? ''), /备用|其他来源/);
});

test('a platform office-kit releases only a library for (no binary, e.g. OpenHarmony) is unsupported, not a Kit that fails to start', async () => {
  const library = { ...manifest, kit: { ...manifest.kit, 'openharmony-arm64': { status: 'test', file: 'ohos.tgz', size: 1, integrity: 'sha512-x', module: 'liboffline_kit.so', files: {} } } };
  const dataRoot = mkdtempSync(join(tmpdir(), 'okrt-'));
  const rt = new OfficeKitRuntime({ manifest: library, dataRoot, packageRoot: join(dataRoot, 'pkg'), host: { platform: 'openharmony', arch: 'arm64' } });
  assert.equal(rt.unsupported.platform, 'openharmony-arm64');
  assert.deepEqual(rt.unsupported.supported.map(row => row.platform), ['darwin-arm64'], 'only runnable platforms are offered');
  await assert.rejects(rt.binary(), error => error.code === 'unsupported-platform');
  assert.equal(rt.status().state, 'unsupported');
});

