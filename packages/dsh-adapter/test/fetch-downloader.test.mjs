import test from 'node:test';
import assert from 'node:assert/strict';
import { createHash, randomBytes } from 'node:crypto';
import { existsSync, mkdirSync, mkdtempSync, readFileSync, statSync, utimesSync, writeFileSync } from 'node:fs';
import { createServer } from 'node:http';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { Readable } from 'node:stream';
import { Downloader, DownloadError, fileProblems, markerOf } from '../src/host/fetch/downloader.mjs';
import { LeaseBusy, acquireLease } from '../src/host/fetch/lease.mjs';
import { SourceError, openSource } from '../src/host/fetch/sources.mjs';
import { NPM_CONNECT_TIMEOUT_MS, NPM_IDLE_MS, downgradeProblem, editorNpmSource, kitSources, officeKitItems, readHighest, recordHighest, versionBelow } from '../src/host/fetch/office-kit.mjs';
import { selectProfile, sourcesFor } from '../src/host/fetch/download-sources.mjs';
import { makeTar } from './fixtures/tar-writer.mjs';

const sri = (algorithm, data) => `${algorithm}-${createHash(algorithm).update(data).digest('base64')}`;
const temp = () => mkdtempSync(join(tmpdir(), 'downloader-'));
const kitFiles = { 'offline-kit-runtime': { data: randomBytes(60_000), mode: 0o755 }, 'modoc.js': { data: 'modoc' } };
const archive = makeTar(Object.entries(kitFiles).map(([name, f]) => ({ name, ...f })), { gzip: true });
const files = Object.fromEntries(Object.entries(kitFiles).map(([name, f]) => [name, { sha256: sri('sha256', Buffer.from(f.data)), ...f.mode && { executable: true } }]));
const item = (root, extra = {}) => ({ key: 'kit/0.1.2/darwin-arm64', dir: join(root, 'kit', '0.1.2', 'darwin-arm64'), archive: { file: 'kit.tar.gz', size: archive.length, integrity: sri('sha512', archive) }, files, sources: [{ id: 'release', kind: 'url', base: 'http://release/' }], ...extra });

/** A local HTTP server for the archive: `plan` decides each request (status, partial body, ETag). */
async function server(plan) {
  const requests = [];
  const http = createServer((req, res) => {
    requests.push({ url: req.url, range: req.headers.range, ifRange: req.headers['if-range'], auth: req.headers.authorization ?? req.headers['private-token'] });
    plan(req, res, requests.length);
  });
  await new Promise(resolve => http.listen(0, '127.0.0.1', resolve));
  const base = `http://127.0.0.1:${http.address().port}/`;
  return { base, requests, close: () => new Promise(resolve => { http.closeAllConnections(); http.close(resolve); }) };
}
const send = (res, bytes, { status = 200, etag = '"e1"', headers = {} } = {}) => { res.writeHead(status, { 'Content-Length': bytes.length, ETag: etag, 'Accept-Ranges': 'bytes', ...headers }); res.end(bytes); };
const fast = { delaysMs: [1, 1, 1] };

test('downloads, checks size and sha512, unpacks, checks every file and lands the directory with a .complete marker', async () => {
  const root = temp();
  const http = await server((req, res) => send(res, archive));
  try {
    const events = [];
    const prepared = [];
    const downloader = new Downloader({ root, ...fast });
    downloader.on('progress', event => events.push(event.phase));
    const got = await downloader.ensure(item(root, { sources: [{ id: 'release', kind: 'url', base: http.base }], prepare: async dir => { prepared.push(existsSync(join(dir, 'modoc.js'))); } }));
    assert.equal(got.source, 'release');
    assert.equal(readFileSync(join(got.dir, 'modoc.js'), 'utf8'), 'modoc');
    if (process.platform !== 'win32') assert.ok(statSync(join(got.dir, 'offline-kit-runtime')).mode & 0o100);
    assert.deepEqual(prepared, [true], 'prepare runs on the unpacked files before they go live');
    assert.equal(JSON.parse(readFileSync(join(got.dir, '.complete'), 'utf8')).marker, sri('sha512', archive));
    for (const phase of ['waiting', 'downloading', 'verifying', 'unpacking', 'ready']) assert.ok(events.includes(phase), phase);
    assert.equal(existsSync(join(root, 'downloads', 'kit.tar.gz.part')), false);
    // Second time: the cache, no request.
    const again = await new Downloader({ root }).ensure(item(root, { sources: [{ id: 'release', kind: 'url', base: http.base }] }));
    assert.equal(again.source, 'cache');
    assert.equal(http.requests.length, 1);
    assert.ok(JSON.parse(readFileSync(join(root, 'state.json'), 'utf8')).items['kit/0.1.2/darwin-arm64'].lastUsed);
  } finally { await http.close(); }
});

test('a broken connection resumes from the .part file with Range and If-Range; a changed ETag starts over', async () => {
  const root = temp();
  const half = 20_000;
  const http = await server((req, res, n) => {
    if (n === 1) { res.writeHead(200, { 'Content-Length': archive.length, ETag: '"e1"', 'Accept-Ranges': 'bytes' }); res.write(archive.subarray(0, half), () => setTimeout(() => res.destroy(), 100)); return; }
    const from = Number(/bytes=(\d+)-/.exec(req.headers.range ?? '')?.[1] ?? 0);
    if (from && req.headers['if-range'] === '"e1"') return send(res, archive.subarray(from), { status: 206, headers: { 'Content-Range': `bytes ${from}-${archive.length - 1}/${archive.length}` } });
    send(res, archive);
  });
  try {
    const got = await new Downloader({ root, ...fast }).ensure(item(root, { sources: [{ id: 'release', kind: 'url', base: http.base }] }));
    assert.equal(got.source, 'release');
    assert.equal(http.requests.length, 2);
    const resumedAt = Number(/^bytes=(\d+)-$/.exec(http.requests[1].range ?? '')?.[1]);
    assert.ok(resumedAt > 0 && resumedAt <= half, `resumed at ${http.requests[1].range}`);
    assert.equal(http.requests[1].ifRange, '"e1"');
  } finally { await http.close(); }
});

test('a corrupted download is thrown away and fetched again; a source that keeps failing gives way to the next one', async () => {
  const root = temp();
  const bad = Buffer.from(archive); bad[bad.length - 30] ^= 0xff;
  const broken = await server((req, res) => send(res, bad));
  const notFound = await server((req, res) => { res.writeHead(404); res.end(); });
  const good = await server((req, res) => send(res, archive));
  try {
    const logs = [];
    const got = await new Downloader({ root, ...fast, log: (level, text) => logs.push(text) }).ensure(item(root, { sources: [
      { id: 'missing', kind: 'url', base: notFound.base }, { id: 'broken', kind: 'url', base: broken.base }, { id: 'good', kind: 'url', base: good.base }] }));
    assert.equal(got.source, 'good');
    assert.equal(notFound.requests.length, 1, 'a 404 is not retried');
    assert.equal(broken.requests.length, 3, 'a bad archive is retried, then the next source');
    assert.ok(logs.some(text => /sha512 does not match/.test(text)));
  } finally { await Promise.all([broken.close(), notFound.close(), good.close()]); }
});

test('when every source fails the error names each, is recorded in state.json, and a later call retries', async () => {
  const root = temp();
  const down = await server((req, res) => { res.writeHead(503); res.end(); });
  try {
    const failed = [];
    const downloader = new Downloader({ root, ...fast });
    downloader.on('progress', event => { if (event.phase === 'failed') failed.push(event.error); });
    await assert.rejects(downloader.ensure(item(root, { sources: [{ id: 'release', kind: 'url', base: down.base }, { id: 'other', kind: 'unknown' }] })),
      error => error instanceof DownloadError && /release: .*HTTP 503/.test(error.message) && /other: unknown source kind/.test(error.message));
    assert.equal(down.requests.length, 3);
    assert.equal(failed.length, 1);
    assert.match(downloader.lastError('kit/0.1.2/darwin-arm64').message, /HTTP 503/);
  } finally { await down.close(); }
});

test('files that differ from the manifest never go live, and neither does a Kit whose self test fails', async () => {
  const root = temp();
  const http = await server((req, res) => send(res, archive));
  try {
    const wrong = { ...files, 'modoc.js': { sha256: sri('sha256', Buffer.from('other')) } };
    await assert.rejects(new Downloader({ root, delaysMs: [1] }).ensure(item(root, { files: wrong, sources: [{ id: 'release', kind: 'url', base: http.base }] })), /changed modoc\.js/);
    await assert.rejects(new Downloader({ root, delaysMs: [1] }).ensure(item(root, { files: { ...files, extra: { sha256: 'sha256-x' } }, sources: [{ id: 'release', kind: 'url', base: http.base }] })), /missing extra/);
    await assert.rejects(new Downloader({ root, delaysMs: [1] }).ensure(item(root, { sources: [{ id: 'release', kind: 'url', base: http.base }], prepare: async () => { throw Object.assign(new Error('the document engine does not start'), { retry: false }); } })), /does not start/);
    assert.equal(existsSync(join(root, 'kit', '0.1.2', 'darwin-arm64')), false);
  } finally { await http.close(); }
});

test('an archive put into downloads/ by hand is used when its sha512 matches', async () => {
  const root = temp();
  mkdirSync(join(root, 'downloads'), { recursive: true });
  writeFileSync(join(root, 'downloads', 'whatever-name.tgz'), archive);
  const got = await new Downloader({ root }).ensure(item(root, { sources: [{ id: 'release', kind: 'url', base: 'http://127.0.0.1:9/' }] }));
  assert.equal(got.source, 'manual');
  assert.equal(existsSync(join(root, 'downloads', 'whatever-name.tgz')), false);
});

test('a bundled copy (offline package) is used in place and prepared once; one download serves concurrent callers', async () => {
  const root = temp();
  const bundledDir = join(root, 'package', 'resources', 'kit');
  mkdirSync(bundledDir, { recursive: true });
  for (const [name, f] of Object.entries(kitFiles)) writeFileSync(join(bundledDir, name), f.data);
  let prepared = 0;
  const downloader = new Downloader({ root });
  const bundled = item(root, { bundledDir, prepare: async () => { prepared++; } });
  assert.deepEqual(await downloader.ensure(bundled), { dir: bundledDir, source: 'bundled' });
  await downloader.ensure(bundled);
  assert.equal(prepared, 1);
  let opened = 0;
  const slow = new Downloader({ root: temp(), open: async () => { opened++; await new Promise(r => setTimeout(r, 30)); return { body: Readable.from([archive]), resumed: false, format: 'tgz' }; } });
  const one = item(slow.root);
  const [a, b] = await Promise.all([slow.ensure(one), slow.ensure(one)]);
  assert.equal(opened, 1);
  assert.equal(a.dir, b.dir);
});

test('two processes: the second waits on the lease and then finds the item complete; a stale lease is taken over', async () => {
  const root = temp();
  const held = acquireLease(join(root, '.lock'));
  assert.throws(() => acquireLease(join(root, '.lock')), LeaseBusy);
  let opened = 0;
  const waiting = new Downloader({ root, lease: { pollMs: 10 }, open: async () => { opened++; return { body: Readable.from([archive]), resumed: false, format: 'tgz' }; } });
  const pending = waiting.ensure(item(root));
  await new Promise(r => setTimeout(r, 40));
  // The holder finishes the same item meanwhile.
  await new Downloader({ root, lease: { holder: 'x' }, open: async () => ({ body: Readable.from([archive]), resumed: false, format: 'tgz' }) }).install(item(root), { id: 'other' }, Readable.from([archive]), 'tgz');
  held.release();
  assert.equal((await pending).source, 'cache');
  assert.equal(opened, 0);
  // Stale: not touched for longer than staleMs.
  const lockFile = join(root, '.lock2');
  const old = acquireLease(lockFile, { renewMs: 1e9 });
  const past = new Date(Date.now() - 120_000);
  utimesSync(lockFile, past, past);
  const taken = acquireLease(lockFile);
  assert.notEqual(taken.holder, old.holder);
  old.release();
  assert.ok(existsSync(lockFile), 'the old holder must not remove a lease it lost');
  taken.release();
});

test('a body that stops sending fails after idleMs; an aborted signal stops waiting', async () => {
  const root = temp();
  const stall = await server((req, res) => { res.writeHead(200, { 'Content-Length': archive.length }); res.write(archive.subarray(0, 100)); });
  try {
    await assert.rejects(new Downloader({ root, delaysMs: [1], idleMs: 50 }).ensure(item(root, { sources: [{ id: 'release', kind: 'url', base: stall.base }] })), /no data for/);
    const controller = new AbortController();
    const pending = new Downloader({ root: temp(), delaysMs: [1], idleMs: 10_000 }).ensure(item(root, { sources: [{ id: 'release', kind: 'url', base: stall.base }] }), { signal: controller.signal });
    setTimeout(() => controller.abort(new Error('stopped')), 50);
    await assert.rejects(pending, /stopped/);
  } finally { await stall.close(); }
});

test('budgetMs gives up waiting but the download goes on and lands', async () => {
  const root = temp();
  let release;
  const gate = new Promise(r => { release = r; });
  const downloader = new Downloader({ root, open: async () => { await gate; return { body: Readable.from([archive]), resumed: false, format: 'tgz' }; } });
  await assert.rejects(downloader.ensure(item(root), { budgetMs: 20 }), error => error.code === 'still-downloading');
  release();
  await downloader.inflight.get('kit/0.1.2/darwin-arm64');
  assert.ok(downloader.isComplete(item(root)));
});

test('sources: npm checks the registry integrity', async () => {
  const http = await server((req, res) => {
    if (req.url.startsWith('/registry/')) return send(res, Buffer.from(JSON.stringify({ versions: { '0.1.2': { dist: { tarball: `http://127.0.0.1:${req.socket.localPort}/tarball.tgz`, integrity: 'sha512-other' } } } })));
    send(res, archive);
  });
  try {
    const it = { file: 'kit.tar.gz', integrity: sri('sha512', archive) };
    await assert.rejects(openSource({ kind: 'npm', registry: `${http.base}registry/`, package: '@officesdk/kit-darwin-arm64', version: '0.1.2' }, it), error => error instanceof SourceError && error.code === 'integrity');
  } finally { await http.close(); }
});

test('office-kit items: paths per version and platform, sources with a mirror first, unsupported platforms', () => {
  const manifest = { officeKit: { version: '0.1.2', protocol: '5' }, sources: [{ id: 'release', kind: 'url', base: 'https://release.example/0.1.2/' }],
    kit: { 'darwin-arm64': { file: 'k.tgz', size: 1, integrity: 'sha512-k', binary: 'offline-kit-runtime', status: 'stable', files: { 'offline-kit-runtime': { sha256: 'x', executable: true } } } },
    editor: { file: 'e.tgz', size: 2, integrity: 'sha512-e', root: 'package/dist/resources', files: { 'catalog.json': { sha256: 'y' } } } };
  const items = officeKitItems(manifest, { root: '/data/office-kit', packageRoot: '/pkg', host: { platform: 'darwin', arch: 'arm64' }, mirror: 'https://mirror.example/okit/' });
  assert.equal(items.kit.dir, join('/data/office-kit', 'kit', '0.1.2', 'darwin-arm64'));
  assert.equal(items.editor.dir, join('/data/office-kit', 'editor', '0.1.2'));
  assert.equal(items.editor.strip, 'package/dist/resources');
  assert.equal(items.kit.bundledDir, join('/pkg', 'resources', 'kit'));
  assert.deepEqual(items.kit.sources.map(s => s.id), ['mirror', 'release']);
  assert.equal(items.kit.sources[0].base, 'https://mirror.example/okit/0.1.2/');
  assert.deepEqual(officeKitItems(manifest, { root: '/d', host: { platform: 'linux', arch: 'x64' } }).unsupported, { platform: 'linux-x64', reason: 'platform', supported: [{ platform: 'darwin-arm64', status: 'stable' }] });
  assert.deepEqual(kitSources(manifest).map(s => s.kind), ['url']);
  assert.notEqual(markerOf({ files: { a: { sha256: '1' } } }), markerOf({ files: { a: { sha256: '2' } } }));
  // A skill's marker is its file list, whether its pack's sha512 is known yet or not (R2-F1).
  const skill = { key: 'skill/x/0.0.1', files: { 'SKILL.md': { sha512: 's' } } };
  assert.equal(markerOf({ ...skill, archive: { integrity: 'sha512-p' } }), markerOf(skill));
  assert.equal(markerOf({ key: 'kit/1', files: {}, archive: { integrity: 'sha512-k' } }), 'sha512-k');
  assert.deepEqual(fileProblems({ a: { sha512: 'x' } }, { a: { sha512: 'x', executable: true }, b: {} }), ['mode a', 'unexpected b']);
});

test('downgrade guard: a Kit below the floor the newest Kit recorded, or one that cannot read its index, is refused', async () => {
  const root = temp();
  const manifest = v => ({ officeKit: { version: v, indexVersion: v === '0.2.0' ? 4 : 3, readsIndexVersions: v === '0.2.0' ? [3, 4] : [1, 2, 3], downgradeFloor: v === '0.2.0' ? '0.1.5' : '0.1.0' } });
  assert.equal(downgradeProblem(manifest('0.1.2'), await readHighest(root)), undefined);
  await recordHighest(root, manifest('0.1.2'));
  await recordHighest(root, manifest('0.2.0'));
  await recordHighest(root, manifest('0.1.2'));
  assert.equal((await readHighest(root)).version, '0.2.0', 'an older Kit never lowers the record');
  assert.match(downgradeProblem(manifest('0.1.2'), await readHighest(root)).message, /0\.2\.0/);
  assert.equal(downgradeProblem(manifest('0.2.0'), await readHighest(root)), undefined);
  assert.ok(versionBelow('0.1.2', '0.1.10') && !versionBelow('0.1.10', '0.1.2'));
});

test('taking over a stale lease puts back a lease another taker created meanwhile (R1-F8)', () => {
  const root = temp();
  const file = join(root, '.lock');
  writeFileSync(file, '1:old');
  const past = new Date(Date.now() - 120_000);
  utimesSync(file, past, past);
  // Between this taker's staleness check and its rename, another taker replaced the stale lease with its own.
  assert.throws(() => acquireLease(file, { beforeTakeover: () => writeFileSync(file, '2:other-taker') }), LeaseBusy);
  assert.equal(readFileSync(file, 'utf8'), '2:other-taker', 'the other taker keeps its lease');
  // A lease still stale when moved is taken over.
  writeFileSync(file, '3:old'); utimesSync(file, past, past);
  const taken = acquireLease(file);
  assert.match(readFileSync(file, 'utf8'), new RegExp(`^${process.pid}:`));
  taken.release();
});

test('a complete, matching .part left by a failed later step installs without downloading again (R2-F4)', async () => {
  const root = temp();
  mkdirSync(join(root, 'downloads'), { recursive: true });
  writeFileSync(join(root, 'downloads', 'kit.tar.gz.part'), archive);
  writeFileSync(join(root, 'downloads', 'kit.tar.gz.part.json'), JSON.stringify({ source: 'release', etag: '"e1"', integrity: sri('sha512', archive) }));
  let opened = 0;
  const got = await new Downloader({ root, open: async () => { opened++; throw new Error('must not download'); } }).ensure(item(root));
  assert.equal(got.source, 'release');
  assert.equal(opened, 0);
});

// The editor from npm (office-kit 0.3: the release carries no editor); a drill profile may add a source after it.
const npmManifest = (kit = {}) => ({ officeKit: { version: '0.1.3', protocol: '5' }, sources: [{ id: 'release', kind: 'url', base: 'https://release.example/0.1.3/' }],
  kit: { 'darwin-arm64': { file: 'k.tgz', size: 1, integrity: 'sha512-k', binary: 'offline-kit-runtime', status: 'stable', files: { 'offline-kit-runtime': { sha256: 'x', executable: true } } }, ...kit },
  editor: { file: 'offline-kit-editor-0.1.3.tgz', size: 2, integrity: 'sha512-e', root: 'package/dist/resources', files: { 'catalog.json': { sha256: 'y' } },
    npm: { registry: 'https://npm.example.invalid/', package: '@example/offline-kit-editor', version: '0.1.3' } } });

test('the editor comes from npm by default; a profile may set its own limits and a later source; the Kit never comes from npm; off and a URL change it', () => {
  const mac = { platform: 'darwin', arch: 'arm64' };
  assert.deepEqual(officeKitItems(npmManifest(), { root: '/d', host: mac }).editor.sources.map(s => s.id), ['npm'], 'by default npm is the editor\'s only source');
  const drill = sourcesFor('public');
  drill.profiles.public.editor = [{ id: 'npm', connectTimeoutMs: 3000, idleMs: 8000, attempts: 1 }, 'mirror'];
  const profile = selectProfile(drill, {});
  const items = officeKitItems(npmManifest(), { root: '/d', host: mac, profile });
  assert.deepEqual(items.editor.sources.map(s => s.id), ['npm'], 'no mirror set: npm alone');
  assert.deepEqual(items.kit.sources.map(s => s.id), ['release']);
  assert.deepEqual(items.editor.sources[0], { id: 'npm', kind: 'npm', registry: 'https://npm.example.invalid/', package: '@example/offline-kit-editor', version: '0.1.3', attempts: 1, connectTimeoutMs: 3000, idleMs: 8000 });
  assert.deepEqual([NPM_CONNECT_TIMEOUT_MS, NPM_IDLE_MS], [15_000, 30_000], 'the defaults: npm is the last source');
  assert.deepEqual(officeKitItems(npmManifest(), { root: '/d', host: mac, profile, mirror: 'https://m/' }).editor.sources.map(s => s.id), ['npm', 'mirror'], 'the profile\'s order: npm, then the mirror');
  assert.deepEqual(officeKitItems(npmManifest(), { root: '/d', host: mac, profile, npm: 'off', mirror: 'https://m/' }).editor.sources.map(s => s.id), ['mirror']);
  assert.equal(officeKitItems(npmManifest(), { root: '/d', host: mac, profile, npm: 'https://10.255.255.1/' }).editor.sources[0].registry, 'https://10.255.255.1/');
  assert.equal(editorNpmSource({ ...npmManifest(), editor: { ...npmManifest().editor, npm: undefined } }), undefined, 'a manifest without editor.npm (0.1.2) keeps the release only');
});

test('test platforms resolve their own binary; a platform the manifest lacks is unsupported and the message names the test ones', async () => {
  const test = (exe = '') => ({ file: 't.tgz', size: 1, integrity: 'sha512-t', binary: `offline-kit-runtime${exe}`, status: 'test', files: { [`offline-kit-runtime${exe}`]: { sha256: 'x', executable: true } } });
  const manifest = npmManifest({ 'linux-x64': test(), 'win32-arm64': test('.exe') });
  assert.equal(officeKitItems(manifest, { root: '/d', host: { platform: 'linux', arch: 'x64' } }).kit.binary, 'offline-kit-runtime');
  assert.equal(officeKitItems(manifest, { root: '/d', host: { platform: 'win32', arch: 'arm64' } }).kit.binary, 'offline-kit-runtime.exe');
  assert.equal(officeKitItems(manifest, { root: '/d', host: { platform: 'win32', arch: 'arm64' } }).kit.status, 'test');
  const { unsupported } = officeKitItems(manifest, { root: '/d', host: { platform: 'freebsd', arch: 'x64' } });
  assert.equal(unsupported.reason, 'platform');
  const { unsupportedError } = await import('../src/host/kit-platform.mjs');
  const message = unsupportedError(unsupported).message;
  assert.match(message, /只支持 darwin-arm64、linux-x64（测试）、win32-arm64（测试），当前系统是 freebsd-x64/);
  assert.match(message, /supports darwin-arm64, linux-x64 \(test\), win32-arm64 \(test\) only/);
});

/** An editor item fetched from a local npm registry stand-in (`npm`) and the release (`release`). */
function npmItem(root, npmSource, releaseBase) {
  return { key: 'editor/0.1.3', label: 'editor', dir: join(root, 'editor', '0.1.3'), archive: { file: 'editor.tgz', size: archive.length, integrity: sri('sha512', archive) }, files,
    sources: [npmSource, { id: 'release', kind: 'url', base: releaseBase }] };
}

test('npm unreachable: packument with no answer gives way to the release within the npm limit, once, and says it fell back', async () => {
  const root = temp();
  const release = await server((req, res) => send(res, archive));
  const logs = [];
  const events = [];
  try {
    // A registry address that accepts nothing in time (outside the company network): fetch never answers.
    const hang = (url, { signal }) => (url.includes('registry') ? new Promise((_, reject) => signal.addEventListener('abort', () => reject(signal.reason), { once: true })) : fetch(url, { signal }));
    const downloader = new Downloader({ root, delaysMs: [2000, 8000, 30000], fetchImpl: hang, log: (level, text) => logs.push(text) });
    downloader.on('progress', event => events.push(event));
    const started = Date.now();
    const result = await downloader.ensure(npmItem(root, { id: 'npm', kind: 'npm', registry: 'https://registry.example/', package: '@example/offline-kit-editor', version: '0.1.3', attempts: 1, connectTimeoutMs: 200, idleMs: 200 }, release.base));
    const took = Date.now() - started;
    assert.equal(result.source, 'release');
    assert.ok(took < 1500, `fell back in ${took} ms (one try, no 2 s retry delay)`);
    assert.ok(logs.some(text => /editor\/0\.1\.3: npm unavailable \(no answer from https:\/\/registry\.example\/ in 0\.2 s\); falling back to release/.test(text)), logs.join('\n'));
    assert.ok(events.some(event => event.phase === 'downloading' && event.source === 'release' && event.fallback?.[0] === 'npm'), 'progress names the source and what it fell back from');
  } finally { await release.close(); }
});

test('npm body that stalls gives way after the npm idle limit; npm failing to open keeps the release\'s .part', async () => {
  const root = temp();
  const npm = await server((req, res) => {
    if (req.url.startsWith('/registry/')) return send(res, Buffer.from(JSON.stringify({ versions: { '0.1.3': { dist: { tarball: `http://127.0.0.1:${req.socket.localPort}/editor.tgz` } } } })));
    res.writeHead(200, { 'Content-Length': archive.length }); res.write(archive.subarray(0, 100));
  });
  const release = await server((req, res) => send(res, archive));
  try {
    const started = Date.now();
    const result = await new Downloader({ root, delaysMs: [2000, 8000, 30000], idleMs: 60_000 }).ensure(npmItem(root, { id: 'npm', kind: 'npm', registry: `${npm.base}registry/`, package: '@example/offline-kit-editor', version: '0.1.3', attempts: 1, connectTimeoutMs: 200, idleMs: 200 }, release.base));
    assert.equal(result.source, 'release');
    assert.ok(Date.now() - started < 1500, 'the npm idle limit applies, not the 60 s one');
    // An the release .part from an earlier run survives an npm source that fails before sending a byte.
    const root2 = temp();
    mkdirSync(join(root2, 'downloads'), { recursive: true });
    writeFileSync(join(root2, 'downloads', 'editor.tgz.part'), archive.subarray(0, 1000));
    writeFileSync(join(root2, 'downloads', 'editor.tgz.part.json'), JSON.stringify({ source: 'release', etag: '"e1"', integrity: sri('sha512', archive) }));
    const resumed = await server((req, res) => (req.headers.range ? send(res, archive.subarray(1000), { status: 206, headers: { 'Content-Range': `bytes 1000-${archive.length - 1}/${archive.length}` } }) : send(res, archive)));
    try {
      const refused = async () => { throw new SourceError('https://registry.example/: ENOTFOUND', { code: 'network' }); };
      const open = (source, it, context) => (source.kind === 'npm' ? refused() : openSource(source, it, context));
      assert.equal((await new Downloader({ root: root2, delaysMs: [1], open }).ensure(npmItem(root2, { id: 'npm', kind: 'npm', registry: 'https://registry.example/', package: 'p', version: '0.1.3', attempts: 1 }, resumed.base))).source, 'release');
      assert.equal(resumed.requests[0].range, 'bytes=1000-', 'the release resumed from its own .part');
    } finally { await resumed.close(); }
  } finally { await npm.close(); await release.close(); }
});

test('OFFICESDK_EDITOR_NPM: off in any case (or false/0/none) leaves npm out; a value that is neither warns; npm errors hide registry credentials (R1-F2, R1-F3)', async () => {
  const mac = { platform: 'darwin', arch: 'arm64' };
  // Off leaves npm out: by default the editor then has only a mirror, when OFFICESDK_KIT_MIRROR sets one.
  for (const value of ['off', 'OFF', ' Off ', 'false', '0', 'none']) assert.deepEqual(officeKitItems(npmManifest(), { root: '/d', host: mac, npm: value }).editor.sources.map(s => s.id), [], value);
  assert.deepEqual(officeKitItems(npmManifest(), { root: '/d', host: mac, npm: 'off', mirror: 'https://m/' }).editor.sources.map(s => s.id), ['mirror']);
  const logs = [];
  assert.equal(editorNpmSource(npmManifest(), 'registry.example', (level, text) => logs.push(`${level} ${text}`)).registry, 'https://npm.example.invalid/');
  assert.match(logs[0], /^warn OFFICESDK_EDITOR_NPM=registry\.example is neither off nor an http\(s\) URL/);
  const http = await server((req, res) => (req.url.includes('missing') ? send(res, Buffer.from('{}'), { status: 404 }) : send(res, Buffer.from(JSON.stringify({ versions: {} })))));
  try {
    const registry = http.base.replace('http://', 'http://user:s3cret@');
    for (const pkg of ['@example/missing', '@example/offline-kit-editor']) {
      await assert.rejects(openSource({ kind: 'npm', registry, package: pkg, version: '0.1.3' }, { file: 'e.tgz' }), error => error instanceof SourceError && !/s3cret/.test(error.message) && /127\.0\.0\.1/.test(error.message));
    }
  } finally { await http.close(); }
});
