import test from 'node:test';
import assert from 'node:assert/strict';
import { existsSync, mkdirSync, mkdtempSync, readFileSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { PluginUpdater, newer } from '../src/host/updater.mjs';
import { ChannelStore, NPMJS, NPMMIRROR, channelVersion, fetchPackument, npmInstall, npmrcRegistry, registryCandidates, setScopedRegistry } from '../src/host/npm.mjs';

/** A DSH profile with @officesdk/officesdk-plugin-dsh installed from npm at `version` (DSH writes the exact version). */
function profile(version = '0.1.1-test.1', spec = version) {
  const base = mkdtempSync(join(tmpdir(), 'officesdk-npm-'));
  const profileDir = join(base, 'desktop');
  const packageRoot = join(profileDir, 'node_modules', '@officesdk', 'officesdk-plugin-dsh');
  mkdirSync(packageRoot, { recursive: true });
  writeFileSync(join(profileDir, 'package.json'), JSON.stringify({ dependencies: { '@officesdk/officesdk-plugin-dsh': spec } }));
  writeFileSync(join(packageRoot, 'package.json'), JSON.stringify({ name: '@officesdk/officesdk-plugin-dsh', version }));
  writeFileSync(join(profileDir, 'cordis.patch.yml'), 'a: 1\n');
  return { base, profileDir, packageRoot, dataRoot: join(base, 'data'), install: next => writeFileSync(join(packageRoot, 'package.json'), JSON.stringify({ name: '@officesdk/officesdk-plugin-dsh', version: next })) };
}
const packument = tags => ({ name: '@officesdk/officesdk-plugin-dsh', 'dist-tags': tags, versions: Object.fromEntries(Object.values(tags).map(v => [v, { version: v, dist: { integrity: `sha512-${v}` } }])) });
const registry = (tags, seen = []) => async url => { seen.push(url); return new Response(JSON.stringify(packument(tags)), { status: 200 }); };

test('an exact version, next or latest is an npm install; git, file and ranges are not', () => {
  assert.equal(npmInstall(profile('0.1.1-test.1').packageRoot).spec, '0.1.1-test.1');
  assert.equal(npmInstall(profile('0.1.1', 'next').packageRoot).kind, 'npm');
  for (const spec of ['git+ssh://git@example.invalid/x.git#dist', 'file:/tmp/x.tgz', '^0.1.0']) assert.equal(npmInstall(profile('0.1.1', spec).packageRoot), undefined, spec);
});

test('the registry: OFFICESDK_NPM_REGISTRY, else the profile .npmrc (scope line first), else npmjs then npmmirror', async () => {
  const p = profile();
  assert.deepEqual(registryCandidates(p.profileDir, {}).map(c => c.from), ['npmjs', 'npmmirror']);
  writeFileSync(join(p.profileDir, '.npmrc'), 'registry=https://r.example/\n//r.example/:_authToken=secret\n');
  assert.equal(npmrcRegistry(p.profileDir), 'https://r.example/');
  writeFileSync(join(p.profileDir, '.npmrc'), 'registry=https://r.example/\n@officesdk:registry=http://127.0.0.1:4873\n');
  assert.deepEqual(registryCandidates(p.profileDir, {}), [{ url: 'http://127.0.0.1:4873/', from: 'npmrc' }]);
  assert.deepEqual(registryCandidates(p.profileDir, { OFFICESDK_NPM_REGISTRY: 'https://env.example' }), [{ url: 'https://env.example/', from: 'env' }]);
});

test('npmjs unreachable falls back to npmmirror; a configured registry does not fall back', async () => {
  const p = profile();
  const seen = [];
  const fetchImpl = async url => { seen.push(url); if (url.startsWith(NPMJS)) throw Object.assign(new Error('fetch failed'), { cause: { code: 'ETIMEDOUT' } }); return new Response(JSON.stringify(packument({ latest: '0.1.1' })), { status: 200 }); };
  const answer = await fetchPackument({ profileDir: p.profileDir, env: {}, fetchImpl });
  assert.equal(answer.from, 'npmmirror');
  assert.deepEqual(seen, [`${NPMJS}@officesdk%2Fofficesdk-plugin-dsh`, `${NPMMIRROR}@officesdk%2Fofficesdk-plugin-dsh`]);
  await assert.rejects(fetchPackument({ profileDir: p.profileDir, env: { OFFICESDK_NPM_REGISTRY: 'https://down.example/' }, fetchImpl: async () => { throw new Error('down'); } }), /could not be read from npm/);
});

test('setScopedRegistry adds or replaces only the @officesdk line and keeps the rest', async () => {
  const p = profile();
  writeFileSync(join(p.profileDir, '.npmrc'), '//r.example/:_authToken=keep\n@other:registry=https://o/\n');
  assert.equal(await setScopedRegistry(p.profileDir, 'http://127.0.0.1:4873'), true);
  assert.equal(readFileSync(join(p.profileDir, '.npmrc'), 'utf8'), '//r.example/:_authToken=keep\n@other:registry=https://o/\n@officesdk:registry=http://127.0.0.1:4873/\n');
  assert.equal(await setScopedRegistry(p.profileDir, 'https://n/'), true);
  assert.match(readFileSync(join(p.profileDir, '.npmrc'), 'utf8'), /^@officesdk:registry=https:\/\/n\/$/m);
  assert.equal(readFileSync(join(p.profileDir, '.npmrc'), 'utf8').match(/@officesdk/g).length, 1);
  assert.equal(await setScopedRegistry(p.profileDir, 'https://n/'), false);
});

test('channels: next takes the newer of next and latest; latest only latest', () => {
  assert.equal(channelVersion(packument({ latest: '0.1.0', next: '0.1.1-test.2' }), 'next', newer), '0.1.1-test.2');
  assert.equal(channelVersion(packument({ latest: '0.1.1', next: '0.1.1-test.2' }), 'next', newer), '0.1.1');
  assert.equal(channelVersion(packument({ latest: '0.1.0', next: '0.1.1-test.2' }), 'latest', newer), '0.1.0');
  assert.equal(channelVersion(packument({ latest: '0.1.0' }), 'next', newer), '0.1.0');
});

test('the channel is kept per profile: a test user who received a release stays on next (R1-D3); OFFICESDK_CHANNEL changes it', async () => {
  const p = profile('0.1.1-test.1');
  const store = new ChannelStore(p.dataRoot);
  assert.equal(await store.resolve(p.profileDir, '0.1.1-test.1', {}), 'next');
  assert.equal(await store.resolve(p.profileDir, '0.1.1', {}), 'next', 'stays next after a release');
  assert.equal(await store.resolve(p.profileDir, '0.1.1', { OFFICESDK_CHANNEL: 'latest' }), 'latest');
  assert.equal(await store.resolve(p.profileDir, '0.1.2-test.1', {}), 'latest', 'written back');
  assert.equal(await new ChannelStore(p.dataRoot).resolve(profile('0.1.0').profileDir, '0.1.0', {}), 'latest');
});

test('an npm install checks its channel and updates with dsh plugin add <name>@<version> (the version the panel showed)', async () => {
  const p = profile('0.1.1-test.1');
  const calls = [];
  const run = async (file, args, options) => { calls.push({ file, args, cwd: options.cwd }); p.install('0.1.1-test.2'); return ''; };
  const updater = new PluginUpdater({ packageRoot: p.packageRoot, cli: '/dsh', run, dataRoot: p.dataRoot, env: {}, fetchImpl: registry({ latest: '0.1.0', next: '0.1.1-test.2' }) });
  const status = await updater.status();
  assert.deepEqual(status, { current: '0.1.1-test.1', latest: '0.1.1-test.2', available: true, channel: 'next' });
  assert.deepEqual(updater.command(), { file: '/dsh', args: ['plugin', '--profile', 'desktop', 'add', '@officesdk/officesdk-plugin-dsh@latest'] }, 'the pre-check needs no version');
  const result = await updater.update();
  assert.deepEqual(calls, [{ file: '/dsh', args: ['plugin', '--profile', 'desktop', 'add', '@officesdk/officesdk-plugin-dsh@0.1.1-test.2'], cwd: p.profileDir }]);
  assert.deepEqual(result, { current: '0.1.1-test.1', installed: '0.1.1-test.2', restartRequired: false, reloading: true });
  assert.match(readFileSync(join(p.profileDir, 'cordis.patch.yml'), 'utf8'), /^# officesdk: loaded 0\.1\.1-test\.2$/m);
  assert.equal(existsSync(join(p.profileDir, '.npmrc')), false, 'npmjs needs no .npmrc');
});

test('a registry from OFFICESDK_NPM_REGISTRY is written to the profile (scoped) before the add; npmmirror is not', async () => {
  for (const [env, fetchImpl, expected] of [
    [{ OFFICESDK_NPM_REGISTRY: 'http://127.0.0.1:4873/' }, registry({ latest: '0.1.1' }), '@officesdk:registry=http://127.0.0.1:4873/\n'],
    [{}, async url => { if (url.startsWith(NPMJS)) throw new Error('blocked'); return new Response(JSON.stringify(packument({ latest: '0.1.1' })), { status: 200 }); }, undefined],
  ]) {
    const p = profile('0.1.0');
    const updater = new PluginUpdater({ packageRoot: p.packageRoot, cli: '/dsh', run: async () => { p.install('0.1.1'); return ''; }, dataRoot: p.dataRoot, env, fetchImpl });
    await updater.status();
    await updater.update();
    let npmrc;
    try { npmrc = readFileSync(join(p.profileDir, '.npmrc'), 'utf8'); } catch { npmrc = undefined; }
    assert.equal(npmrc, expected);
  }
});

test('nothing newer, or no npm answer, offers nothing', async () => {
  const p = profile('0.1.1');
  const same = new PluginUpdater({ packageRoot: p.packageRoot, cli: '/dsh', run: async () => '', dataRoot: p.dataRoot, env: {}, fetchImpl: registry({ latest: '0.1.1' }) });
  assert.equal((await same.status()).available, false);
  const down = new PluginUpdater({ packageRoot: p.packageRoot, cli: '/dsh', run: async () => '', dataRoot: p.dataRoot, env: {}, fetchImpl: async () => { throw new Error('offline'); } });
  assert.deepEqual(await down.status(), { current: '0.1.1', latest: undefined, available: false, channel: 'latest' });
});

test('a downgrade is the same add with an older version', async () => {
  const p = profile('0.1.1');
  const calls = [];
  const updater = new PluginUpdater({ packageRoot: p.packageRoot, cli: '/dsh', run: async (file, args) => { calls.push(args.at(-1)); p.install('0.1.0'); return ''; }, dataRoot: p.dataRoot, env: {}, fetchImpl: registry({ latest: '0.1.1' }) });
  await assert.rejects(updater.update('0.1.0'), /not newer than the installed 0\.1\.1/, 'a click never downgrades (R2-F9)');
  await assert.rejects(updater.update('0.1.1'), /not newer/);
  assert.deepEqual(calls, []);
  const result = await updater.update('0.1.0', { downgrade: true });
  assert.deepEqual(calls, ['@officesdk/officesdk-plugin-dsh@0.1.0']);
  assert.equal(result.installed, '0.1.0');
  assert.equal(result.reloading, undefined, 'an older copy is not reloaded in place');
});

test('the update check carries the publish time of the version it offers, read from the full packument (V3.0 bubble date)', async () => {
  const p = profile('0.1.0');
  const accepts = [];
  const fetchImpl = async (url, init) => {
    accepts.push(init?.headers?.Accept);
    const data = packument({ latest: '0.1.1' });
    data.versions = {
      '0.1.0': { officeSdkReleaseNotes: { version: '0.1.0', highlights: ['旧版'], highlightsEn: ['Previous'] } },
      '0.1.1': { officeSdkReleaseNotes: { version: '0.1.1', highlights: ['新版'], highlightsEn: ['New'] } },
    };
    return new Response(JSON.stringify({ ...data, time: { created: '2026-11-01T00:00:00.000Z', '0.1.0': '2026-11-02T08:00:00.000Z', '0.1.1': '2026-11-12T03:00:00.000Z', bad: 'nope' } }), { status: 200 });
  };
  const updater = new PluginUpdater({ packageRoot: p.packageRoot, cli: '/dsh', run: async () => '', dataRoot: p.dataRoot, env: {}, fetchImpl });
  const status = await updater.status();
  assert.deepEqual([status.highlights, status.highlightsEn], [['新版'], ['New']]);
  assert.equal(status.releasedAt, '2026-11-12T03:00:00.000Z');
  assert.equal(status.latest, '0.1.1');
  assert.deepEqual(accepts, ['application/json'], 'the abbreviated install document has no publish times');
  // A registry without times: no date, the rest unchanged.
  const bare = new PluginUpdater({ packageRoot: p.packageRoot, cli: '/dsh', run: async () => '', dataRoot: p.dataRoot, env: {}, fetchImpl: registry({ latest: '0.1.1' }) });
  assert.equal(Object.hasOwn(await bare.status(), 'releasedAt'), false);
  // Other callers keep the abbreviated document.
  const seen = [];
  await fetchPackument({ profileDir: p.profileDir, env: {}, fetchImpl: async (url, init) => { seen.push(init.headers.Accept); return new Response(JSON.stringify(packument({ latest: '0.1.1' })), { status: 200 }); } });
  assert.deepEqual(seen, ['application/vnd.npm.install-v1+json, application/json']);
});
