import test from 'node:test';
import assert from 'node:assert/strict';
import { mkdirSync, mkdtempSync, readFileSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { PluginUpdater, cliCommand, findCli, redactCredentials, updateWithKit, withProfileLock } from '../src/host/updater.mjs';

/** A DSH profile with this package installed from npm by `spec` (an exact version unless given), at `version`. */
function profile(spec, version = '0.0.21') {
  const base = mkdtempSync(join(tmpdir(), 'officesdk-updater-'));
  const profileDir = join(base, 'desktop');
  const packageRoot = join(profileDir, 'node_modules', '@officesdk', 'officesdk-plugin-dsh');
  mkdirSync(packageRoot, { recursive: true });
  writeFileSync(join(profileDir, 'package.json'), JSON.stringify({ dependencies: { '@officesdk/officesdk-plugin-dsh': spec ?? version } }));
  writeFileSync(join(packageRoot, 'package.json'), JSON.stringify({ name: '@officesdk/officesdk-plugin-dsh', version }));
  return { profileDir, packageRoot, dataRoot: join(base, 'data'), install: next => writeFileSync(join(packageRoot, 'package.json'), JSON.stringify({ name: '@officesdk/officesdk-plugin-dsh', version: next })) };
}
const registry = latest => async () => new Response(JSON.stringify({ name: '@officesdk/officesdk-plugin-dsh', 'dist-tags': { latest } }), { status: 200 });
const npmUpdater = (p, options) => new PluginUpdater({ packageRoot: p.packageRoot, dataRoot: p.dataRoot, env: {}, fetchImpl: registry('0.0.22'), ...options });

test('only an npm install by exact version or channel updates itself', async () => {
  for (const spec of ['git+ssh://git@example.invalid/x.git#dist', 'file:/tmp/officesdk.tgz', '^0.0.21']) {
    const updater = npmUpdater(profile(spec), { run: async () => '' });
    assert.deepEqual(await updater.status(), { current: '0.0.21', available: false }, spec);
    await assert.rejects(updater.update(), { code: 'unsupported-operation' });
  }
});

test('after an update a marker line in the profile patch makes DSH recompose; the line is replaced, never repeated', async () => {
  const p = profile();
  const patch = join(p.profileDir, 'cordis.patch.yml');
  writeFileSync(patch, '- id: locale\n  config:\n    preference: zh');
  let next = '0.0.22';
  const updater = npmUpdater(p, { cli: '/dsh', run: async file => { if (file === '/dsh') p.install(next); return ''; } });
  assert.deepEqual(await updater.update('0.0.22'), { current: '0.0.21', installed: '0.0.22', restartRequired: false, reloading: true });
  assert.equal(readFileSync(patch, 'utf8'), '- id: locale\n  config:\n    preference: zh\n# officesdk: loaded 0.0.22\n');
  next = '0.0.23';
  await updater.update('0.0.23');
  assert.equal(readFileSync(patch, 'utf8'), '- id: locale\n  config:\n    preference: zh\n# officesdk: loaded 0.0.23\n');
});

test('credentials never reach the client in update errors', () => {
  assert.equal(redactCredentials('npm error 401 https://user:npm-SECRET@registry.example.invalid/@officesdk%2fofficesdk-plugin-dsh'),
    'npm error 401 https://registry.example.invalid/@officesdk%2fofficesdk-plugin-dsh');
  assert.equal(redactCredentials('git+ssh://git@example.invalid/x.git#dist'), 'git+ssh://example.invalid/x.git#dist');
});

const WIN = { platform: 'win32', resourcesPath: 'C:\\Program Files (x86)\\DeepSeek Harness\\resources', execPath: 'C:\\Program Files (x86)\\DeepSeek Harness\\DeepSeek Harness.exe', env: {} };
const LAUNCHER = 'C:\\Program Files (x86)\\DeepSeek Harness\\resources\\runtime\\cli\\bin\\dsh.cmd';
const ENTRY = 'C:\\Program Files (x86)\\DeepSeek Harness\\resources\\app.asar\\dsh\\node_modules\\@deepseek-ai\\dsh-desktop-host\\lib\\cli.js';
const UPDATE = ['plugin', '--profile', 'desktop', 'add', '@officesdk/officesdk-plugin-dsh@0.0.22'];

test('the DSH command line: macOS keeps bin/dsh; Windows takes bin/dsh.cmd, with the app running its CLI entry as the fallback', () => {
  assert.deepEqual(findCli({ platform: 'darwin', resourcesPath: '/App/Resources', env: {}, exists: path => path === '/App/Resources/runtime/cli/bin/dsh' }), { file: '/App/Resources/runtime/cli/bin/dsh', args: [] });
  assert.deepEqual(findCli({ platform: 'darwin', env: {}, exists: () => false }), { file: 'dsh', args: [] });
  assert.deepEqual(findCli({ platform: 'darwin', env: { DSH_CLI: '/gone/dsh' }, resourcesPath: '/App/Resources', exists: path => path === '/App/Resources/runtime/cli/bin/dsh' }).file, '/App/Resources/runtime/cli/bin/dsh', 'a DSH_CLI that is not there is skipped, as before');
  const direct = { file: WIN.execPath, args: ['--expose-internals', ENTRY], env: { ELECTRON_RUN_AS_NODE: '1' } };
  assert.deepEqual(findCli({ ...WIN, exists: path => path === LAUNCHER }), { file: LAUNCHER, args: [], cmd: true, fallback: direct });
  assert.deepEqual(findCli({ ...WIN, exists: () => false }), direct, 'no launcher: the CLI entry directly');
  assert.deepEqual(findCli({ ...WIN, env: { DSH_CLI: 'D:\\dsh\\dsh.cmd' }, exists: () => false }), { file: 'D:\\dsh\\dsh.cmd', args: [], cmd: true, fallback: direct });
  assert.deepEqual(findCli({ platform: 'win32', env: {}, exists: () => false }), { file: 'dsh.cmd', args: [], cmd: true });
});

test('a .cmd launcher runs through cmd.exe only when cmd.exe cannot change what runs; otherwise the direct entry does', () => {
  const cli = findCli({ ...WIN, exists: path => path === LAUNCHER });
  assert.deepEqual(cliCommand(cli, UPDATE, { comspec: 'C:\\Windows\\system32\\cmd.exe' }), { file: 'C:\\Windows\\system32\\cmd.exe',
    args: ['/d', '/s', '/c', `""${LAUNCHER}" plugin --profile desktop add @officesdk/officesdk-plugin-dsh@0.0.22"`], windowsVerbatimArguments: true });
  // A profile name cmd.exe would split or expand: the arguments go to the app's executable as an argument list instead.
  for (const profile of ['my profile', 'a&calc', '%PATH%', 'x"y', 'a^b', 'p(1)', 'a,b', 'a=b', 'a;b']) {
    assert.deepEqual(cliCommand(cli, ['plugin', '--profile', profile, 'add', '@officesdk/officesdk-plugin-dsh@0.0.22']),
      { file: WIN.execPath, args: ['--expose-internals', ENTRY, 'plugin', '--profile', profile, 'add', '@officesdk/officesdk-plugin-dsh@0.0.22'], env: { ELECTRON_RUN_AS_NODE: '1' } }, profile);
  }
  assert.equal(cliCommand({ ...cli, file: 'C:\\100%\\dsh.cmd' }, UPDATE).file, WIN.execPath, 'a launcher path cmd.exe would expand');
  assert.throws(() => cliCommand({ file: 'dsh.cmd', args: [], cmd: true }, ['plugin', '--profile', 'a&b']), { code: 'unsupported-operation' });
  // macOS and a bare path string run as they always did.
  assert.deepEqual(cliCommand('/dsh', UPDATE), { file: '/dsh', args: UPDATE });
});

test('on Windows the update goes through cmd.exe with its arguments verbatim and no console window', async () => {
  const p = profile();
  const calls = [];
  const updater = npmUpdater(p, { cli: findCli({ ...WIN, exists: path => path === LAUNCHER }), run: async (file, args, options) => { calls.push([file, args, options]); return ''; } });
  await updater.update('0.0.22');
  const [[file, args, options]] = calls;
  assert.match(file, /cmd(\.exe)?$/i);
  assert.equal(args.at(-1), `""${LAUNCHER}" plugin --profile desktop add @officesdk/officesdk-plugin-dsh@0.0.22"`);
  assert.deepEqual([options.windowsVerbatimArguments, options.cwd], [true, p.profileDir]);
});

test('an update waits for every Kit view to be saved; on Windows Kit stops for it and starts again if the update fails', async () => {
  const kit = () => { const events = []; return { events, suspend: async () => { events.push('suspend'); }, resume: () => { events.push('resume'); } }; };
  const updater = (result, events) => ({ command: () => {}, update: async () => { events?.push('update'); if (result instanceof Error) throw result; return result; } });
  // Unsaved edits or an AI writing, in any window: refused before anything stops.
  for (const view of [{ unsaved: true }, { aiActive: 1 }]) {
    const k = kit();
    await assert.rejects(updateWithKit({ updater: updater({}, k.events), views: [{}, view], kit: k, platform: 'win32' }), error => error.code === 'unsaved-changes' && /1 open document has unsaved changes/.test(error.message));
    assert.deepEqual(k.events, []);
  }
  // Windows, reloaded by DSH: Kit stopped before the update and left stopped for the new instance.
  let k = kit();
  assert.deepEqual(await updateWithKit({ updater: updater({ installed: '0.0.40', reloading: true }, k.events), views: [{}], kit: k, platform: 'win32' }), { installed: '0.0.40', reloading: true });
  assert.deepEqual(k.events, ['suspend', 'update']);
  // Windows, waiting for a restart: the old instance keeps working, so Kit may start again.
  k = kit();
  await updateWithKit({ updater: updater({ restartRequired: true }, k.events), views: [], kit: k, platform: 'win32' });
  assert.deepEqual(k.events, ['suspend', 'update', 'resume']);
  // Windows, failed: Kit may start again and the failure stays the answer.
  k = kit();
  await assert.rejects(updateWithKit({ updater: updater(new Error('pnpm failed'), k.events), views: [], kit: k, platform: 'win32' }), /pnpm failed/);
  assert.deepEqual(k.events, ['suspend', 'update', 'resume']);
  // A copy that cannot update at all refuses before Kit stops.
  k = kit();
  const fixed = npmUpdater(profile('file:/tmp/officesdk.tgz'), { run: async () => '' });
  await assert.rejects(updateWithKit({ updater: fixed, views: [], kit: k, platform: 'win32' }), { code: 'unsupported-operation' });
  assert.deepEqual(k.events, []);
  // macOS replaces files under a running Kit as before: nothing stops.
  k = kit();
  await updateWithKit({ updater: updater({ reloading: true }, k.events), views: [{}], kit: k, platform: 'darwin' });
  assert.deepEqual(k.events, ['update']);
});

test('the update installs the version the panel showed; profile changes run one at a time (R1-F3, R1-F6)', async () => {
  const targets = [];
  await updateWithKit({ updater: { command() {}, update: async target => { targets.push(target); return {}; } }, target: '0.1.2' });
  assert.deepEqual(targets, ['0.1.2']);
  const order = [];
  const task = (name, ms) => () => new Promise(done => { order.push(`${name} start`); setTimeout(() => { order.push(`${name} end`); done(); }, ms); });
  await Promise.all([withProfileLock(task('a', 30)), withProfileLock(task('b', 1)), withProfileLock(() => { throw new Error('c fails'); }).catch(() => order.push('c failed')), withProfileLock(task('d', 1))]);
  assert.deepEqual(order, ['a start', 'a end', 'b start', 'b end', 'c failed', 'd start', 'd end']);
});
