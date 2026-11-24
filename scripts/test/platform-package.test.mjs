import test from 'node:test';
import assert from 'node:assert/strict';
import { chmodSync, existsSync, mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { GITATTRIBUTES, applyPlatform, checkPackage, kitExecutables, platformManifest, platformOf, tarballModeProblems, testVersion, versionAllowed } from '../platform-package.mjs';

const write = (file, text, mode) => { mkdirSync(path.dirname(file), { recursive: true }); writeFileSync(file, text); if (mode) chmodSync(file, mode); };
const SHARED = ['bos-pin.txt', 'modoc.js', 'mow_wasm.d.ts', 'mow_wasm.js', 'mow_wasm_bg.wasm'];
// The manifest fields of a released (0.0.35) package that the transform touches, in their order.
const MANIFEST = { name: '@officesdk/officesdk-plugin-dsh', version: '0.0.35', files: ['lib/host', 'resources/kit'], engines: {}, os: ['darwin'], cpu: ['arm64'], dsh: {} };

function base(t) {
  const dir = mkdtempSync(path.join(tmpdir(), 'platform-package-'));
  t.after(() => rmSync(dir, { recursive: true, force: true }));
  return dir;
}

/** An unpacked darwin-arm64 package like 0.0.35. */
function darwinPackage(dir) {
  const root = path.join(dir, 'package');
  write(path.join(root, 'package.json'), JSON.stringify(MANIFEST, null, 2) + '\n');
  for (const file of ['offline-kit-runtime', 'compose-worker']) write(path.join(root, 'resources', 'kit', file), file, 0o755);
  for (const file of SHARED) write(path.join(root, 'resources', 'kit', file), file);
  return root;
}

function windowsKit(dir) {
  const kit = path.join(dir, 'kit-win32');
  for (const file of ['offline-kit-runtime.exe', 'compose-worker.exe']) write(path.join(kit, file), file, 0o755);
  for (const file of SHARED) write(path.join(kit, file), file);
  return kit;
}

test('the platform table: darwin-arm64 is the default; win32-x64 is office-kit\'s MSVC build with .exe names', () => {
  assert.deepEqual(platformOf(), { name: 'darwin-arm64', os: 'darwin', cpu: 'arm64', triple: 'aarch64-apple-darwin', exe: '', gitattributes: false });
  assert.equal(platformOf('win32-x64').triple, 'x86_64-pc-windows-msvc');
  assert.deepEqual(kitExecutables('win32-x64'), ['offline-kit-runtime.exe', 'compose-worker.exe']);
  assert.throws(() => platformOf('linux-x64'), /unknown platform linux-x64/);
});

test('the manifest keeps its key order; .gitattributes is listed only where it ships', () => {
  const darwin = platformManifest(MANIFEST, 'darwin-arm64');
  assert.equal(JSON.stringify(darwin), JSON.stringify(MANIFEST), 'darwin-arm64 manifest is unchanged byte for byte');
  const win = platformManifest(MANIFEST, 'win32-x64');
  assert.deepEqual(Object.keys(win), Object.keys(MANIFEST));
  assert.deepEqual([win.os, win.cpu], [['win32'], ['x64']]);
  assert.deepEqual(win.files, [...MANIFEST.files, '.gitattributes']);
  assert.deepEqual(platformManifest(win, 'win32-x64').files, win.files, 'idempotent');
  assert.deepEqual(platformManifest(win, 'darwin-arm64').files, MANIFEST.files);
});

test('darwin-arm64: a released package passes unchanged and gains no file', async t => {
  const root = darwinPackage(base(t));
  const before = readFileSync(path.join(root, 'package.json'), 'utf8');
  await applyPlatform(root, 'darwin-arm64');
  assert.equal(readFileSync(path.join(root, 'package.json'), 'utf8'), before);
  assert.equal(existsSync(path.join(root, '.gitattributes')), false);
  assert.deepEqual(await checkPackage(root, 'darwin-arm64'), []);
});

test('win32-x64: the transform with a Windows Kit yields a complete package', async t => {
  const dir = base(t);
  const root = darwinPackage(dir);
  await applyPlatform(root, 'win32-x64', { kit: windowsKit(dir) });
  assert.deepEqual(await checkPackage(root, 'win32-x64'), []);
  assert.equal(readFileSync(path.join(root, '.gitattributes'), 'utf8'), GITATTRIBUTES);
  assert.equal(existsSync(path.join(root, 'resources', 'kit', 'offline-kit-runtime')), false, 'the darwin Kit is replaced, not merged');
  const manifest = JSON.parse(readFileSync(path.join(root, 'package.json'), 'utf8'));
  assert.deepEqual([manifest.os, manifest.cpu, manifest.files.at(-1)], [['win32'], ['x64'], '.gitattributes']);
  // The same tree is not a darwin package.
  const asDarwin = await checkPackage(root, 'darwin-arm64');
  for (const problem of ['package.json os is ["win32"], expected ["darwin"]', 'resources/kit/offline-kit-runtime is missing', 'resources/kit/offline-kit-runtime.exe belongs to another platform',
    '.gitattributes must not ship on this platform']) assert.ok(asDarwin.includes(problem), asDarwin.join('\n'));
});

test('win32-x64: a darwin Kit, a missing or wrong .gitattributes or an unlisted .gitattributes fail', async t => {
  const dir = base(t);
  // A darwin prebuilt scoped as Windows: the transform alone cannot make it one.
  const root = darwinPackage(dir);
  await applyPlatform(root, 'win32-x64');
  const problems = await checkPackage(root, 'win32-x64');
  for (const problem of ['resources/kit/offline-kit-runtime.exe is missing', 'resources/kit/compose-worker.exe is missing', 'resources/kit/offline-kit-runtime belongs to another platform'])
    assert.ok(problems.includes(problem), problems.join('\n'));

  const other = darwinPackage(path.join(dir, 'other'));
  await applyPlatform(other, 'win32-x64', { kit: windowsKit(dir) });
  rmSync(path.join(other, '.gitattributes'));
  write(path.join(other, 'resources', 'kit', 'stray'), 'x');
  chmodSync(path.join(other, 'resources', 'kit', 'compose-worker.exe'), 0o644);
  const manifest = JSON.parse(readFileSync(path.join(other, 'package.json'), 'utf8'));
  writeFileSync(path.join(other, 'package.json'), JSON.stringify({ ...manifest, files: manifest.files.filter(file => file !== '.gitattributes') }));
  assert.deepEqual((await checkPackage(other, 'win32-x64')).sort(), [
    '.gitattributes is missing (Git for Windows would convert line endings)',
    'package.json files does not list .gitattributes (npm pack would drop it)',
    'resources/kit/compose-worker.exe is not executable',
    'resources/kit/stray has no extension',
  ]);
  writeFileSync(path.join(other, '.gitattributes'), '* text=auto\n');
  assert.ok((await checkPackage(other, 'win32-x64')).includes('.gitattributes is "* text=auto\\n", expected "* -text\\n"'));
});

test('tarball modes are read from bsdtar and GNU tar listings', () => {
  const bsd = '-rwxr-xr-x  0 0      0     2048 Oct 26  1985 package/resources/kit/offline-kit-runtime\n-rw-r--r--  0 0      0       10 Oct 26  1985 package/resources/kit/compose-worker\n';
  assert.deepEqual(tarballModeProblems(bsd, 'darwin-arm64'), ['tarball mode of resources/kit/compose-worker is -rw-r--r--, expected -rwxr-xr-x']);
  const gnu = '-rwxr-xr-x 0/0    2048 1985-10-26 08:15 package/resources/kit/offline-kit-runtime.exe\n-rwxr-xr-x 0/0      10 1985-10-26 08:15 package/resources/kit/compose-worker.exe\n';
  assert.deepEqual(tarballModeProblems(gnu, 'win32-x64'), []);
  assert.deepEqual(tarballModeProblems(gnu, 'darwin-arm64'), ['tarball has no package/resources/kit/offline-kit-runtime', 'tarball has no package/resources/kit/compose-worker']);
});

test('versions: X.Y.Z and X.Y.Z-test.N everywhere; X.Y.Z-win32-test.N only for a Windows package', () => {
  for (const name of ['darwin-arm64', 'win32-x64']) {
    assert.ok(versionAllowed('0.0.47', name));
    assert.ok(versionAllowed('0.0.47-test.3', name));
    assert.ok(!versionAllowed('0.0.47-win32', name), 'the -win32 suffix belongs to the tag, not the package');
    assert.ok(!versionAllowed('v0.0.47', name));
  }
  assert.ok(versionAllowed('0.0.47-win32-test.12', 'win32-x64'));
  assert.ok(!versionAllowed('0.0.47-win32-test.12', 'darwin-arm64'));
  assert.equal(testVersion('0.0.47', 5, 'darwin-arm64'), '0.0.47-test.5');
  assert.equal(testVersion('0.0.47', 5, 'win32-x64'), '0.0.47-win32-test.5');
});
