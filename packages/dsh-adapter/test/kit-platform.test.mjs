import test from 'node:test';
import assert from 'node:assert/strict';
import { join } from 'node:path';
import { KIT_TRIPLES, PLATFORMS } from '../../../scripts/platform-package.mjs';
import { KIT_PLATFORMS, NO_CONVERTER, platformKey, resolveKit, unsupportedError } from '../src/host/kit-platform.mjs';

const ROOT = '/pkg/resources/kit';
const files = names => path => names.includes(path.slice(ROOT.length + 1));
const MAC = ['offline-kit-runtime', 'compose-worker', 'modoc.js', 'mow_wasm.js', 'bos-pin.txt'];
const WINDOWS = ['offline-kit-runtime.exe', 'compose-worker.exe', 'modoc.js', 'mow_wasm.js', 'bos-pin.txt'];
const darwin = { platform: 'darwin', arch: 'arm64' }, win32 = { platform: 'win32', arch: 'x64' };

test('the platforms a package can carry: darwin-arm64 as today, win32-x64 with .exe names', () => {
  assert.deepEqual(Object.keys(KIT_PLATFORMS), ['darwin-arm64', 'win32-x64']);
  assert.equal(platformKey(darwin), 'darwin-arm64');
  assert.equal(platformKey(), `${process.platform}-${process.arch}`);
});

test('the Host\'s platform table matches the build\'s: same platforms, same executable suffixes', () => {
  assert.deepEqual(Object.keys(KIT_PLATFORMS).sort(), Object.keys(PLATFORMS).sort());
  for (const [name, { exe }] of Object.entries(KIT_PLATFORMS)) assert.equal(exe, PLATFORMS[name].exe, name);
  // Every packaged platform is an office-kit release target with the same triple and suffix.
  for (const [name, { triple, exe }] of Object.entries(PLATFORMS)) assert.deepEqual([KIT_TRIPLES[triple]?.platform, KIT_TRIPLES[triple]?.exe, KIT_TRIPLES[triple]?.skip], [name, exe, undefined], name);
  for (const { platform, exe } of Object.values(KIT_TRIPLES)) assert.equal(exe, platform.startsWith('win32-') ? '.exe' : '', platform);
});

test('darwin-arm64 launches the extension-less Kit, win32-x64 the .exe', () => {
  assert.deepEqual(resolveKit(ROOT, darwin, files(MAC)), { binary: join(ROOT, 'offline-kit-runtime') });
  assert.deepEqual(resolveKit(ROOT, win32, files(WINDOWS)), { binary: join(ROOT, 'offline-kit-runtime.exe') });
});

test('a missing, foreign or incomplete Kit is never launched and says why', () => {
  assert.deepEqual(resolveKit(ROOT, win32, files([])), { unsupported: { platform: 'win32-x64', reason: 'missing', file: 'offline-kit-runtime.exe' } });
  assert.deepEqual(resolveKit(ROOT, darwin, files([])), { unsupported: { platform: 'darwin-arm64', reason: 'missing', file: 'offline-kit-runtime' } });
  // The macOS package on Windows, and the Windows package on a Mac.
  assert.equal(resolveKit(ROOT, win32, files(MAC)).unsupported.reason, 'mismatch');
  assert.equal(resolveKit(ROOT, darwin, files(WINDOWS)).unsupported.reason, 'mismatch');
  // Kit checks its layout from its own directory: the worker and modoc.js must sit beside it.
  assert.deepEqual(resolveKit(ROOT, win32, files(['offline-kit-runtime.exe', 'compose-worker', 'modoc.js'])).unsupported, { platform: 'win32-x64', reason: 'incomplete', file: 'compose-worker.exe' });
  assert.deepEqual(resolveKit(ROOT, darwin, files(['offline-kit-runtime', 'compose-worker'])).unsupported, { platform: 'darwin-arm64', reason: 'incomplete', file: 'modoc.js' });
  // No Kit is built for an Intel Mac or Windows on Arm, whatever the folder holds.
  assert.deepEqual(resolveKit(ROOT, { platform: 'darwin', arch: 'x64' }, files(MAC)), { unsupported: { platform: 'darwin-x64', reason: 'platform' } });
  assert.deepEqual(resolveKit(ROOT, { platform: 'win32', arch: 'arm64' }, files(WINDOWS)), { unsupported: { platform: 'win32-arm64', reason: 'platform' } });
});

test('the user-facing answer starts with 本平台暂不支持 and names what to install', () => {
  const cases = [
    [{ platform: 'win32-x64', reason: 'missing', file: 'offline-kit-runtime.exe' }, /没有 win32-x64 的文档服务（offline-kit-runtime\.exe）.*安装 win32-x64 版本/],
    [{ platform: 'win32-x64', reason: 'mismatch' }, /其他平台.*安装 win32-x64 版本/],
    [{ platform: 'darwin-arm64', reason: 'incomplete', file: 'modoc.js' }, /缺少 modoc\.js/],
    [{ platform: 'darwin-x64', reason: 'platform' }, /只支持 darwin-arm64、win32-x64，当前系统是 darwin-x64/],
  ];
  for (const [unsupported, text] of cases) {
    const error = unsupportedError(unsupported);
    assert.equal(error.name, 'SuiteError');
    assert.equal(error.code, 'unsupported-operation', 'a fixed failure: the client shows no retry');
    assert.match(error.message, /^本平台暂不支持打开 Office 文档：/);
    assert.match(error.message, text);
    assert.match(error.message, /This platform cannot open Office documents yet: /, 'English follows for an English UI');
    assert.deepEqual(error.details, { platform: unsupported.platform, reason: unsupported.reason });
  }
});

test('without Core\'s converter, Core claims no operation and converts nothing', async () => {
  assert.deepEqual(NO_CONVERTER.capabilities(), {});
  await assert.rejects(NO_CONVERTER.import({}), error => error.code === 'unavailable');
  await assert.rejects(NO_CONVERTER.export({}), error => error.code === 'unavailable');
});
