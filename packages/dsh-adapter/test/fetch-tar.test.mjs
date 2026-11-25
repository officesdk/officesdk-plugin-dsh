import test from 'node:test';
import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import { existsSync, mkdtempSync, readFileSync, statSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { Readable } from 'node:stream';
import { TarError, extractTar, hashTar, safeName } from '../src/host/fetch/tar.mjs';
import { makeTar } from './fixtures/tar-writer.mjs';

const sri = (algorithm, data) => `${algorithm}-${createHash(algorithm).update(data).digest('base64')}`;
const stream = (bytes, chunk = 333) => Readable.from((function* () { for (let i = 0; i < bytes.length; i += chunk) yield bytes.subarray(i, i + chunk); })());
const dir = () => mkdtempSync(join(tmpdir(), 'tar-test-'));

test('extracts files and directories from a gzip stream in small chunks, hashing each file and keeping the executable bit', async () => {
  const big = Buffer.alloc(5000, 7);
  const tar = makeTar([{ name: './bin/', type: '5' }, { name: './bin/offline-kit-runtime', data: big, mode: 0o755 }, { name: 'modoc.js', data: 'x = 1' }], { gzip: true });
  const dest = dir();
  const { files } = await extractTar(stream(tar), dest, { gzip: true });
  assert.deepEqual(Object.keys(files).sort(), ['bin/offline-kit-runtime', 'modoc.js']);
  assert.deepEqual(files['bin/offline-kit-runtime'], { size: 5000, sha256: sri('sha256', big), sha512: sri('sha512', big), executable: true });
  assert.equal(readFileSync(join(dest, 'modoc.js'), 'utf8'), 'x = 1');
  if (process.platform !== 'win32') {
    assert.ok(statSync(join(dest, 'bin/offline-kit-runtime')).mode & 0o100);
    assert.equal(statSync(join(dest, 'modoc.js')).mode & 0o111, 0);
  }
});

test('pax paths, GNU long names and the ustar prefix carry long and UTF-8 names', async () => {
  const long = `${'a'.repeat(120)}/文件.txt`;
  const tar = makeTar([
    { name: 'short', pax: { path: 'sdk/中文名.png' }, data: 'p' },
    { name: 'ignored', longName: `gnu/${'b'.repeat(130)}.js`, data: 'g' },
    { name: long, data: 'u' },
  ]);
  const { files } = await hashTar(stream(tar));
  assert.deepEqual(Object.keys(files).sort(), [long, 'sdk/中文名.png', `gnu/${'b'.repeat(130)}.js`].sort());
});

test('strip keeps one subtree: other entries are rejected, or skipped when asked', async () => {
  const tar = makeTar([{ name: 'package/dist/resources/catalog.json', data: '{}' }, { name: 'package/package.json', data: '{}' }]);
  await assert.rejects(hashTar(stream(tar), { strip: 'package/dist/resources' }), /outside package\/dist\/resources/);
  const { files } = await hashTar(stream(tar), { strip: 'package/dist/resources', outside: 'skip' });
  assert.deepEqual(Object.keys(files), ['catalog.json']);
  // git archive output: the skill directory's prefix comes off.
  const git = makeTar([{ name: 'packages/dsh-adapter/skills/shimo-pptx/', type: '5' }, { name: 'packages/dsh-adapter/skills/shimo-pptx/styles/default/style.json', data: '{}' }]);
  assert.deepEqual(Object.keys((await hashTar(stream(git), { strip: 'packages/dsh-adapter/skills/shimo-pptx/' })).files), ['styles/default/style.json']);
});

test('links, absolute names, .. and duplicate files are refused before anything escapes the target', async () => {
  for (const [entries, pattern] of [
    [[{ name: 'evil', type: '2', linkname: '/etc/passwd' }], /not a regular file/],
    [[{ name: 'hard', type: '1', linkname: 'x' }], /not a regular file/],
    [[{ name: '/abs.txt', data: 'x' }], /absolute path/],
    [[{ name: 'C:/win.txt', data: 'x' }], /absolute path/],
    [[{ name: 'a/../../up.txt', data: 'x' }], /\.\./],
    [[{ name: 'twice', data: '1' }, { name: './twice', data: '2' }], /appears twice/],
  ]) {
    const dest = dir();
    await assert.rejects(extractTar(stream(makeTar(entries)), join(dest, 'out')), error => error instanceof TarError && pattern.test(error.message));
    assert.equal(existsSync(join(dest, 'up.txt')), false);
  }
  assert.equal(safeName('./a//b/./c'), 'a/b/c');
  assert.throws(() => safeName('a\\..\\b'), /\.\./);
});

test('a truncated archive, a corrupted header and a non-tar fail as bad archives', async () => {
  const tar = makeTar([{ name: 'f', data: Buffer.alloc(2000, 1) }]);
  await assert.rejects(hashTar(stream(tar.subarray(0, 1200))), /ended/);
  const corrupt = Buffer.from(tar); corrupt[10] ^= 0xff;
  await assert.rejects(hashTar(stream(corrupt)), /checksum/);
  await assert.rejects(hashTar(stream(makeTar([{ name: 'f', data: 'x' }], { end: false }))), /end blocks/);
  await assert.rejects(hashTar(stream(Buffer.from('not a gzip at all'.repeat(40))), { gzip: true }), error => error.code === 'Z_DATA_ERROR' || error instanceof TarError);
});

test('macOS AppleDouble companions (._name) are skipped: office-kit\'s Mac archive has them and npm drops them', async () => {
  const { files } = await hashTar(stream(makeTar([{ name: './._offline-kit-runtime', data: 'xattr' }, { name: './offline-kit-runtime', data: 'bin', mode: 0o755 }, { name: 'guide/._README.md', data: 'x' }])));
  assert.deepEqual(Object.keys(files), ['offline-kit-runtime']);
});
