import test from 'node:test';
import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import { sriHex, sumsProblems } from '../kit-smoke.mjs';
import { parseSums } from '../vendor.mjs';

const sri = text => `sha256-${createHash('sha256').update(text).digest('base64')}`;
const hex = text => createHash('sha256').update(text).digest('hex');

test('an SRI sha256 reads as the hex SHA256SUMS uses', () => {
  assert.equal(sriHex(sri('a')), hex('a'));
  assert.equal(sriHex('sha512-x'), undefined);
});

test('every archive of the manifest must be listed in SHA256SUMS with the same sha256', () => {
  const manifest = { kit: { 'darwin-arm64': { file: 'k.tar.gz', sha256: sri('k') }, 'win32-x64': { file: 'w.tar.gz', sha256: sri('w') } }, editor: { file: 'e.tgz', sha256: sri('e') } };
  const sums = parseSums(`${hex('k')}  k.tar.gz\r\n${hex('w')} *w.tar.gz\n${hex('e')}  e.tgz\n\nnot a line\n`);
  assert.deepEqual(sumsProblems(manifest, sums), []);
  sums.set('w.tar.gz', hex('other'));
  sums.delete('e.tgz');
  assert.deepEqual(sumsProblems(manifest, sums), [`SHA256SUMS lists w.tar.gz as ${hex('other')}, kit-manifest.json says ${hex('w')} (kit win32-x64)`, 'SHA256SUMS does not list e.tgz (editor)']);
});
