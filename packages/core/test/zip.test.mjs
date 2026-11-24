import test from 'node:test';
import assert from 'node:assert/strict';
import { deflateRawSync } from 'node:zlib';
import { checkOfficePackage, dropWriterRuntimeNodeIds, listZip, readZipEntry } from '../src/index.mjs';
import { zip } from './helpers.mjs';

test('office packages need content types and the format main part', () => {
  checkOfficePackage(zip({ '[Content_Types].xml': 'x', 'word/document.xml': 'y' }), 'word/document.xml');
  assert.throws(() => checkOfficePackage(zip({ '[Content_Types].xml': 'x' }), 'word/document.xml'), { code: 'invalid-output' });
  assert.throws(() => checkOfficePackage(Buffer.from('not a zip'), 'word/document.xml'), { code: 'invalid-format' });
});

test('entries are bounded, encrypted entries refused, and deflate supported', () => {
  const archive = listZip(zip({ 'a.txt': 'hello' }));
  assert.equal(readZipEntry(archive, archive.entries[0], 10).toString(), 'hello');
  assert.throws(() => readZipEntry(archive, archive.entries[0], 2), { code: 'input-too-large' });
  assert.throws(() => listZip(zip({ a: '1', b: '2' }), { maxEntries: 1 }), { code: 'input-too-large' });
  const encrypted = zip({ 'a.txt': 'x' });
  encrypted.writeUInt16LE(0x801, encrypted.length - 22 - 46 - 5 + 8);
  assert.throws(() => listZip(encrypted), { code: 'password-protected' });
  const bomb = listZip(zip({ 'a.txt': 'x' }));
  const deflated = deflateRawSync(Buffer.alloc(1024));
  const entry = { ...bomb.entries[0], method: 8, compressedSize: deflated.length, uncompressedSize: 0 };
  const buffer = Buffer.concat([bomb.buffer.subarray(0, 30 + 5), deflated]);
  assert.throws(() => readZipEntry({ buffer }, { ...entry, local: 0 }, 100), { code: 'invalid-format' });
});

test('Writer runtime nodeIds are removed at every depth', () => {
  const document = { nodeId: 1, blocks: [{ table: { rows: [{ nodeId: 'x', text: 'a' }] } }] };
  assert.equal(dropWriterRuntimeNodeIds(document), true);
  assert.deepEqual(document, { blocks: [{ table: { rows: [{ text: 'a' }] } }] });
  assert.equal(dropWriterRuntimeNodeIds({ a: [1, 2] }), false);
});
