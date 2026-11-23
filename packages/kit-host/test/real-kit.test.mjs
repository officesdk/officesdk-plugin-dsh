import test from 'node:test';
import assert from 'node:assert/strict';
import { copyFileSync, existsSync, mkdtempSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { KitProcess } from '../src/index.mjs';

const repo = fileURLToPath(new URL('../../../', import.meta.url));
// An unpacked office-kit Kit release for this machine (the plugin downloads the same archive on first use).
const kitDir = process.env.OFFICESDK_KIT_DIR;
const binary = kitDir ? path.join(kitDir, process.platform === 'win32' ? 'offline-kit-runtime.exe' : 'offline-kit-runtime') : '';
const skip = binary && existsSync(binary) ? false : 'no Kit: set OFFICESDK_KIT_DIR to an unpacked office-kit Kit release';

test('the staged Kit opens a PPTX, reads it back and stops without leaving a process', { skip, timeout: 120_000 }, async t => {
  const dir = mkdtempSync(path.join(tmpdir(), 'real-kit-'));
  t.after(() => rmSync(dir, { recursive: true, force: true }));
  const deck = path.join(dir, 'deck.pptx');
  copyFileSync(path.join(repo, 'test', 'fixtures', 'sample.pptx'), deck);
  const kit = new KitProcess({ binary, work: path.join(dir, 'work') });
  t.after(() => kit.stop());
  const client = await kit.start();
  const opened = await client.request('open', { path: deck });
  assert.deepEqual([opened.ok, opened.suite, opened.unsaved, opened.nextChangeSetRevision], [true, 'presentation', false, 0], JSON.stringify({ ...opened, content: undefined }));
  assert.ok(JSON.parse(opened.content), 'content is a MOP JSON string');
  const again = await client.request('content', { documentId: opened.documentId });
  assert.equal(again.content, opened.content);
  const unknownField = await client.request('content', { documentId: opened.documentId, workdir: '/tmp' });
  assert.deepEqual([unknownField.ok, unknownField.code], [false, 'forbidden-field']);
  const fonts = await client.request('font-catalog');
  assert.equal(fonts.ok, true);
  const { pid } = kit;
  await kit.stop();
  assert.equal(kit.client, null);
  assert.throws(() => process.kill(pid, 0), /ESRCH/, 'no Kit process is left behind');
});
