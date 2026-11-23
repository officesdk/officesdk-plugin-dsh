#!/usr/bin/env node
// A scripted stand-in for `offline-kit-runtime serve --work DIR`: protocol "5" frames on stdin/stdout and the
// side stream on fd 3. FAKE_KIT_PROTOCOL changes the ready frame; ops script the edge cases under test.
// FAKE_KIT_SPLIT=1 is Windows' layout: Kit reads the Host's media on fd 3 and writes its own on fd 4 only.
import { appendFileSync, writeFileSync } from 'node:fs';
import net from 'node:net';

const work = process.argv[process.argv.indexOf('--work') + 1];
// fs streams on a socket fd park reads in the thread pool and keep process.exit from returning.
const socket = new net.Socket({ fd: 3, readable: true, writable: true });
const side = process.env.FAKE_KIT_SPLIT ? { in: socket, out: new net.Socket({ fd: 4, readable: true, writable: true }) } : { out: socket, in: socket };
const send = frame => { const body = Buffer.from(JSON.stringify(frame)); const head = Buffer.alloc(4); head.writeUInt32BE(body.length); process.stdout.write(Buffer.concat([head, body])); };
const sideFrame = (stream, bytes) => { const head = Buffer.alloc(16); head.writeBigUInt64BE(BigInt(stream)); head.writeBigUInt64BE(BigInt(bytes.length), 8); side.out.write(Buffer.concat([head, bytes])); };
let outstanding = 0, maxOutstanding = 0, received = Buffer.alloc(0), sideReceived = Buffer.alloc(0);
side.in.on('data', chunk => { sideReceived = Buffer.concat([sideReceived, chunk]); });
appendFileSync(`${work}/fake-kit.log`, `start ${process.pid}\n`);

send({ ready: true, protocolVersion: process.env.FAKE_KIT_PROTOCOL ?? '5' });
process.stdin.on('data', chunk => {
  received = Buffer.concat([received, chunk]);
  while (received.length >= 4 && received.length >= 4 + received.readUInt32BE(0)) {
    const length = received.readUInt32BE(0);
    const command = JSON.parse(received.subarray(4, 4 + length));
    received = received.subarray(4 + length);
    handle(command);
  }
});
process.stdin.on('end', () => { appendFileSync(`${work}/fake-kit-stdin.log`, `${Date.now()}\n`); if (!process.env.FAKE_KIT_IGNORE_STDIN_END) process.exit(0); });
process.on('SIGTERM', () => { if (process.env.FAKE_KIT_IGNORE_TERM) return; process.exit(0); });

const unsaved = new Set(), changes = new Map();
function handle({ id, op, ...fields }) {
  appendFileSync(`${work}/commands.log`, `${JSON.stringify({ op, ...fields })}\n`);
  switch (op) {
    case 'open':
      if (fields.path.includes('locked')) return send({ id, ok: false, code: 'need-password', error: 'password required' });
      return send({ id, ok: true, documentId: `doc-${fields.path}`, suite: fields.path.includes('mismatch') ? (fields.path.endsWith('.xlsx') ? 'document' : 'sheet') : fields.path.endsWith('.docx') ? 'document' : fields.path.endsWith('.xlsx') ? 'sheet' : fields.path.endsWith('.pptx') ? 'presentation' : 'unknown', content: '{"slides":[]}', unsaved: unsaved.has(`doc-${fields.path}`), nextChangeSetRevision: 0, ...(fields.path.includes('disk') && { diskChanged: true }) });
    case 'change':
      if (fields.body === 'reject') return send({ id, ok: false, code: 'change-body-rejected', error: 'rejected' });
      unsaved.add(fields.documentId);
      changes.set(fields.documentId, (changes.get(fields.documentId) ?? 0) + 1);
      return setTimeout(() => send({ id, ok: true }), 20);
    case 'content': return send({ id, ok: true, documentId: fields.documentId, suite: 'presentation', content: `{"slides":[],"changes":${changes.get(fields.documentId) ?? 0}}`, unsaved: unsaved.has(fields.documentId), nextChangeSetRevision: changes.get(fields.documentId) ?? 0 });
    case 'save':
      if (fields.target.includes('protected')) return send({ id, ok: false, code: 'protection-unsupported', error: 'per-user rules' });
      if (fields.target.includes('changed') && !fields.overwrite) return send({ id, ok: false, code: 'file-changed', error: 'changed on disk' });
      unsaved.delete(fields.documentId);
      // A slow save lets the test land a change while it runs (Kit runs frames in order).
      return setTimeout(() => send({ id, ok: true }), fields.target.includes('slow') ? 150 : 0);
    case 'export': writeFileSync(fields.target, 'pptx bytes'); return send({ id, ok: true });
    case 'import': unsaved.delete(fields.documentId); return send({ id, ok: true, documentId: fields.documentId, suite: 'presentation', content: '{"slides":[]}', unsaved: false, nextChangeSetRevision: 0 });
    case 'open-media-read': {
      if (fields.range === 'bytes=999-') return send({ id, ok: false, code: 'media-range-unsatisfiable', error: 'range' });
      if (fields.range) { const bytes = Buffer.from('dia'); sideFrame(13, bytes); return send({ id, ok: true, stream: '13', length: 3, contentType: 'video/mp4', offset: 2, totalLength: 10 }); }
      const bytes = Buffer.from(`media ${fields.reference}`); sideFrame(11, bytes); return send({ id, ok: true, stream: '11', length: bytes.length, contentType: 'image/png' }); }
    case 'begin-media-write': return send({ id, ok: true, stream: '12' });
    case 'finish-media-write': return setTimeout(() => send({ id, ok: true, reference: `mop-asset:/${fields.kind ?? 'media'}/${'a'.repeat(64)}.png` }), 20);
    case 'fail': return send({ id, ok: false, code: 'file-changed', error: 'The file changed on disk.' });
    case 'hold':
      outstanding++; maxOutstanding = Math.max(maxOutstanding, outstanding);
      return setTimeout(() => { outstanding--; send({ id, ok: true }); }, 60);
    case 'stats': return send({ id, ok: true, maxOutstanding });
    case 'media-read': {
      const bytes = Buffer.from(fields.text ?? 'picture bytes');
      sideFrame(7, bytes); // bytes first, the JSON that names the stream afterwards
      return setTimeout(() => send({ id, ok: true, stream: '7', length: bytes.length, contentType: 'image/png' }), 30);
    }
    case 'media-write': return send({ id, ok: true, stream: '9' });
    case 'media-written': return setTimeout(() => send({ id, ok: true, received: sideReceived.subarray(16).toString() }), 30);
    case 'no-id': return send({ ok: true });
    case 'garbage': { const body = Buffer.from('not json'); const head = Buffer.alloc(4); head.writeUInt32BE(body.length); return process.stdout.write(Buffer.concat([head, body])); }
    case 'crash': return process.exit(3);
    case 'never': return undefined;
    default: return send({ id, ok: false, code: 'unknown-operation', error: `unknown op ${op}` });
  }
}
