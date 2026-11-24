import { PassThrough } from 'node:stream';

/**
 * Media and font bytes travel on Kit's side stream (kit-usage/02-frames.md): a 16-byte header, then the body.
 * Unix has one duplex pipe at stdio index 3; Windows has two one-way pipes (kit-usage/01-program.md).
 */
export const MAX_SIDE_BODY = 268_435_456;
const STREAM = /^[0-9]{1,20}$/;

/**
 * Demultiplexes Kit's side stream. Bytes for a stream may arrive before the JSON response that names it:
 * they wait until `openBody` claims the stream with the length from that response. `input` carries Kit's
 * bytes to the Host; `output` (the same pipe on Unix) carries the Host's media writes to Kit.
 */
export function createSideStream(input, output = input) {
  let buffer = Buffer.alloc(0), failed = null, writing = Promise.resolve();
  let parse = { mode: 'header' };
  const slots = new Map(), early = new Map();
  const fail = error => {
    if (failed) return;
    failed = error;
    for (const slot of slots.values()) slot.output?.destroy(error);
    slots.clear(); early.clear();
  };
  const keepEarly = (id, chunk, done) => {
    const queued = early.get(id) ?? { chunks: [], done: false };
    if (chunk.length) queued.chunks.push(chunk);
    queued.done ||= done;
    early.set(id, queued);
  };
  const deliver = (id, chunk, done) => {
    const slot = slots.get(id);
    if (!slot) return keepEarly(id, chunk, done);
    slot.received += chunk.length;
    if (slot.received > slot.length) { slot.output?.destroy(new Error('side stream body longer than announced')); slot.output = null; }
    if (chunk.length) slot.output?.write(chunk);
    if (done) {
      if (slot.received !== slot.length) slot.output?.destroy(new Error('side stream body length mismatch'));
      else slot.output?.end();
      slots.delete(id);
    }
  };
  const consume = () => {
    while (!failed) {
      if (parse.mode === 'header') {
        if (buffer.length < 16) return;
        const id = buffer.readBigUInt64BE(0).toString();
        const length = buffer.readBigUInt64BE(8);
        buffer = buffer.subarray(16);
        if (length > BigInt(MAX_SIDE_BODY)) return fail(new Error('side stream body exceeds 256 MiB'));
        parse = { mode: 'body', id, remaining: Number(length) };
        if (!parse.remaining) { deliver(id, Buffer.alloc(0), true); parse = { mode: 'header' }; }
        continue;
      }
      if (!buffer.length) return;
      const take = Math.min(buffer.length, parse.remaining);
      const chunk = Buffer.from(buffer.subarray(0, take));
      buffer = buffer.subarray(take);
      parse.remaining -= take;
      const { id } = parse, done = parse.remaining === 0;
      if (done) parse = { mode: 'header' };
      deliver(id, chunk, done);
    }
  };
  input.on('data', chunk => { if (failed) return; buffer = buffer.length ? Buffer.concat([buffer, chunk]) : Buffer.from(chunk); consume(); });
  input.on('error', fail);
  input.on('end', () => fail(new Error('side stream closed')));
  if (output !== input) output.on('error', fail);

  return {
    get failed() { return failed; },
    /** The body Kit announced as `stream`/`length`; cancelling the reader still drains the shared channel. */
    openBody(stream, length) {
      const output = new PassThrough();
      if (failed || !STREAM.test(String(stream)) || !Number.isSafeInteger(length) || length < 0 || length > MAX_SIDE_BODY || slots.has(String(stream))) {
        output.destroy(failed ?? new Error('side stream unavailable'));
        return output;
      }
      const id = String(stream);
      const slot = { length, received: 0, output };
      slots.set(id, slot);
      output.on('close', () => { slot.output = null; });
      const queued = early.get(id);
      early.delete(id);
      if (queued) {
        for (const chunk of queued.chunks) if (slots.has(id)) deliver(id, chunk, false);
        if (queued.done && slots.has(id)) deliver(id, Buffer.alloc(0), true);
      }
      return output;
    },
    discardBody(stream, length) { this.openBody(stream, length).on('error', () => {}).resume(); },
    /** Writes one framed body for a `begin-media-write` stream; writes never interleave. */
    writeBody(stream, bytes) {
      if (failed) return Promise.reject(failed);
      if (!STREAM.test(String(stream)) || bytes.length > MAX_SIDE_BODY) return Promise.reject(new Error('side stream unavailable'));
      const header = Buffer.alloc(16);
      header.writeBigUInt64BE(BigInt(stream), 0);
      header.writeBigUInt64BE(BigInt(bytes.length), 8);
      const run = writing.then(() => new Promise((resolve, reject) => {
        output.write(Buffer.concat([header, Buffer.from(bytes)]), error => error ? reject(error) : resolve());
      }));
      writing = run.catch(() => {});
      return run;
    },
  };
}
