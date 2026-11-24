/** Kit command frames (protocol "5", kit-usage/02-frames.md): 4-byte big-endian length, then UTF-8 JSON. */
export const PROTOCOL = '5';
export const MAX_FRAME = 67_108_864;
export const MAX_IN_FLIGHT = 8;

export class FrameError extends Error {
  constructor(message) { super(message); this.name = 'FrameError'; }
}

/** One command as bytes; a body over the limit is an explicit failure, nothing is written. */
export function encodeFrame(command) {
  const body = Buffer.from(JSON.stringify(command), 'utf8');
  if (body.length > MAX_FRAME) throw new FrameError('frame exceeds 64 MiB');
  const header = Buffer.alloc(4);
  header.writeUInt32BE(body.length, 0);
  return Buffer.concat([header, body]);
}

/** Splits stdout bytes into JSON frames. An oversized or non-JSON frame breaks the channel (result unknown). */
export class FrameDecoder {
  #buffer = Buffer.alloc(0);
  push(chunk) {
    this.#buffer = this.#buffer.length ? Buffer.concat([this.#buffer, chunk]) : Buffer.from(chunk);
    const frames = [];
    while (this.#buffer.length >= 4) {
      const length = this.#buffer.readUInt32BE(0);
      if (length > MAX_FRAME) throw new FrameError('response frame exceeds 64 MiB');
      if (this.#buffer.length < 4 + length) break;
      const body = this.#buffer.subarray(4, 4 + length);
      this.#buffer = this.#buffer.subarray(4 + length);
      let frame;
      try { frame = JSON.parse(body.toString('utf8')); }
      catch { throw new FrameError('response frame is not JSON'); }
      if (!frame || typeof frame !== 'object' || Array.isArray(frame)) throw new FrameError('response frame is not an object');
      frames.push(frame);
    }
    return frames;
  }
}

/** String and number ids are different ids (kit-usage/02-frames.md). */
export const idKey = id => `${typeof id}:${id}`;
