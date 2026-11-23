import { FrameDecoder, FrameError, MAX_IN_FLIGHT, PROTOCOL, encodeFrame, idKey } from './frames.mjs';

/** A command whose frame reached Kit but whose response never came: Kit may or may not have run it. */
export const UNKNOWN = Object.freeze({ ok: false, unknown: true });

/**
 * One connection to a running Kit. `request` resolves with Kit's response frame, an explicit failure
 * (`ok: false` with a `code`, nothing reached Kit or Kit refused it) or `UNKNOWN`. It never rejects.
 * The channel breaks — every in-flight command becomes unknown — when stdout ends, a response is not a
 * JSON object or exceeds 64 MiB, or a frame after `ready` carries no id (kit-usage/05-host.md).
 */
export class KitClient {
  #decoder = new FrameDecoder();
  #pending = new Map();
  #waiting = [];
  #inFlight = 0;
  #nextId = 1;
  #ready;
  #readyFrame;
  #broken = null;

  constructor({ stdin, stdout, side = null, onBroken = () => {} }) {
    this.stdin = stdin;
    this.side = side;
    this.onBroken = onBroken;
    this.#ready = new Promise((resolve, reject) => { this.#readyFrame = { resolve, reject }; });
    this.#ready.catch(() => {});
    stdout.on('data', chunk => this.#consume(chunk));
    stdout.on('end', () => this.#break(new Error('Kit closed stdout')));
    stdout.on('error', error => this.#break(error));
    stdin.on('error', error => this.#break(error));
  }

  get broken() { return this.#broken; }
  /** Resolves once Kit announced protocol "5"; rejects for any other first frame. */
  ready() { return this.#ready; }

  request(op, fields = {}, { signal } = {}) {
    if (this.#broken) return Promise.resolve({ ok: false, code: 'kit-unavailable', error: 'The document service is not running.' });
    return new Promise(resolve => {
      const entry = { op, fields, resolve, written: false, signal };
      const abort = () => {
        if (entry.written) return;
        this.#waiting = this.#waiting.filter(item => item !== entry);
        resolve({ ok: false, code: 'cancelled', error: 'The request was cancelled before it reached the document service.' });
      };
      if (signal?.aborted) return abort();
      signal?.addEventListener('abort', abort, { once: true });
      entry.done = frame => { signal?.removeEventListener('abort', abort); resolve(frame); };
      this.#waiting.push(entry);
      this.#pump();
    });
  }

  /** Every command waits for `ready`; at most 8 are written and unanswered at once. */
  async #pump() {
    try { await this.#ready; } catch { return; }
    while (!this.#broken && this.#inFlight < MAX_IN_FLIGHT && this.#waiting.length) {
      const entry = this.#waiting.shift();
      const id = String(this.#nextId++);
      let bytes;
      try { bytes = encodeFrame({ ...entry.fields, op: entry.op, id }); }
      catch (error) { entry.done({ ok: false, code: 'frame-too-large', error: error.message }); continue; }
      entry.written = true;
      this.#inFlight++;
      this.#pending.set(idKey(id), entry);
      try { this.stdin.write(bytes); }
      catch (error) { this.#break(error); }
    }
  }

  #consume(chunk) {
    let frames;
    try { frames = this.#decoder.push(chunk); }
    catch (error) { return this.#break(error instanceof FrameError ? error : new FrameError(String(error))); }
    for (const frame of frames) {
      if (this.#broken) return;
      if (!this.#readyFrame.settled) {
        this.#readyFrame.settled = true;
        if (frame.ready === true && frame.protocolVersion === PROTOCOL && !('id' in frame)) this.#readyFrame.resolve(frame);
        else {
          const error = new Error(frame.ready === true ? `Kit speaks protocol ${frame.protocolVersion}, expected ${PROTOCOL}` : 'Kit did not become ready');
          this.#readyFrame.reject(error);
          this.#break(error);
        }
        continue;
      }
      if (typeof frame.id !== 'string' && typeof frame.id !== 'number') return this.#break(new FrameError('response without id'));
      const entry = this.#pending.get(idKey(frame.id));
      if (!entry) {
        // A response nobody waits for still owns its side-stream bytes.
        if (frame.ok && typeof frame.stream === 'string' && Number.isSafeInteger(frame.length)) this.side?.discardBody(frame.stream, frame.length);
        continue;
      }
      this.#pending.delete(idKey(frame.id));
      this.#inFlight--;
      entry.done(frame);
      this.#pump();
    }
  }

  #break(error) {
    if (this.#broken) return;
    this.#broken = error;
    if (!this.#readyFrame.settled) { this.#readyFrame.settled = true; this.#readyFrame.reject(error); }
    for (const entry of this.#pending.values()) entry.done(UNKNOWN);
    this.#pending.clear();
    for (const entry of this.#waiting.splice(0)) entry.done({ ok: false, code: 'kit-unavailable', error: 'The document service stopped before the request was sent.' });
    this.#inFlight = 0;
    this.onBroken(error);
  }

  /** Stops using the channel; callers end the process separately. */
  close() { this.#break(new Error('Kit client closed')); }
}
