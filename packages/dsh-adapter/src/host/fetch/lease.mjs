import { randomUUID } from 'node:crypto';
import { closeSync, linkSync, mkdirSync, openSync, readFileSync, renameSync, rmSync, statSync, utimesSync, writeSync } from 'node:fs';
import { dirname } from 'node:path';

/**
 * A lock file held as a lease (spec 261110-universal-package §4.2 `.lock`): created with O_EXCL and holding
 * `<pid>:<holder uuid>` (one DSH process may run several profiles' copies of the plugin, so the pid alone does not tell
 * holders apart). The holder touches it every `renewMs`; one not touched for `staleMs` is stale whoever wrote it, and is
 * taken over by renaming it aside and creating a new one, then reading it back to be sure the new one is ours.
 */
export const RENEW_MS = 10_000;
export const STALE_MS = 60_000;

export class LeaseBusy extends Error {
  constructor(file, holder) { super(`${file} is held by ${holder}`); this.name = 'LeaseBusy'; this.code = 'busy'; this.holder = holder; }
}

function create(file, mine) {
  const fd = openSync(file, 'wx', 0o600);
  try { writeSync(fd, mine); } finally { closeSync(fd); }
}
const read = file => { try { return readFileSync(file, 'utf8'); } catch { return undefined; } };

/**
 * Takes the lease or throws LeaseBusy. Returns `{ release, holder }`; the lease renews itself until released.
 * `now`/`setInterval` are injectable for tests.
 */
export function acquireLease(file, { holder = randomUUID(), staleMs = STALE_MS, renewMs = RENEW_MS, now = Date.now, timers = { setInterval, clearInterval }, beforeTakeover = () => {} } = {}) {
  mkdirSync(dirname(file), { recursive: true });
  const mine = `${process.pid}:${holder}`;
  try { create(file, mine); }
  catch (error) {
    if (error?.code !== 'EEXIST') throw error;
    let age = 0;
    try { age = now() - statSync(file).mtimeMs; } catch { /* vanished: try again below */ }
    const stale = read(file);
    if (age < staleMs && stale !== undefined) throw new LeaseBusy(file, stale);
    // Stale: move it aside (only one taker wins the rename of this exact file), then create ours. What was moved must be
    // the lease judged stale: another taker may have replaced it in between, and then it goes back and this one waits.
    const aside = `${file}.stale-${randomUUID()}`;
    beforeTakeover();
    let moved = false;
    try { renameSync(file, aside); moved = true; } catch { /* another taker moved it first */ }
    if (moved) {
      let fresh = false;
      try { fresh = readFileSync(aside, 'utf8') !== stale || now() - statSync(aside).mtimeMs < staleMs; } catch {}
      if (fresh) {
        // link, not rename: never overwrite a lease someone created meanwhile.
        try { linkSync(aside, file); } catch { /* a newer lease is there */ }
        rmSync(aside, { force: true });
        throw new LeaseBusy(file, read(file));
      }
      rmSync(aside, { force: true });
    }
    try { create(file, mine); } catch (again) { if (again?.code === 'EEXIST') throw new LeaseBusy(file, read(file)); throw again; }
  }
  if (read(file) !== mine) throw new LeaseBusy(file, read(file));
  const timer = timers.setInterval(() => { try { if (read(file) === mine) { const t = new Date(now()); utimesSync(file, t, t); } } catch { /* the next renewal tries again */ } }, renewMs);
  timer?.unref?.();
  let released = false;
  return {
    holder: mine,
    release() {
      if (released) return;
      released = true;
      timers.clearInterval(timer);
      if (read(file) === mine) rmSync(file, { force: true });
    },
  };
}

/** Waits for the lease, polling every `pollMs`, until `signal` aborts or `timeoutMs` passes. */
export async function waitForLease(file, { signal, pollMs = 1000, timeoutMs = Infinity, onWait = () => {}, ...options } = {}) {
  const started = Date.now();
  for (;;) {
    try { return acquireLease(file, options); }
    catch (error) {
      if (!(error instanceof LeaseBusy)) throw error;
      if (Date.now() - started > timeoutMs) throw error;
      onWait(error.holder);
    }
    signal?.throwIfAborted();
    await new Promise((resolve, reject) => {
      const timer = setTimeout(resolve, pollMs);
      signal?.addEventListener('abort', () => { clearTimeout(timer); reject(signal.reason); }, { once: true });
    });
  }
}
