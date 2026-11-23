import { spawn as spawnChild } from 'node:child_process';
import { EventEmitter } from 'node:events';
import { randomUUID } from 'node:crypto';
import { mkdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import * as nodePath from 'node:path';
import { KitClient } from './client.mjs';
import { createSideStream } from './side-stream.mjs';

const RESTART_MS = [500, 1000, 2000, 5000, 10_000];
const MAX_RESTARTS = 5;
const LOCK = '.officesdk-kit.lock';
/** `{ pid, binary }` of the Kit started on this work root last. */
export const RUNNING = '.officesdk-kit.running';

const alive = pid => { try { process.kill(pid, 0); return true; } catch (error) { return error.code === 'EPERM'; } };

/** Takes one lock file, or throws `work-root-locked`; returns its release. */
function lockFile(file, mine) {
  for (let attempt = 0; attempt < 2; attempt++) {
    try {
      writeFileSync(file, mine, { flag: 'wx' });
      return () => { try { if (readFileSync(file, 'utf8') === mine) rmSync(file); } catch {} };
    } catch (error) {
      if (error.code !== 'EEXIST') throw error;
      const owner = Number(readFileSync(file, 'utf8').split(':')[0]);
      if (Number.isSafeInteger(owner) && owner > 0 && alive(owner)) throw Object.assign(new Error(`The document work root is in use by process ${owner}.`), { code: 'work-root-locked' });
      rmSync(file, { force: true });
    }
  }
  throw Object.assign(new Error('The document work root could not be locked.'), { code: 'work-root-locked' });
}

/**
 * One Kit per work root (Kit itself does not lock it). The lock names the process and the holder, so a
 * second plugin instance in the same process (a reload overlapping dispose) is refused too; only a lock
 * whose process is gone is taken over.
 */
export function lockWorkRoot(work, holder = randomUUID()) {
  mkdirSync(work, { recursive: true });
  const mine = `${process.pid}:${holder}`;
  return lockFile(nodePath.join(work, LOCK), mine);
}

/**
 * Supervises `offline-kit-runtime serve --work <work>`. Events: `ready` ({ generation }) after each start,
 * `down` ({ generation, error }) when the channel or process is lost — every view of that generation is
 * stale and must reopen — and `log` with stderr text. Restarts back off; `stop` ends Kit and waits for it;
 * `suspend` also keeps it from starting again until `resume` (a plugin update replaces its executables).
 */
export class KitProcess extends EventEmitter {
  #child = null;
  #client = null;
  #generation = 0;
  #release = null;
  #stopping = false;
  #restartTimer = null;
  #attempt = 0;
  #starting = null;
  #suspended = false;

  /** `platform` picks the side-stream pipes and how Kit is ended; it is Node's `process.platform`. */
  constructor({ binary, work, spawn = spawnChild, restartMs = RESTART_MS, maxRestarts = MAX_RESTARTS, killGraceMs = 2000, platform = process.platform }) {
    super();
    Object.assign(this, { binary, work, spawn, restartMs, maxRestarts, killGraceMs, platform });
  }

  get generation() { return this.#generation; }
  /** The current Kit process id, if one was started. */
  get pid() { return this.#child?.pid; }
  /** The live client, or null while Kit is starting, restarting or stopped. */
  get client() { return this.#client && !this.#client.broken ? this.#client : null; }

  /** Starts Kit unless one is already running; concurrent callers (a restart timer, a request) share one launch. */
  start() {
    if (this.#suspended) return Promise.reject(Object.assign(new Error('The document service is stopped while the plugin updates.'), { code: 'kit-suspended' }));
    this.#stopping = false;
    if (this.client) return Promise.resolve(this.client);
    this.#release ??= lockWorkRoot(this.work);
    clearTimeout(this.#restartTimer);
    this.#starting ??= this.#launch().finally(() => { this.#starting = null; });
    return this.#starting;
  }

  /** Resolves with a live client, starting Kit if needed. */
  async ready() { return this.client ?? await this.start(); }

  async #launch() {
    const windows = this.platform === 'win32';
    // Windows has no duplex fd 3: Kit reads the Host's media on index 3 and writes media and fonts on index 4.
    // windowsHide keeps a console window from flashing up when the desktop app starts Kit.
    // `binary` may be a function resolving the path (the universal package downloads Kit first, spec
    // 261110-universal-package §4.2); its error (download failed, self test failed, downgrade refused) is this start's.
    const binary = typeof this.binary === 'function' ? await this.binary() : this.binary;
    // Stopped or suspended while the binary was being fetched (a reload, an update): this instance no longer holds the
    // work directory, so it must not start a Kit there.
    if (this.#stopping || this.#suspended) throw Object.assign(new Error('The document service was stopped while it was being prepared.'), { code: 'kit-suspended' });
    const child = this.spawn(binary, ['serve', '--work', this.work], { stdio: windows ? ['pipe', 'pipe', 'pipe', 'pipe', 'pipe'] : ['pipe', 'pipe', 'pipe', 'pipe'], env: { ...process.env }, windowsHide: true });
    this.#child = child;
    // Which Kit runs on this work root (spec 261110-universal-package §4.3 cleanup keeps its directory).
    try { writeFileSync(nodePath.join(this.work, RUNNING), JSON.stringify({ pid: child.pid ?? null, binary })); } catch { /* best effort */ }
    const exited = new Promise(resolve => child.once('exit', (code, signal) => resolve({ code, signal })));
    child.exited = exited;
    child.stderr?.setEncoding('utf8');
    child.stderr?.on('data', text => this.emit('log', text));
    const side = windows ? (child.stdio?.[3] && child.stdio?.[4] ? createSideStream(child.stdio[4], child.stdio[3]) : null)
      : child.stdio?.[3] ? createSideStream(child.stdio[3]) : null;
    const generation = this.#generation + 1;
    const client = new KitClient({ stdin: child.stdin, stdout: child.stdout, side, onBroken: error => this.#lost(child, generation, error) });
    child.once('error', error => client.close(error));
    exited.then(({ code, signal }) => { if (!client.broken) client.close(); this.#lost(child, generation, new Error(`Kit exited (${signal ?? code})`)); });
    try { await client.ready(); }
    catch (error) { await this.#end(child); throw error; }
    this.#client = client;
    this.#generation = generation;
    this.#attempt = 0;
    this.emit('ready', { generation });
    return client;
  }

  #lost(child, generation, error) {
    if (child.lost) return;
    child.lost = true;
    if (this.#child === child) { this.#client = null; }
    if (generation <= this.#generation) this.emit('down', { generation, error });
    void this.#end(child);
    if (this.#stopping || this.#child !== child) return;
    // A Kit that keeps failing is left down; the next request starts it again (see `ready`).
    if (this.#attempt >= this.maxRestarts) { this.emit('gave-up', { error }); return; }
    const delay = this.restartMs[Math.min(this.#attempt++, this.restartMs.length - 1)];
    clearTimeout(this.#restartTimer);
    this.#restartTimer = setTimeout(() => { this.start().catch(() => this.#lost(this.#child ?? child, this.#generation, error)); }, delay);
    this.#restartTimer.unref?.();
  }

  /**
   * Ends one process and waits for its exit: stdin closed, then SIGTERM and SIGKILL after the grace period.
   * Windows has no SIGTERM — kill() terminates at once — so there Kit gets the grace period after stdin
   * closes to leave on its own, and is killed only then.
   */
  #end(child) {
    if (child.exitCode !== null || child.signalCode !== null) return child.exited;
    // stop() and the broken channel both end the same child: one grace period, one kill.
    child.ending ??= this.#terminate(child);
    return child.ending;
  }

  async #terminate(child) {
    try { child.stdin?.end(); } catch {}
    const windows = this.platform === 'win32';
    if (!windows) try { child.kill('SIGTERM'); } catch {}
    const timer = setTimeout(() => { try { windows ? child.kill() : child.kill('SIGKILL'); } catch {} }, this.killGraceMs);
    timer.unref?.();
    await child.exited;
    clearTimeout(timer);
  }

  get suspended() { return this.#suspended; }
  /** Stops Kit and keeps it stopped: requests fail with `kit-suspended` and no restart runs until `resume`. */
  async suspend() {
    this.#suspended = true;
    await this.stop();
  }
  /** Lets Kit start again; the next request starts it. */
  resume() { this.#suspended = false; }

  async stop() {
    this.#stopping = true;
    clearTimeout(this.#restartTimer);
    const child = this.#child;
    this.#client?.close();
    this.#client = null;
    if (child) await this.#end(child);
    this.#release?.();
    this.#release = null;
  }
}
