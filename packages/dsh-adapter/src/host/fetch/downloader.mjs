import { createHash, randomUUID } from 'node:crypto';
import { EventEmitter, once } from 'node:events';
import { createReadStream, createWriteStream, existsSync, readFileSync, readdirSync, statSync } from 'node:fs';
import { mkdir, readFile, rename, rm, stat, writeFile } from 'node:fs/promises';
import { basename, dirname, join } from 'node:path';
import { waitForLease } from './lease.mjs';
import { SourceError, openSource } from './sources.mjs';
import { extractTar } from './tar.mjs';

/**
 * One downloader for everything the universal package fetches (spec 261110-universal-package §4.2, §4.5): office-kit's
 * Kit and editor resources, and skill resources. An item is
 *
 *   { key, dir, files, sources, archive?, strip?, outside?, bundledDir?, prepare?, label? }
 *
 * `dir` is where the verified files end up (written once, then never changed: a new version gets a new dir); `files` the
 * manifest's `{ [path]: { sha256 | sha512, executable? } }`; `archive` `{ file, size, integrity }` for a `.tar.gz` whose
 * sha512 the manifest pins (Kit, editor); without it the source streams a plain tar checked file by file.
 * `prepare(tmpDir)` runs on the unpacked files before they go live (Kit: executable bits, quarantine, self test).
 *
 * Downloads run under one lease per data root (several DSH windows and profiles share it), resume from `.part` files,
 * try each source `attempts` times with growing delays, and emit `progress` events
 * `{ key, label, phase: 'waiting'|'downloading'|'verifying'|'unpacking'|'ready'|'failed', received?, total?, error? }`.
 */
export class DownloadError extends Error {
  constructor(item, failures) {
    const reasons = failures.map(failure => `${failure.source}: ${failure.message}`);
    super(`${item.label ?? item.key} could not be downloaded (${reasons.join('; ') || 'no source'})`);
    this.name = 'DownloadError';
    this.code = failures.length && failures.every(failure => failure.code === 'no-token' || failure.code === 'denied') ? 'denied' : 'download-failed';
    this.failures = failures;
  }
}

export const COMPLETE = '.complete';
const DELAYS_MS = [2000, 8000, 30000];
const IDLE_MS = 60_000;
const RENAME_RETRIES = [100, 200, 400, 800, 1600, 2000, 2000, 2000, 2000, 2000];
const sleep = (ms, signal) => new Promise((resolve, reject) => {
  if (signal?.aborted) return reject(signal.reason);
  const timer = setTimeout(resolve, ms);
  signal?.addEventListener('abort', () => { clearTimeout(timer); reject(signal.reason); }, { once: true });
});
const sourceName = source => source.id ?? source.kind;

/** A stable marker for an item's content: its archive sha512, else a hash of its file list. */
export function markerOf(item) {
  // A skill is the same files whether its pack's sha512 is known yet or not (review R2-F1): its file list decides.
  if (item.archive?.integrity && !String(item.key ?? '').startsWith('skill/')) return item.archive.integrity;
  const text = JSON.stringify(Object.keys(item.files).sort().map(name => [name, item.files[name].sha512 ?? item.files[name].sha256, Boolean(item.files[name].executable)]));
  return `files-sha256-${createHash('sha256').update(text).digest('base64')}`;
}

async function hashFile(file, algorithm) {
  const hash = createHash(algorithm);
  for await (const chunk of createReadStream(file)) hash.update(chunk);
  return `${algorithm}-${hash.digest('base64')}`;
}

/** Problems between the unpacked files and the manifest's list (missing, extra, changed, executable bit). */
export function fileProblems(expected, actual) {
  const problems = [];
  for (const [name, want] of Object.entries(expected)) {
    const got = actual[name];
    if (!got) { problems.push(`missing ${name}`); continue; }
    const algorithm = want.sha512 ? 'sha512' : 'sha256';
    if (got[algorithm] !== want[algorithm]) problems.push(`changed ${name}`);
    if (Boolean(want.executable) !== Boolean(got.executable)) problems.push(`mode ${name}`);
  }
  for (const name of Object.keys(actual)) if (!Object.hasOwn(expected, name)) problems.push(`unexpected ${name}`);
  return problems;
}

export class Downloader extends EventEmitter {
  constructor({ root, open = openSource, delaysMs = DELAYS_MS, idleMs = IDLE_MS, fetchImpl, spawn, now = Date.now, lease = {}, log = () => {}, renameDelaysMs = RENAME_RETRIES }) {
    super();
    Object.assign(this, { root, open, delaysMs, idleMs, fetchImpl, spawn, now, leaseOptions: lease, log, renameDelaysMs });
    this.inflight = new Map();
    this.prepared = new Set();
    this.via = new Map();
  }

  get downloads() { return join(this.root, 'downloads'); }
  get stateFile() { return join(this.root, 'state.json'); }

  /** Progress of `item`; while a download runs it names the source in use and the sources it fell back from. */
  progress(item, fields) { this.emit('progress', { key: item.key, label: item.label, source: undefined, fallback: undefined, ...this.via.get(item.key), ...fields }); }

  /** Whether `item.dir` holds this exact item: its `.complete` marker and every listed file present. Cheap: no hashing. */
  isComplete(item) {
    try {
      const marker = JSON.parse(readFileSync(join(item.dir, COMPLETE), 'utf8'));
      if (marker.marker !== markerOf(item)) return false;
      return Object.keys(item.files).every(name => existsSync(join(item.dir, ...name.split('/'))));
    } catch { return false; }
  }

  /**
   * The directory holding `item`, downloading it when needed. Bundled copies (an offline package) are used in place
   * when every listed file is there. Concurrent calls for one key share one download.
   * Options: `signal`, `budgetMs` (give up waiting after this; the download keeps going in the background unless
   * `signal` aborts), `firstSourceOnly`.
   */
  async ensure(item, { signal, budgetMs, firstSourceOnly = false } = {}) {
    const bundled = item.bundledDir && Object.keys(item.files).every(name => existsSync(join(item.bundledDir, ...name.split('/'))));
    if (bundled) {
      if (item.prepare && !this.prepared.has(item.bundledDir)) { await item.prepare(item.bundledDir, { bundled: true }); this.prepared.add(item.bundledDir); }
      // Not touched: `lastUsed` belongs to the data directory copy, which a packaged copy leaves to age out (prune).
      return { dir: item.bundledDir, source: 'bundled' };
    }
    if (this.isComplete(item)) { await this.touch(item); return { dir: item.dir, source: 'cache' }; }
    let running = this.inflight.get(item.key);
    if (!running) {
      running = this.download(item, { signal: budgetMs ? undefined : signal, firstSourceOnly }).finally(() => this.inflight.delete(item.key));
      this.inflight.set(item.key, running);
    }
    if (!budgetMs) return running;
    let timer;
    const budget = new Promise((_, reject) => { timer = setTimeout(() => reject(Object.assign(new Error(`${item.label ?? item.key} is still downloading`), { code: 'still-downloading' })), budgetMs); });
    const aborted = signal ? new Promise((_, reject) => signal.addEventListener('abort', () => reject(signal.reason), { once: true })) : new Promise(() => {});
    try { return await Promise.race([running, budget, aborted]); } finally { clearTimeout(timer); }
  }

  async download(item, { signal, firstSourceOnly }) {
    this.progress(item, { phase: 'waiting' });
    const lease = await waitForLease(join(this.root, '.lock'), { signal, onWait: () => this.progress(item, { phase: 'waiting' }), ...this.leaseOptions });
    try {
      // Another process may have finished it while this one waited.
      if (this.isComplete(item)) { this.progress(item, { phase: 'ready' }); return { dir: item.dir, source: 'cache' }; }
      const manual = item.archive && await this.manualCopy(item);
      if (manual) {
        try { await this.install(item, { kind: 'manual' }, createReadStream(manual), 'tgz', signal); await rm(manual, { force: true }); return { dir: item.dir, source: 'manual' }; }
        catch (error) { this.log('warn', `ignored ${manual}: ${error.message}`); }
      }
      const failures = [];
      const sources = firstSourceOnly ? item.sources.slice(0, 1) : item.sources;
      const fellBack = [];
      for (const [index, source] of sources.entries()) {
        // A source may set its own number of tries (download-sources.json `attempts`; the editor's npm source: NPM_DEFAULTS).
        const attempts = firstSourceOnly ? 1 : Math.max(1, Math.min(source.attempts ?? this.delaysMs.length, this.delaysMs.length));
        this.via.set(item.key, { source: sourceName(source), ...fellBack.length && { fallback: [...fellBack] } });
        for (let attempt = 1; attempt <= attempts; attempt++) {
          signal?.throwIfAborted();
          try {
            await this.fetchFrom(item, source, signal);
            this.progress(item, { phase: 'ready' });
            await this.touch(item, { lastError: null, source: sourceName(source) });
            return { dir: item.dir, source: sourceName(source) };
          } catch (error) {
            if (signal?.aborted) throw signal.reason;
            failures.push({ source: sourceName(source), message: error?.message ?? String(error), code: error?.code });
            this.log('warn', `${item.key} from ${sourceName(source)} (try ${attempt}): ${error?.message ?? error}`);
            // `final`: the bytes were fine and another source would deliver the same (a Kit that does not start).
            if (error?.final) { const failed = new DownloadError(item, failures); this.progress(item, { phase: 'failed', error: failed.message }); throw failed; }
            if (error?.retry === false || attempt === attempts) break;
            await sleep(this.delaysMs[attempt - 1], signal);
          }
        }
        fellBack.push(sourceName(source));
        const next = sources[index + 1];
        if (next) this.log('warn', `${item.key}: ${sourceName(source)} unavailable (${failures.at(-1)?.message}); falling back to ${sourceName(next)}`);
      }
      const error = new DownloadError(item, failures);
      this.progress(item, { phase: 'failed', error: error.message });
      await this.touch(item, { lastError: { at: this.now(), message: error.message } }, { used: false });
      throw error;
    } finally { this.via.delete(item.key); lease.release(); }
  }

  /** An archive the user put into downloads/ by hand (any name) with the manifest's size and sha512 (spec §5). */
  async manualCopy(item) {
    let names;
    try { names = readdirSync(this.downloads); } catch { return undefined; }
    for (const name of names) {
      if (name.endsWith('.part') || name.endsWith('.part.json')) continue;
      const file = join(this.downloads, name);
      try {
        if (statSync(file).size !== item.archive.size) continue;
        if (await hashFile(file, 'sha512') === item.archive.integrity) return file;
      } catch { /* unreadable: not ours */ }
    }
    return undefined;
  }

  async fetchFrom(item, source, signal) {
    // One controller per attempt: an idle body is cut off by aborting its request, not by waiting for its reader.
    const attempt = new AbortController();
    const requestSignal = signal ? AbortSignal.any([signal, attempt.signal]) : attempt.signal;
    const cut = error => attempt.abort(error);
    // Without a pinned archive the source streams the files themselves (checked one by one).
    if (!item.archive) {
      const opened = await this.open(source, item, { signal: requestSignal, fetchImpl: this.fetchImpl, spawn: this.spawn });
      return this.install(item, source, this.watch(opened.body, signal, cut, source.idleMs), opened.format ?? 'tar', signal);
    }
    await mkdir(this.downloads, { recursive: true });
    const part = join(this.downloads, `${basename(item.archive.file)}.part`);
    const metaFile = `${part}.json`;
    let meta = {};
    try { meta = JSON.parse(await readFile(metaFile, 'utf8')); } catch {}
    let offset = 0;
    try { offset = meta.source === sourceName(source) && meta.integrity === item.archive.integrity && meta.etag ? (await stat(part)).size : 0; } catch {}
    // A complete .part that already matches (a run whose later step failed, e.g. the Kit self test): no new download.
    if (offset === item.archive.size && await hashFile(part, 'sha512').catch(() => '') === item.archive.integrity) {
      await this.install(item, source, createReadStream(part), 'tgz', signal);
      await rm(part, { force: true }); await rm(metaFile, { force: true });
      return;
    }
    if (offset >= item.archive.size) offset = 0;
    const opened = await this.open(source, item, { offset, etag: meta.etag, signal: requestSignal, fetchImpl: this.fetchImpl, spawn: this.spawn });
    if (!opened.resumed) offset = 0;
    await writeFile(metaFile, JSON.stringify({ source: sourceName(source), etag: opened.etag ?? null, integrity: item.archive.integrity }));
    const out = createWriteStream(part, { flags: offset ? 'a' : 'w' });
    const closed = new Promise((resolve, reject) => { out.once('close', resolve); out.once('error', reject); });
    let received = offset;
    try {
      for await (const chunk of this.watch(opened.body, signal, cut, source.idleMs)) {
        received += chunk.length;
        if (received > item.archive.size) throw new SourceError(`${item.archive.file} is larger than ${item.archive.size} bytes`, { code: 'size' });
        if (!out.write(chunk)) await once(out, 'drain');
        this.progress(item, { phase: 'downloading', received, total: item.archive.size });
      }
    } finally { out.end(); await closed.catch(() => {}); }
    if (received !== item.archive.size) throw new SourceError(`${item.archive.file}: got ${received} of ${item.archive.size} bytes`, { code: 'size' });
    this.progress(item, { phase: 'verifying', received, total: item.archive.size });
    const integrity = await hashFile(part, 'sha512');
    if (integrity !== item.archive.integrity) {
      await rm(part, { force: true }); await rm(metaFile, { force: true });
      throw new SourceError(`${item.archive.file}: sha512 does not match the manifest`, { code: 'integrity' });
    }
    await this.install(item, source, createReadStream(part), 'tgz', signal);
    await rm(part, { force: true }); await rm(metaFile, { force: true });
  }

  /** Passes the body through, failing (and cutting the request off) when no data arrives for `idleMs` (the source's own, if it sets one). */
  async* watch(body, signal, cut = () => {}, idleMs = this.idleMs) {
    const iterator = body[Symbol.asyncIterator]();
    for (;;) {
      let timer;
      const idle = new Promise((_, reject) => { timer = setTimeout(() => { const error = new SourceError(`no data for ${idleMs / 1000} s`, { code: 'timeout' }); cut(error); reject(error); }, idleMs); });
      let next;
      try { next = await Promise.race([iterator.next(), idle]); }
      catch (error) {
        // Never awaited: a reader stuck in a pending read would hold this attempt forever.
        iterator.return?.()?.catch?.(() => {});
        throw signal?.aborted ? signal.reason : error;
      } finally { clearTimeout(timer); }
      if (next.done) return;
      signal?.throwIfAborted();
      yield next.value;
    }
  }

  /** Unpacks into a sibling temporary directory, checks every file, prepares it, then renames it into place. */
  async install(item, source, stream, format, signal) {
    const parent = dirname(item.dir);
    await mkdir(parent, { recursive: true });
    const tmp = join(parent, `.${basename(item.dir)}.tmp-${randomUUID()}`);
    try {
      this.progress(item, { phase: 'unpacking' });
      // `strip` is the archive's prefix.
      const strip = item.strip ?? '';
      const { files } = await extractTar(stream, tmp, { gzip: format === 'tgz', strip, outside: item.outside ?? 'reject' });
      signal?.throwIfAborted();
      const problems = fileProblems(item.files, files);
      if (problems.length) throw new SourceError(`${item.label ?? item.key} does not match the manifest: ${problems.slice(0, 5).join(', ')}${problems.length > 5 ? ` and ${problems.length - 5} more` : ''}`, { code: 'integrity' });
      if (item.prepare) await item.prepare(tmp, { bundled: false });
      await writeFile(join(tmp, COMPLETE), JSON.stringify({ marker: markerOf(item), source: sourceName(source), at: new Date(this.now()).toISOString() }));
      // A leftover incomplete directory (an interrupted run before this layout) is moved aside first.
      if (existsSync(item.dir)) {
        const aside = join(parent, `.${basename(item.dir)}.old-${randomUUID()}`);
        await this.retryRename(item.dir, aside);
        await rm(aside, { recursive: true, force: true }).catch(() => {});
      }
      await this.retryRename(tmp, item.dir);
    } catch (error) {
      await rm(tmp, { recursive: true, force: true }).catch(() => {});
      throw error;
    }
  }

  /** Windows: a freshly written exe is often held briefly by Defender or the indexer (EPERM/EBUSY/EACCES). */
  async retryRename(from, to) {
    for (let attempt = 0; ; attempt++) {
      try { return await rename(from, to); }
      catch (error) {
        if (!['EPERM', 'EBUSY', 'EACCES'].includes(error?.code) || attempt >= this.renameDelaysMs.length) throw error;
        await sleep(this.renameDelaysMs[attempt]);
      }
    }
  }

  readState() { try { return JSON.parse(readFileSync(this.stateFile, 'utf8')); } catch { return { items: {} }; } }
  /** Records use (and the last error) per item key; best effort, never fatal. */
  async touch(item, fields = {}, { used = true } = {}) {
    try {
      await mkdir(this.root, { recursive: true });
      const state = this.readState();
      state.items ??= {};
      const previous = state.items[item.key] ?? {};
      state.items[item.key] = { ...previous, dir: item.dir, ...used && { lastUsed: this.now() }, ...fields };
      if (fields.lastError === null) delete state.items[item.key].lastError;
      const tmp = `${this.stateFile}.${randomUUID()}.tmp`;
      await writeFile(tmp, JSON.stringify(state, null, 2));
      await rename(tmp, this.stateFile);
    } catch (error) { this.log('debug', `state.json not written: ${error?.message}`); }
  }

  /** The last failure recorded for `key`, if any. */
  lastError(key) { return this.readState().items?.[key]?.lastError ?? undefined; }

  /**
   * Removes item directories under `dirs` that are not kept: `keep(dir)` true, used within `maxAgeMs`, or locked.
   * Returns the removed directories. A directory that cannot be removed (Windows: in use) is left for next time.
   */
  async prune({ dirs, keep = () => false, maxAgeMs = 30 * 24 * 3600_000 }) {
    const state = this.readState();
    const lastUsed = new Map(Object.values(state.items ?? {}).filter(entry => entry?.dir).map(entry => [entry.dir, entry.lastUsed ?? 0]));
    const removed = [];
    for (const dir of dirs) {
      if (keep(dir)) continue;
      if (this.now() - (lastUsed.get(dir) ?? 0) < maxAgeMs) continue;
      try { await rm(dir, { recursive: true, force: true }); removed.push(dir); } catch (error) { this.log('debug', `kept ${dir}: ${error?.message}`); }
    }
    return removed;
  }
}
