import { randomBytes } from 'node:crypto';
import { lstat, mkdir, mkdtemp, readdir, readFile, rename, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { basename, dirname, join } from 'node:path';
import { DraftStore } from './drafts.mjs';
import { EDITS } from './edits.mjs';
import { SuiteError, fail, fromAbort } from './errors.mjs';
import { FORMATS, capabilities, contentType, formatOf } from './formats.mjs';
import { normalizeLimits } from './limits.mjs';
import { RuntimeReaders } from './runtime-readers.mjs';
import { confinedFile, measure, mime, readConfined, safeRelative } from './paths.mjs';
import { checkOfficePackage } from './zip.mjs';

const TASK_PREFIX = 'officesdk-';
const never = new AbortController().signal;
const scopeKey = scope => JSON.stringify([scope.sessionId, scope.workspaceRoot]);
const sameVersion = (a, b) => a.version === b.version && a.absolutePath === b.absolutePath;
const anySignal = values => AbortSignal.any(values.filter(Boolean));
const token = bytes => randomBytes(bytes).toString('hex');

/** Signals settle as their own error code: a deadline is a timeout, everything else a cancellation. */
async function guarded(signal, operation) {
  try { signal.throwIfAborted(); return await operation(); }
  catch (error) { if (signal.aborted) throw fromAbort(signal); throw error; }
}

/**
 * One document session per open: imported snapshot, optional edits, saves and exports.
 * Hosts inject reading, writing, authorization and conversion; nothing here knows DSH.
 */
export class SuiteService {
  constructor({ source, writer, authorizer, converter, runtimeRoot, tempRoot = tmpdir(), dataRoot, routePrefix = '/shimo', limits, now = Date.now, hostCapabilities = {} }) {
    if (!source || !authorizer || !converter || !runtimeRoot || !dataRoot) throw new TypeError('SuiteService requires source, authorizer, converter, runtimeRoot and dataRoot.');
    this.source = source;
    this.writer = writer;
    this.authorizer = authorizer;
    this.converter = converter;
    this.runtimeRoot = runtimeRoot;
    this.tempRoot = tempRoot;
    this.routePrefix = routePrefix;
    this.limits = normalizeLimits(limits);
    this.now = now;
    this.hostCapabilities = writer ? hostCapabilities : { ...hostCapabilities, save: false, edit: false };
    this.drafts = new DraftStore({ root: join(dataRoot, 'drafts'), limits: this.limits, now });
    this.lifetime = new AbortController();
    this.sessions = new Map();
    this.downloads = new Map();
    this.operations = new Set();
    this.closing = new Set();
    this.waiters = [];
    this.active = 0;
    // Over the limit, reads wait their turn: a deck can reference more pictures than there are slots.
    this.readers = new RuntimeReaders(this.limits.maxReaders);
    this.reservedBytes = 0;
    this.retainedBytes = 0;
    this.started = undefined;
    this.sweeper = setInterval(() => { this.expire().catch(() => {}); }, Math.min(this.limits.leaseTtlMs, 30_000));
    this.sweeper.unref();
  }

  capabilities(format) { return capabilities(format, this.converter.capabilities?.() ?? {}, this.hostCapabilities); }

  /** Orphaned task directories come from a crashed host; nothing else uses the prefix. */
  start() {
    this.started ??= (async () => {
      await mkdir(this.tempRoot, { recursive: true, mode: 0o700 });
      for (const name of await readdir(this.tempRoot)) if (name.startsWith(TASK_PREFIX)) await rm(join(this.tempRoot, name), { recursive: true, force: true });
      await this.drafts.load();
    })();
    return this.started;
  }

  track(operation) {
    this.operations.add(operation);
    return operation.finally(() => this.operations.delete(operation));
  }

  async acquire(signal) {
    signal.throwIfAborted();
    if (this.active < this.limits.maxConcurrent) { this.active++; return; }
    if (this.waiters.length >= this.limits.maxQueued) fail('busy', 'The conversion queue is full.');
    await new Promise((resolve, reject) => {
      const entry = { resolve, signal, abort: () => { this.waiters = this.waiters.filter(v => v !== entry); reject(fromAbort(signal)); } };
      signal.addEventListener('abort', entry.abort, { once: true });
      this.waiters.push(entry);
    });
  }
  relinquish() {
    const next = this.waiters.shift();
    if (next) { next.signal.removeEventListener('abort', next.abort); next.resolve(); }
    else this.active--;
  }
  reserve(bytes) {
    if (this.retainedBytes + this.reservedBytes + bytes > this.limits.maxRetainedBytes) fail('busy', 'The retained byte limit is reached.');
    this.reservedBytes += bytes;
    let released = false;
    return () => { if (!released) { released = true; this.reservedBytes -= bytes; } };
  }
  deadline(operation, format) {
    const controller = new AbortController();
    const ms = this.limits.timeoutMs[operation][format];
    const timer = setTimeout(() => controller.abort(new SuiteError('timeout', 'The conversion timed out.')), ms);
    timer.unref();
    return { signal: controller.signal, abort: reason => controller.abort(reason), clear: () => clearTimeout(timer) };
  }

  /** Live budget checks: output is strict, scratch only counts bytes, and one total covers both. */
  monitor(output, scratch, abort) {
    let scanning = false;
    const timer = setInterval(async () => {
      if (scanning) return;
      scanning = true;
      try {
        const published = await measure(output, { maxBytes: this.limits.maxOutputBytes, maxEntries: this.limits.maxOutputFiles });
        await measure(scratch, { maxBytes: this.limits.maxOutputBytes - published, strict: false });
      } catch (error) { if (error.code !== 'ENOENT') abort(error); }
      finally { scanning = false; }
    }, 50);
    timer.unref();
    return () => clearInterval(timer);
  }

  async authorize(peer, scope, signal, action, path, revision) {
    if (peer.signal?.aborted) fail('forbidden', 'The document peer is no longer available.');
    await this.authorizer.authorize(peer, scope, signal, { action, path, revision });
    signal.throwIfAborted();
  }

  /** A file still being written by another program is read again a couple of times before giving up. */
  open(scope, path, revision, upstream = never, peer, { mode = 'view' } = {}) {
    return this.track((async () => {
      for (let attempt = 0; ; attempt++) {
        try { return await this.openSession(scope, path, revision, upstream, peer, mode); }
        catch (error) {
          if (error?.code !== 'source-changed' || attempt >= 2 || upstream.aborted || peer?.signal?.aborted) throw error;
          await new Promise(resolve => setTimeout(resolve, 300 * (attempt + 1)));
        }
      }
    })());
  }

  async openSession(scope, path, revision, upstream, peer, mode) {
    if (!peer?.id || !scope?.sessionId || typeof scope.workspaceRoot !== 'string') fail('forbidden', 'Opening a document requires an authorized Session.');
    if (!Number.isSafeInteger(revision) || revision < 0) fail('invalid-revision', 'Revision must be a nonnegative safe integer.');
    if (mode !== 'view' && mode !== 'edit') fail('unsupported-operation', 'Mode must be view or edit.');
    const format = formatOf(path);
    const caps = this.capabilities(format);
    if (!caps.preview || (mode === 'edit' && !caps.edit)) fail('unsupported-operation', `This document cannot be opened for ${mode === 'edit' ? 'editing' : 'preview'}.`);
    await this.start();
    const deadline = this.deadline('import', format);
    const signal = anySignal([upstream, peer.signal, deadline.signal, this.lifetime.signal]);
    let directory, acquired = false, release = () => {}, published = false, stopMonitor = () => {};
    try {
      return await guarded(signal, async () => {
        await this.acquire(signal); acquired = true;
        await this.expire();
        if (this.sessions.size >= this.limits.maxLeases) fail('busy', 'The document session limit is reached.');
        const action = mode === 'edit' ? 'edit' : 'open';
        await this.authorize(peer, scope, signal, action, path, revision);
        const source = await this.source.stat(scope, path, signal);
        const parked = this.parkedFor(scope, source, peer);
        if (parked) {
          if (sameVersion(source, { version: parked.baseVersion, absolutePath: parked.absolutePath })) return this.resume(parked, peer, revision, await this.writeAccess(parked));
          // The file moved on while the edits waited: keep them as a draft and open the new file.
          await this.closeSession(parked, { keepDraft: true });
        }
        if (source.bytes > this.limits.maxInputBytes) fail('input-too-large', 'The document exceeds the input limit.', { limit: 'maxInputBytes' });
        release = this.reserve(this.limits.maxInputBytes + this.limits.maxOutputBytes);
        const input = await this.source.readBytes(scope, path, {}, signal);
        if (!sameVersion(source, input)) fail('source-changed', 'The document changed while it was read.');
        if (!(input.data instanceof Uint8Array) || input.data.byteLength > this.limits.maxInputBytes || input.eof === false) fail('input-too-large', 'The document exceeds the input limit.', { limit: 'maxInputBytes' });
        directory = await mkdtemp(join(this.tempRoot, TASK_PREFIX));
        const inputPath = join(directory, 'input', `source.${format}`);
        const output = join(directory, 'output'), scratch = join(directory, 'scratch');
        for (const dir of ['input', 'output', 'scratch']) await mkdir(join(directory, dir), { mode: 0o700 });
        await writeFile(inputPath, input.data, { flag: 'wx', mode: 0o600, signal });
        stopMonitor = this.monitor(output, scratch, deadline.abort);
        const result = await this.converter.import({ format, input: inputPath, output, scratch }, signal);
        stopMonitor();
        signal.throwIfAborted();
        const snapshot = await this.validateImport(format, output, result);
        await rm(join(directory, 'input'), { recursive: true, force: true });
        await rm(scratch, { recursive: true, force: true });
        const bytes = await measure(output, { maxBytes: this.limits.maxOutputBytes, maxEntries: this.limits.maxOutputFiles });
        await this.authorize(peer, scope, signal, action, path, revision);
        if (!sameVersion(source, await this.source.stat(scope, path, signal))) fail('source-changed', 'The document changed during conversion.');
        if (this.sessions.size >= this.limits.maxLeases) fail('busy', 'The document session limit is reached.');
        const id = token(24);
        const session = {
          id, scope: { ...scope }, scopeKey: scopeKey(scope), peer, principal: peer.principal ?? 'owner', path, name: basename(source.absolutePath || path),
          absolutePath: source.absolutePath, baseVersion: source.version, revision, format, mode, directory,
          snapshot: { directory: output, ...snapshot }, warnings: result.warnings ?? [], component: result.component,
          bytes, editRevision: 0, savedEditRevision: 0, state: 'active', expires: this.now() + this.limits.leaseTtlMs,
          uses: new Set(), closed: false, queue: Promise.resolve(), generation: 0,
        };
        this.sessions.set(id, session);
        this.retainedBytes += bytes;
        this.attach(session, peer);
        published = true;
        return this.handle(session);
      });
    } finally {
      deadline.clear(); stopMonitor(); release();
      if (directory && !published) await rm(directory, { recursive: true, force: true });
      if (acquired) this.relinquish();
    }
  }

  async validateImport(format, output, result) {
    if (result?.format !== format || typeof result.content !== 'string' || !Array.isArray(result.assets) || result.assets.length > 32) fail('invalid-output', 'Invalid converter result.');
    safeRelative(result.content);
    await confinedFile(output, result.content);
    for (const asset of result.assets) {
      safeRelative(asset);
      const info = await lstat(join(output, asset));
      if (info.isSymbolicLink() || !info.isDirectory()) fail('invalid-output', 'Invalid converter asset directory.');
    }
    return { content: result.content, assets: [...result.assets], snapshot: FORMATS[format].snapshot };
  }

  /** A peer that goes away leaves unsaved edits parked for the next open; a clean session just closes. */
  attach(session, peer) {
    session.peer = peer;
    session.peerAbort = () => { this.leave(session).catch(() => {}); };
    peer.signal?.addEventListener('abort', session.peerAbort, { once: true });
  }

  dirty(session) { return session.mode === 'edit' && session.editRevision !== session.savedEditRevision; }

  parkedFor(scope, source, peer) {
    const key = scopeKey(scope), principal = peer.principal ?? 'owner';
    for (const session of this.sessions.values()) {
      if (session.parked && !session.closed && session.scopeKey === key && session.principal === principal && session.absolutePath === source.absolutePath) return session;
    }
  }

  /** Autosave could not write a file outside the workspace: the panel says the edits wait for a confirmed write-back. */
  resume(session, peer, revision, access) {
    session.parked = false;
    session.revision = revision;
    session.expires = this.now() + this.limits.leaseTtlMs;
    this.attach(session, peer);
    return { ...this.handle(session), ...access === 'outside' && this.dirty(session) && { saveNeedsApproval: true } };
  }

  /**
   * The tab (or its peer) is gone. Unsaved edits are saved when the file may be written without
   * asking; otherwise the session stays parked in memory until it is reopened, saved or expires.
   */
  async leave(session) {
    if (session.closed || session.parked) return;
    if (!this.dirty(session)) return this.closeSession(session);
    session.peer.signal?.removeEventListener('abort', session.peerAbort);
    session.parked = true;
    session.peer = { id: undefined, principal: session.principal };
    session.expires = this.now() + this.limits.parkTtlMs;
    try {
      if (!(await this.autosaves(session))) return;
      const deadline = AbortSignal.any([AbortSignal.timeout(this.limits.timeoutMs.export[session.format] ?? 180_000), this.lifetime.signal]);
      await this.saveSession(session, deadline, { system: true });
      if (session.parked && !this.dirty(session)) await this.closeSession(session);
    } catch { /* stays parked: the next open resumes it, expiry keeps it as a draft */ }
  }

  /**
   * Saving without a peer skips `authorize`, so only the owner autosaves: a remote principal may be
   * allowed to edit yet denied `save`, and its edits wait as a parked session, then a draft.
   */
  async autosaves(session) {
    return session.principal === 'owner' && this.canWrite(session);
  }

  /** 'allowed' | 'outside' (one user approval writes it) | 'denied'; see the writer's `writeAccess`. */
  async writeAccess(session) {
    if (!this.writer || this.hostCapabilities.save === false) return 'denied';
    if (typeof this.writer.writeAccess !== 'function') return 'allowed';
    return this.writer.writeAccess(session.scope, session.path).catch(() => 'denied');
  }

  async canWrite(session) {
    return (await this.writeAccess(session)) === 'allowed';
  }

  /** The open (or parked) session for a file in this scope, by its absolute path. */
  async find(scope, path, peer, signal) {
    const source = await this.source.stat(scope, path, signal);
    const key = scopeKey(scope), principal = peer?.principal ?? 'owner';
    for (const session of this.sessions.values()) {
      if (!session.closed && session.scopeKey === key && session.principal === principal && session.absolutePath === source.absolutePath) return session;
    }
  }

  handle(session) {
    const prefix = `${this.routePrefix}/doc/${session.id}`;
    return { id: session.id, format: session.format, mode: session.mode, snapshot: session.snapshot.snapshot, version: session.baseVersion,
      revision: session.revision, editRevision: session.editRevision, unsaved: this.dirty(session), capabilities: this.capabilities(session.format),
      expiresInMs: this.limits.leaseTtlMs, entry: `${prefix}/index.html`, warnings: session.warnings };
  }

  meta(session) {
    const prefix = `${this.routePrefix}/doc/${session.id}`;
    return { v: 1, id: session.id, name: session.name, format: session.format, mode: session.mode, snapshot: session.snapshot.snapshot,
      version: session.baseVersion, revision: session.revision, editRevision: session.editRevision, capabilities: this.capabilities(session.format),
      routePrefix: this.routePrefix, contentUrl: `${prefix}/content`, resourcesBaseUrl: `${prefix}/resources/`, runtimeBaseUrl: `${this.routePrefix}/runtime/`,
      sdkApiBaseUrl: `${prefix}/sdk/v2/api`, stageUrl: session.mode === 'edit' ? `${prefix}/stage/` : undefined, assetRoots: session.snapshot.assets };
  }

  async session(id, peer, signal) {
    signal.throwIfAborted();
    const session = this.sessions.get(id);
    if (!session || session.closed || session.parked || session.peer.id !== peer?.id || peer.signal?.aborted || session.expires <= this.now()) {
      if (session && !session.parked && session.expires <= this.now()) await this.leave(session);
      fail('lease-unavailable', 'The document session is unavailable.');
    }
    return session;
  }

  /** A diverged session still re-authorizes, but no longer compares against a source the user chose to leave behind. */
  async check(session, signal) {
    await this.authorize(session.peer, session.scope, signal, session.mode === 'edit' ? 'edit' : 'open', session.path, session.revision);
    if (session.state === 'diverged') return;
    const current = await this.source.readBytes(session.scope, session.path, { range: { offset: 0, length: 1 } }, signal);
    const expected = { version: session.baseVersion, absolutePath: session.absolutePath };
    if (!sameVersion(expected, current) || !sameVersion(expected, await this.source.stat(session.scope, session.path, signal))) fail('source-changed', 'The document changed.');
  }

  /** Readers wait for no writer, but closing waits for every reader. */
  async use(session, operation) {
    let finish;
    const done = new Promise(resolve => { finish = resolve; });
    session.uses.add(done);
    try {
      if (session.closed) fail('lease-unavailable', 'The document session is unavailable.');
      return await operation();
    } finally { session.uses.delete(done); finish(); }
  }

  /** Writers on one session run one at a time so revisions cannot interleave. */
  exclusive(session, operation) {
    const run = session.queue.then(() => this.use(session, operation));
    session.queue = run.catch(() => {});
    return run;
  }

  async resource(id, route, peer, upstream = never) {
    safeRelative(route);
    const signal = anySignal([upstream, peer?.signal, this.lifetime.signal]);
    return guarded(signal, async () => {
      const session = await this.session(id, peer, signal);
      return this.readers.read(signal, () => this.use(session, async () => {
        if (route.startsWith('export/')) return this.download(route.slice('export/'.length), peer, session.id);
        await this.check(session, signal);
        if (route === 'meta') return { body: Buffer.from(JSON.stringify(this.meta(session))), type: 'application/json; charset=utf-8' };
        if (route === 'index.html') return { body: await readConfined(this.runtimeRoot, 'preview.html', 1024 * 1024, signal), type: 'text/html; charset=utf-8' };
        if (route === 'content') {
          const body = await readConfined(session.snapshot.directory, session.snapshot.content, this.limits.maxStageBytes, signal);
          return { body, type: contentType(session.format, session.snapshot.content),
            headers: { 'x-amz-meta-bytes': String(body.length), 'x-amz-meta-head': '1', 'x-amz-meta-length': String(body.toString('utf8').length), 'x-amz-meta-delta-version': '1' } };
        }
        if (route.startsWith('resources/')) return this.snapshotResource(session, safeRelative(route.slice('resources/'.length)), signal);
        fail('not-found', 'Route not found.');
      }));
    });
  }

  /** Direct asset-root paths first, then the category aliases SDKs use for embedded media. */
  async snapshotResource(session, resource, signal) {
    const candidates = [];
    if (session.snapshot.assets.some(root => resource.startsWith(root + '/'))) candidates.push(resource);
    const [category, ...rest] = resource.split('/');
    const mapped = { images: 'image', image: 'image', embedding: 'embedding', media: 'media', embeddings: 'embeddings' }[category];
    // Roots may be the category itself (workbook.modoc/media, media) or contain it (result.bin.dir/image).
    if (mapped && rest.length) for (const root of session.snapshot.assets) {
      candidates.push(root.split('/').at(-1) === mapped ? `${root}/${rest.join('/')}` : `${root}/${mapped}/${rest.join('/')}`);
    }
    for (const candidate of candidates) {
      try { return { body: await readConfined(session.snapshot.directory, candidate, this.limits.maxOutputBytes, signal), type: mime(candidate) }; }
      catch (error) { if (error.code !== 'ENOENT') throw error; }
    }
    fail('not-found', 'Resource not found.');
  }

  async renew(scope, id, upstream = never, peer) {
    const signal = anySignal([upstream, peer?.signal, this.lifetime.signal]);
    return guarded(signal, async () => {
      const session = await this.session(id, peer, signal);
      if (session.scopeKey !== scopeKey(scope)) fail('lease-unavailable', 'The document session is unavailable.');
      await this.use(session, () => this.check(session, signal));
      session.expires = this.now() + this.limits.leaseTtlMs;
      return { expiresInMs: this.limits.leaseTtlMs };
    });
  }

  /**
   * Harness bumps its revision on every source change, including our own save. Rule 1 keeps
   * the session when the file is what we last wrote; rule 3 keeps a dirty session but marks it diverged.
   */
  async rebind(scope, id, revision, { diverged = false } = {}, upstream = never, peer) {
    if (!Number.isSafeInteger(revision) || revision < 0) fail('invalid-revision', 'Revision must be a nonnegative safe integer.');
    const signal = anySignal([upstream, peer?.signal, this.lifetime.signal]);
    return guarded(signal, async () => {
      const session = await this.session(id, peer, signal);
      if (session.scopeKey !== scopeKey(scope)) fail('lease-unavailable', 'The document session is unavailable.');
      return this.exclusive(session, async () => {
        if (diverged) {
          if (session.mode !== 'edit' || session.editRevision === session.savedEditRevision) fail('unsupported-operation', 'Only a session with unsaved edits can diverge.');
          await this.authorize(peer, session.scope, signal, 'edit', session.path, revision);
          session.state = 'diverged';
        } else await this.check(session, signal);
        session.revision = revision;
        session.expires = this.now() + this.limits.leaseTtlMs;
        return { revision, state: session.state, version: session.baseVersion };
      });
    });
  }

  async stage(id, part, body, { baseEditRevision } = {}, peer, upstream = never) {
    if (typeof part !== 'string' || !part) fail('invalid-path', 'Invalid document part.');
    if (!(body instanceof Uint8Array)) fail('invalid-body', 'Document parts must be bytes.');
    if (body.byteLength > this.limits.maxStageBytes) fail('input-too-large', 'The document part exceeds the stage limit.', { limit: 'maxStageBytes' });
    const signal = anySignal([upstream, peer?.signal, this.lifetime.signal]);
    return guarded(signal, async () => {
      const session = await this.session(id, peer, signal);
      if (session.mode !== 'edit') fail('read-only', 'This document session is read only.');
      return this.exclusive(session, async () => {
        if (baseEditRevision !== session.editRevision) fail('edit-conflict', 'The document was edited elsewhere in this session.', { editRevision: session.editRevision });
        await this.authorize(peer, session.scope, signal, 'edit', session.path, session.revision);
        const release = this.reserve(body.byteLength);
        try {
          const change = await EDITS[session.format].package(session.snapshot, part, Buffer.from(body.buffer, body.byteOffset, body.byteLength), this.limits);
          await this.applyChange(session, change);
        } finally { release(); }
        session.editRevision++;
        session.expires = this.now() + this.limits.leaseTtlMs;
        return { editRevision: session.editRevision };
      });
    });
  }

  /** Full replacements are built beside the snapshot and swapped in; single files are renamed into place. */
  async applyChange(session, change) {
    const root = session.snapshot.directory;
    if (change.replace) {
      const next = join(session.directory, `snapshot-${++session.generation}`);
      await mkdir(next, { mode: 0o700 });
      try {
        for (const [name, data] of change.files) {
          await mkdir(dirname(join(next, name)), { recursive: true, mode: 0o700 });
          await writeFile(join(next, name), data, { flag: 'wx', mode: 0o600 });
        }
      } catch (error) { await rm(next, { recursive: true, force: true }); throw error; }
      const previous = root;
      session.snapshot = { ...session.snapshot, directory: next, content: change.content, assets: change.assets };
      await rm(previous, { recursive: true, force: true });
    } else {
      for (const [name, data] of change.files) {
        safeRelative(name);
        const target = join(root, name);
        await mkdir(dirname(target), { recursive: true, mode: 0o700 });
        const temporary = `${target}.${token(6)}.tmp`;
        await writeFile(temporary, data, { flag: 'wx', mode: 0o600 });
        await rename(temporary, target);
      }
    }
    const bytes = await measure(session.snapshot.directory, { maxBytes: this.limits.maxStageBytes, maxEntries: this.limits.maxOutputFiles });
    this.retainedBytes += bytes - session.bytes;
    session.bytes = bytes;
  }

  /** Converts a snapshot back to its Office format inside a fresh scratch directory. */
  async exportOffice(format, snapshot, upstream) {
    const deadline = this.deadline('export', format);
    const signal = anySignal([upstream, deadline.signal, this.lifetime.signal]);
    let acquired = false, directory, release = () => {}, stopMonitor = () => {};
    try {
      return await guarded(signal, async () => {
        await this.acquire(signal); acquired = true;
        release = this.reserve(this.limits.maxInputBytes);
        directory = await mkdtemp(join(this.tempRoot, TASK_PREFIX));
        const scratch = join(directory, 'scratch'), output = join(directory, 'output');
        await mkdir(scratch, { mode: 0o700 }); await mkdir(output, { mode: 0o700 });
        const prepared = await EDITS[format].prepareExport(snapshot, scratch);
        const file = join(output, `document.${format}`);
        stopMonitor = this.monitor(output, scratch, deadline.abort);
        await this.converter.export({ format, snapshot: prepared, content: snapshot.content, output: file, scratch }, signal);
        stopMonitor();
        const info = await lstat(file).catch(() => fail('invalid-output', 'The converter did not write the exported file.'));
        if (!info.isFile()) fail('invalid-output', 'The exported file is not a regular file.');
        if (info.size > this.limits.maxInputBytes) fail('output-too-large', 'The exported file would exceed the input limit and could not be reopened.', { limit: 'maxInputBytes' });
        const data = await readFile(file);
        checkOfficePackage(data, FORMATS[format].mainPart);
        return data;
      });
    } finally {
      deadline.clear(); stopMonitor(); release();
      if (directory) await rm(directory, { recursive: true, force: true });
      if (acquired) this.relinquish();
    }
  }

  owner(peer, session) { return { principal: peer.principal ?? 'owner', sessionId: session.scope.sessionId, path: session.absolutePath ?? session.path }; }

  /** At most one draft per edit revision: shutdown, a failed save and a later close reuse it. */
  async keepDraft(session, kind, data) {
    if (session.draft?.editRevision === session.editRevision) {
      // Drafts are evicted by count, bytes and age: reuse only one that is still there.
      if (await this.drafts.get(this.owner(session.peer, session), session.draft.id).then(() => true, () => false)) return session.draft;
      session.draft = undefined;
    }
    const temporary = await mkdtemp(join(this.tempRoot, TASK_PREFIX));
    try {
      let source = session.snapshot.directory;
      if (kind === 'office') { source = join(temporary, `document.${session.format}`); await writeFile(source, data, { mode: 0o600 }); }
      const draft = await this.drafts.add({ owner: this.owner(session.peer, session), format: session.format, kind, source, name: session.name,
        sourceVersion: session.baseVersion, content: session.snapshot.content, assets: session.snapshot.assets });
      session.draft = { id: draft.id, editRevision: session.editRevision };
      return draft;
    } finally { await rm(temporary, { recursive: true, force: true }); }
  }

  async save(scope, id, { editRevision, approved = false } = {}, upstream = never, peer) {
    const signal = anySignal([upstream, peer?.signal, this.lifetime.signal]);
    return guarded(signal, async () => {
      const session = await this.session(id, peer, signal);
      if (session.scopeKey !== scopeKey(scope)) fail('lease-unavailable', 'The document session is unavailable.');
      return this.saveSession(session, signal, { editRevision, approved, peer });
    });
  }

  /**
   * Converts the staged snapshot and writes it over the source. `approved` carries this one user
   * consent for a target outside the writable roots. `system` saves have no peer to authorize (tab gone,
   * shutdown, a parked session saved by the AI after the Session check).
   */
  async saveSession(session, signal, { editRevision = session.editRevision, approved = false, peer = session.peer, system = false } = {}) {
    if (session.mode !== 'edit' || !this.capabilities(session.format).save) fail('unsupported-operation', 'This document cannot be saved here.');
    return guarded(signal, () => this.exclusive(session, async () => {
        if (editRevision !== session.editRevision) fail('edit-conflict', 'Unsent edits are pending.', { editRevision: session.editRevision });
        if (!system) await this.authorize(peer, session.scope, signal, 'save', session.path, session.revision);
        const data = await this.exportOffice(session.format, session.snapshot, signal);
        if (session.state === 'diverged') {
          const draft = await this.keepDraft(session, 'office', data);
          fail('source-changed', 'The document changed outside this editor; your edits were kept as a draft.', { draftId: draft.id });
        }
        let written;
        try { written = await this.writer.replace(session.scope, session.path, data, { expectedVersion: session.baseVersion, approved }, signal); }
        catch (error) {
          if (signal.aborted) throw error;
          // Outside the writable roots nothing is lost: the edits stay in this session until the user agrees.
          if (error?.code === 'sandbox-denied') { error.details = { ...error.details, outside: true }; throw error; }
          const draft = await this.keepDraft(session, 'office', data);
          if (error instanceof SuiteError) { error.details = { ...error.details, draftId: draft.id }; throw error; }
          fail('unavailable', 'The document could not be written; your edits were kept as a draft.', { draftId: draft.id });
        }
        session.baseVersion = written.version;
        if (written.absolutePath) session.absolutePath = written.absolutePath;
        session.savedEditRevision = session.editRevision;
        return { version: written.version, editRevision: session.editRevision, path: session.absolutePath };
    }));
  }

  async exportTo(scope, id, { editRevision, target } = {}, upstream = never, peer) {
    const signal = anySignal([upstream, peer?.signal, this.lifetime.signal]);
    return guarded(signal, async () => {
      const session = await this.session(id, peer, signal);
      if (session.scopeKey !== scopeKey(scope)) fail('lease-unavailable', 'The document session is unavailable.');
      if (!this.capabilities(session.format).export) fail('unsupported-operation', 'This document cannot be exported.');
      return this.exclusive(session, async () => {
        if (editRevision !== session.editRevision) fail('edit-conflict', 'Unsent edits are pending.', { editRevision: session.editRevision });
        await this.authorize(peer, session.scope, signal, 'export', session.path, session.revision);
        const data = await this.exportOffice(session.format, session.snapshot, signal);
        return this.deliver(data, session.format, session.name, target, peer, session.scope, signal, `doc/${session.id}`, session.id);
      });
    });
  }

  async deliver(data, format, name, target, peer, scope, signal, routeBase, sessionId) {
    if (target === 'download') {
      const id = token(24);
      const directory = await mkdtemp(join(this.tempRoot, TASK_PREFIX));
      const file = join(directory, `document.${format}`);
      await writeFile(file, data, { mode: 0o600 });
      const fileName = name.toLowerCase().endsWith(`.${format}`) ? name : `${name}.${format}`;
      const entry = { id, directory, file, name: fileName, type: mime(file), peerId: peer.id, sessionId, expires: this.now() + this.limits.exportTtlMs };
      this.downloads.set(id, entry);
      setTimeout(() => { this.dropDownload(id).catch(() => {}); }, this.limits.exportTtlMs).unref();
      return { url: `${this.routePrefix}/${routeBase}/export/${id}`, name: fileName, bytes: data.length };
    }
    if (target && typeof target === 'object' && typeof target.path === 'string') {
      if (!this.writer || this.hostCapabilities.save === false) fail('unsupported-operation', 'This host cannot write files.');
      if (formatOf(target.path) !== format) fail('unsupported-format', `Save as ${format.toUpperCase()} only.`);
      await this.authorize(peer, scope, signal, 'export', target.path, 0);
      return await this.writer.create(scope, target.path, data, signal);
    }
    fail('invalid-body', 'Export target must be download or { path }.');
  }

  async download(id, peer, sessionId) {
    const entry = this.downloads.get(id);
    if (!entry || entry.peerId !== peer?.id || entry.sessionId !== sessionId || entry.expires <= this.now()) fail('not-found', 'Export not found.');
    this.downloads.delete(id);
    try {
      const body = await readFile(entry.file);
      return { body, type: entry.type, headers: { 'Content-Disposition': `attachment; filename*=UTF-8''${encodeURIComponent(entry.name)}` } };
    } finally { await rm(entry.directory, { recursive: true, force: true }); }
  }
  async dropDownload(id) {
    const entry = this.downloads.get(id);
    if (!entry) return;
    this.downloads.delete(id);
    await rm(entry.directory, { recursive: true, force: true });
  }

  draftOwner(scope, peer) { return { principal: peer?.principal ?? 'owner', sessionId: scope.sessionId }; }

  async listDrafts(scope, upstream = never, peer) {
    const signal = anySignal([upstream, peer?.signal, this.lifetime.signal]);
    return guarded(signal, async () => {
      if (!peer?.id || !scope?.sessionId) fail('forbidden', 'Drafts require an authorized Session.');
      await this.authorize(peer, scope, signal, 'export', '', 0);
      return this.drafts.list(this.draftOwner(scope, peer));
    });
  }

  async discardDraft(scope, draftId, upstream = never, peer) {
    const signal = anySignal([upstream, peer?.signal, this.lifetime.signal]);
    return guarded(signal, async () => {
      await this.authorize(peer, scope, signal, 'export', '', 0);
      await this.drafts.discard(this.draftOwner(scope, peer), draftId);
    });
  }

  /** Snapshot drafts convert lazily here; office drafts are delivered as written. */
  async exportDraft(scope, draftId, { target } = {}, upstream = never, peer) {
    const signal = anySignal([upstream, peer?.signal, this.lifetime.signal]);
    return guarded(signal, async () => {
      await this.authorize(peer, scope, signal, 'export', '', 0);
      const draft = await this.drafts.get(this.draftOwner(scope, peer), draftId);
      const data = draft.kind === 'office'
        ? await readFile(join(draft.directory, 'data'))
        : await this.exportOffice(draft.format, { directory: join(draft.directory, 'data'), content: draft.content, assets: draft.assets }, signal);
      return this.deliver(data, draft.format, draft.name, target, peer, scope, signal, `draft/${draft.id}`, `draft:${draft.id}`);
    });
  }

  async draftResource(draftId, route, peer, upstream = never) {
    const signal = anySignal([upstream, peer?.signal, this.lifetime.signal]);
    return guarded(signal, async () => {
      if (!route.startsWith('export/')) fail('not-found', 'Route not found.');
      return this.download(safeRelative(route.slice('export/'.length)), peer, `draft:${draftId}`);
    });
  }

  async release(scope, id, upstream = never, peer) {
    upstream.throwIfAborted();
    const session = this.sessions.get(id);
    if (!session) return;
    if (session.parked) return;
    if (session.peer.id !== peer?.id || session.scopeKey !== scopeKey(scope)) fail('lease-unavailable', 'The document session is unavailable.');
    await this.leave(session);
  }

  /** Unsaved edits leave as a raw snapshot draft: closing never waits for a conversion. */
  async closeSession(session, { keepDraft = false } = {}) {
    if (session.closing) return await session.closing;
    session.closed = true;
    this.sessions.delete(session.id);
    session.peer.signal?.removeEventListener('abort', session.peerAbort);
    session.closing = (async () => {
      await Promise.allSettled([session.queue, ...session.uses]);
      if (keepDraft && this.dirty(session)) {
        await this.keepDraft(session, 'snapshot').catch(() => {});
      }
      await rm(session.directory, { recursive: true, force: true });
      this.retainedBytes -= session.bytes;
      for (const [id, entry] of this.downloads) if (entry.sessionId === session.id) await this.dropDownload(id);
    })();
    this.closing.add(session.closing);
    try { return await session.closing; } finally { this.closing.delete(session.closing); }
  }

  async expire() {
    await Promise.all([...this.sessions.values()].filter(s => s.expires <= this.now())
      .map(s => s.parked ? this.closeSession(s, { keepDraft: true }) : this.leave(s)));
  }

  async dispose() {
    clearInterval(this.sweeper);
    await Promise.allSettled([...this.sessions.values()].filter(session => this.dirty(session)).map(async session => {
      const draft = await this.keepDraft(session, 'snapshot');
      if (!(await this.autosaves(session))) return;
      await this.saveSession(session, AbortSignal.timeout(this.limits.autosaveMs), { system: true });
      await this.drafts.discard(this.owner(session.peer, session), draft.id).catch(() => {});
      session.draft = undefined;
    }));
    this.lifetime.abort(new SuiteError('cancelled', 'The document service is shutting down.'));
    await Promise.allSettled([...this.operations]);
    await Promise.all([...this.sessions.values()].map(session => this.closeSession(session, { keepDraft: true })));
    await Promise.all([...this.closing]);
    for (const id of [...this.downloads.keys()]) await this.dropDownload(id);
  }

  async runtimeResource(relative, signal = never) {
    return { body: await readConfined(this.runtimeRoot, relative, 32 * 1024 * 1024, signal), type: mime(relative) };
  }
}

export const createSuiteService = options => new SuiteService(options);
