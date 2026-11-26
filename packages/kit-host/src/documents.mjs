import { randomUUID } from 'node:crypto';
import { existsSync, mkdirSync, rmSync } from 'node:fs';
import { open as openFile, readdir, realpath, rename, rm, stat as statPath } from 'node:fs/promises';
import * as nodePath from 'node:path';

/**
 * Host Adapter documents over one Kit (spec §12.3). Kit holds the document facts; this layer decides who
 * may do what before any command reaches Kit: Session authorization, the writer's write access on the
 * real path, one writer view per document, a gap-free change queue and the pre-save barrier.
 *
 * Injected: `kit` (KitProcess), `authorize(peer, scope, signal, { action, path })`, `stat(scope, path, signal)`
 * → `{ absolutePath }` (the Session's own file view), `writeAccess(scope, realPath)` → 'allowed' | 'outside'
 * | 'denied', `exportRoot` (Host-private directory for downloads). `platform` and `rename` default to this process's;
 * tests replace them to act out Windows.
 */
export class KitError extends Error {
  constructor(code, message, details) { super(message); this.name = 'KitError'; this.code = code; if (details) this.details = details; }
}
const fail = (code, message, details) => { throw new KitError(code, message, details); };

/** Kit failures the Host maps to a shared code; anything else stays opaque behind `kitCode`. */
const KIT_CODES = {
  'file-changed': ['source-changed', 'The file changed on disk since it was opened.', { overwrite: true }],
  'target-in-use': ['exists', 'Another open document is bound to that file.'],
  'need-password': ['password-protected', 'Password-protected files are not supported yet.'],
  'wrong-password': ['password-protected', 'Password-protected files are not supported yet.'],
  'path-missing': ['not-found', 'The file no longer exists.'],
  'unsupported-format': ['unsupported-format', 'That file format cannot be written; the edits are kept.'],
  'protection-unsupported': ['unsupported-operation', 'This document\'s editing restrictions cannot be written back offline; the edits are kept.'],
  'media-kind-rejected': ['invalid-body', 'The document service refused that kind of media.'],
  'change-body-rejected': ['invalid-body', 'The document service refused an edit.'],
};
function kitFailure(frame, what) {
  if (frame.unknown) return new KitError('unavailable', `The document service stopped while ${what}; the document is being reopened.`, { rebuild: true });
  const known = KIT_CODES[frame.code];
  if (known) return new KitError(known[0], known[1], { kitCode: frame.code, ...known[2] });
  if (frame.code === 'kit-unavailable' || frame.code === 'cancelled') return new KitError('unavailable', 'The document service is not available right now.', { kitCode: frame.code });
  return new KitError('unavailable', `The document service could not finish ${what}.`, { kitCode: frame.code ?? 'unknown' });
}

/** The reference Host never lets these reach the page (kit-usage/05-host.md). */
const PRIVATE_KEYS = new Set(['workdir', 'workingdirectory', 'modocpath', 'cachedir', 'cachepath', 'contentpath', 'directory', 'password', 'sessionpassword']);
/** Per suite (spec §13): the file extension it opens, its change format and the media kinds a write carries. */
const SUITES = {
  document: { extension: '.docx', format: 'mow-v1', kinds: ['image', 'embedding'], type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' },
  sheet: { extension: '.xlsx', format: 'modoc-v1', kinds: [undefined], type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' },
  presentation: { extension: '.pptx', format: 'mop-v1', kinds: ['media', 'embeddings'], type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation' },
};
export const suiteOf = filePath => Object.keys(SUITES).find(suite => filePath.toLowerCase().endsWith(SUITES[suite].extension));
export const contentTypeOf = suite => SUITES[suite]?.type ?? 'application/octet-stream';
/** A workspace-relative path as the Client takes it: '/'-separated on every platform (Windows gives '\'). */
export const slashed = (relative, sep = nodePath.sep) => (sep === '/' ? relative : relative.split(sep).join('/'));

export class KitDocuments {
  views = new Map();
  downloads = new Map();
  /**
   * Documents whose Kit record differs from the file only by the editor's own load-time changes (`inert`): Kit
   * calls them unsaved, but nothing was edited, so they are neither marked unsaved nor written back. In memory only;
   * closing the last view puts the disk copy back into Kit (`import`), so a later open or Kit restart starts clean.
   */
  #inert = new Set();

  constructor({ kit, authorize, stat, writeAccess, exportRoot, now = () => Date.now(), viewTtlMs = 15 * 60_000, warn, platform = process.platform, rename: renameFile = rename }) {
    Object.assign(this, { kit, stat, writeAccess, exportRoot, now, viewTtlMs, warn, platform, renameFile });
    this.authorize = (peer, scope, signal, request) => authorize(peer, scope, signal ?? peer?.signal ?? new AbortController().signal, request);
    // A lost channel makes every view of that Kit stale: the page reopens and rebuilds its editor.
    kit.on?.('down', ({ generation }) => { for (const view of this.views.values()) if (view.generation <= generation) view.stale = true; });
  }

  async #client() {
    try { return await this.kit.ready(); }
    catch (error) {
      // Errors that say what to do (an update in progress, the document engine download or its checks) keep their words.
      const explained = ['kit-suspended', 'download-failed', 'denied', 'kit-selftest', 'kit-downgrade', 'still-downloading', 'unsupported-platform'].includes(error?.code);
      fail('unavailable', explained ? error.message : 'The document service could not start.', { cause: error?.message, ...explained && error.code !== 'kit-suspended' && { engine: error.code } });
    }
  }

  /** Every command to Kit for a document goes through its view; a stale view never reaches Kit. */
  view(viewId, peer) {
    const view = this.views.get(viewId);
    if (!view || view.closed) fail('lease-unavailable', 'The document view is no longer open.');
    if (peer && view.peerId !== peer.id) fail('lease-unavailable', 'The document view belongs to another tab.');
    if (view.stale) fail('edit-conflict', 'The document service restarted; reopen the document.', { rebuild: true });
    view.expires = this.now() + this.viewTtlMs;
    return view;
  }

  #writerOf(documentId) {
    for (const view of this.views.values()) if (!view.closed && view.documentId === documentId && view.writer && !view.stale) return view;
    return undefined;
  }

  /**
   * Opens `path` of the Session for one tab. The first edit-mode view of a document is its writer; any
   * other view — another tab, conversation or principal — is read only.
   */
  async open(scope, filePath, peer, { mode = 'edit', signal } = {}) {
    await this.authorize(peer, scope, signal, { action: mode === 'edit' ? 'edit' : 'open', path: filePath });
    const { absolutePath, version } = await this.stat(scope, filePath, signal);
    const real = await realpath(absolutePath);
    const client = await this.#client();
    const generation = this.kit.generation;
    const opened = await client.request('open', { path: real }, { signal });
    if (opened.code === 'unsupported-format') fail('unsupported-format', 'The document service cannot open this file format.', { kitCode: opened.code });
    if (!opened.ok) throw kitFailure(opened, 'opening the file');
    for (const key of Object.keys(opened)) if (PRIVATE_KEYS.has(key.toLowerCase())) fail('unavailable', 'The document service returned a private field.');
    // A record Kit opened as another suite gets no view; it is never discarded (that would drop unsaved edits).
    const expected = suiteOf(real);
    if (!expected || opened.suite !== expected) fail('unsupported-format', `This file opened as ${opened.suite}, not as ${expected ?? 'a supported document'}.`);
    if (typeof opened.content !== 'string' || !opened.content.trim()) fail('invalid-format', 'The document service returned no content.');
    // The same tab reopening (a reload, a disk change) takes over its own writer view at once instead of waiting
    // for that view's release: the old view closes without autosaving; its edits stay in Kit's record.
    const previous = this.#writerOf(opened.documentId);
    if (mode === 'edit' && previous?.peerId === peer.id) { previous.closed = true; this.views.delete(previous.id); }
    const writer = mode === 'edit' && !this.#writerOf(opened.documentId);
    const view = { id: randomUUID(), peerId: peer.id, principal: peer.principal ?? 'owner', peer, scope, path: filePath, realPath: real,
      documentId: opened.documentId, suite: opened.suite, writer, generation, nextSeq: 0, accepted: 0, queue: Promise.resolve(),
      // Kit restarted between the request and now: the document id may belong to the old process.
      stale: this.kit.generation !== generation || Boolean(client.broken) || (this.kit.client !== undefined && this.kit.client !== client),
      fault: null, unknown: false, holds: new Set(), aiActive: 0, diskChanged: opened.diskChanged === true, unsaved: this.#unsaved(opened.documentId, opened.unsaved),
      // The Session's version of the file as opened (and after each save): the tab tells its own writes from others'.
      version, expires: this.now() + this.viewTtlMs, initial: { content: opened.content, nextChangeSetRevision: opened.nextChangeSetRevision } };
    this.views.set(view.id, view);
    return { id: view.id, suite: view.suite, mode: writer ? 'edit' : 'view', version, content: opened.content,
      nextChangeSetRevision: opened.nextChangeSetRevision, unsaved: view.unsaved, diskChanged: view.diskChanged };
  }

  /**
   * Content for the page to mount. The first call returns what `open` read; later calls are a rebuild:
   * queued changes finish, Kit's composed content comes back with its next revision, and the change
   * sequence restarts at 0 for the new editor.
   */
  async content(viewId, peer) {
    const view = this.view(viewId, peer);
    if (view.initial) { const first = view.initial; view.initial = undefined; return first; }
    await this.whenIdle(view.id);
    if (view.unknown) fail('edit-conflict', 'The document service restarted; reopen the document.', { rebuild: true });
    const client = await this.#client();
    const read = await client.request('content', { documentId: view.documentId });
    if (!read.ok) throw kitFailure(read, 'reading the document');
    view.nextSeq = 0;
    view.fault = null;
    view.unsaved = this.#unsaved(view.documentId, read.unsaved);
    return { content: read.content, nextChangeSetRevision: read.nextChangeSetRevision };
  }

  /** The page's changes, numbered from 0 without gaps; one `change` at a time, in that order. */
  change(viewId, peer, change) {
    try { return this.#enqueue(viewId, peer, change); } catch (error) { return Promise.reject(error); }
  }

  /** The last open view of a document whose Kit record holds nothing but the editor's housekeeping. */
  #onlyInert(view) {
    return !view.unsaved && this.#inert.has(view.documentId) && !view.stale && !view.fault && !view.unknown
      && ![...this.views.values()].some(other => other !== view && !other.closed && other.documentId === view.documentId);
  }

  /** Kit's own unsaved flag, unless all Kit holds beyond the file is the editor's load-time housekeeping. */
  #unsaved(documentId, reported) { return reported === true && !this.#inert.has(documentId); }

  #enqueue(viewId, peer, { seq, format, body, inert = false }) {
    const view = this.view(viewId, peer);
    if (!view.writer) fail('read-only', 'Only the editing tab can change this document.');
    if (view.tableLock) fail('busy', 'A table is being safely applied; wait for it to finish.');
    if (view.fault) fail('edit-conflict', 'The editor must reload from the document service.', { rebuild: true });
    if (typeof body !== 'string' || !body) fail('invalid-body', 'A change must be the suite change text.');
    if (format !== SUITES[view.suite].format) fail('invalid-body', `A ${view.suite} change must be ${SUITES[view.suite].format}.`);
    if (seq !== view.nextSeq) {
      view.fault = new KitError('edit-conflict', `Change ${seq} arrived, ${view.nextSeq} was expected.`, { rebuild: true });
      throw view.fault;
    }
    view.nextSeq++;
    const run = view.queue.then(async () => {
      if (view.fault) throw view.fault;
      const client = await this.#client();
      const result = await client.request('change', { documentId: view.documentId, body });
      if (!result.ok) { view.unknown ||= Boolean(result.unknown); view.fault = kitFailure(result, 'recording the change'); view.fault.details.rebuild = true; throw view.fault; }
      // The editor's own change (no one has acted on the document yet) keeps Kit in step but edits nothing.
      if (inert) { if (!view.unsaved) this.#inert.add(view.documentId); }
      else { view.unsaved = true; view.accepted++; this.#inert.delete(view.documentId); }
      // An inserted picture is safe to save once a change that references it is in Kit.
      for (const reference of view.holds) if (body.includes(reference)) view.holds.delete(reference);
      return { accepted: seq };
    });
    view.queue = run.catch(() => {});
    return run;
  }

  /** Resolves when every change already queued for this view has its Kit result. */
  whenIdle(viewId) { return this.views.get(viewId)?.queue ?? Promise.resolve(); }

  /** Any view of the document still waiting for an inserted picture to reach Kit through a change. */
  #held(documentId) {
    for (const other of this.views.values()) if (!other.closed && other.documentId === documentId && other.holds.size) return true;
    return false;
  }

  /** A recovered page starts a new change sequence on a fresh view of the same document. */
  #ready(view, what) {
    if (view.fault) fail('edit-conflict', 'The editor must reload from the document service.', { rebuild: true });
    if (view.unknown) fail('unavailable', 'The document service result is unknown; reopen the document.', { rebuild: true });
    if (view.aiActive) fail('busy', `An AI edit is running; ${what} after it finishes.`);
    if (this.#held(view.documentId)) fail('busy', `An inserted picture is not in the document yet; ${what} in a moment.`);
  }

  /**
   * Writes the document to its own file. The target is always this view's authorized real path, so a
   * rebinding by another opener can never redirect the write; Kit still refuses a changed file without
   * `overwrite`.
   */
  async save(viewId, peer, { approved = false, overwrite = false, signal } = {}) {
    const view = this.view(viewId, peer);
    if (!view.writer) fail('read-only', 'Only the editing tab can save this document.');
    this.#ready(view, 'save');
    await this.authorize(peer, view.scope, signal, { action: 'save', path: view.path });
    await this.#mayWrite(view, view.realPath, approved && view.principal === 'owner');
    await this.whenIdle(view.id);
    this.#ready(view, 'save');
    const client = await this.#client();
    // A change Kit takes after this save frame (the page is not frozen) keeps the document unsaved.
    const accepted = view.accepted;
    // From the write until the new version is known, a disk check waits: the change it sees is this save.
    const writing = (async () => {
      const saved = await client.request('save', { documentId: view.documentId, target: view.realPath, ...(overwrite && { overwrite: true }) }, { signal });
      if (!saved.ok) { if (saved.unknown) view.unknown = true; throw kitFailure(saved, 'saving'); }
      if (view.accepted === accepted) view.unsaved = false;
      this.#inert.delete(view.documentId);
      const after = await this.stat(view.scope, view.path, signal).catch(error => { this.warn?.(`could not read the version of ${view.path} after saving: ${error?.message}`); return undefined; });
      view.version = after?.version;
    })();
    const settled = writing.catch(() => {});
    view.writing = settled;
    try { await writing; } finally { if (view.writing === settled) view.writing = undefined; }
    return { saved: true, path: view.realPath, ...(view.version !== undefined && { version: view.version }) };
  }

  async #mayWrite(view, target, approved) {
    const access = await this.writeAccess(view.scope, target);
    if (access === 'denied') fail('sandbox-denied', 'This Session cannot write that file (read only).');
    if (access === 'outside' && !approved) fail('sandbox-denied', 'The file is outside the workspace; saving it needs the user\'s approval.', { outside: true });
    if (access !== 'allowed' && access !== 'outside') fail('sandbox-denied', 'The file cannot be written from this Session.');
  }

  /**
   * Where a picked path really is, before the tab opens or writes it: the same read authorization as `open`, the
   * canonical path (workspace-relative inside the workspace, absolute outside) and the real path, which tells one
   * file under two spellings apart. A folder also reports whether this Session may write into it.
   */
  async probe(scope, filePath, peer, { kind = 'file', signal } = {}) {
    if (typeof filePath !== 'string' || !filePath || /[\0-\x1f]/.test(filePath)) fail('invalid-path', 'That path is not valid.');
    await this.authorize(peer, scope, signal, { action: 'open', path: filePath });
    let real;
    if (kind === 'folder') {
      real = await realpath(nodePath.resolve(scope.workspaceRoot, filePath)).catch(() => fail('not-found', 'That folder does not exist.'));
      if (!(await statPath(real)).isDirectory()) fail('not-found', 'That is not a folder.');
    } else {
      const { absolutePath } = await this.stat(scope, filePath, signal);
      real = await realpath(absolutePath);
      if (!suiteOf(real)) fail('unsupported-format', 'Only Word, Excel and PowerPoint files open here.');
      // OOXML is a ZIP: a file without its signature is damaged or something else under a document name.
      const handle = await openFile(real, 'r');
      try {
        const head = Buffer.alloc(4);
        const { bytesRead } = await handle.read(head, 0, 4, 0);
        if (bytesRead < 4 || head.readUInt32LE(0) !== 0x04034b50) fail('invalid-format', 'The file is damaged or not a Word, Excel or PowerPoint file.');
      } finally { await handle.close(); }
    }
    const root = await realpath(scope.workspaceRoot).catch(() => scope.workspaceRoot);
    const relative = nodePath.relative(root, real);
    const inside = relative === '' || (!relative.startsWith('..') && !nodePath.isAbsolute(relative));
    return { realPath: real, path: inside ? slashed(relative) : real, inWorkspace: inside,
      ...(kind === 'folder' && { access: await this.writeAccess(scope, real).catch(() => 'denied') }) };
  }

  /** One workspace folder's subfolders and Office files, for a picker without the desktop's own; never outside it. */
  async browse(scope, folder, peer, { signal } = {}) {
    await this.authorize(peer, scope, signal, { action: 'open', path: folder || '.' });
    const root = await realpath(scope.workspaceRoot);
    const real = await realpath(nodePath.resolve(root, folder || '.')).catch(() => fail('not-found', 'That folder does not exist.'));
    const relative = nodePath.relative(root, real);
    if (relative.startsWith('..') || nodePath.isAbsolute(relative)) fail('forbidden', 'Only folders inside the workspace can be browsed here.');
    const entries = [];
    for (const entry of await readdir(real, { withFileTypes: true })) {
      if (entry.name.startsWith('.')) continue;
      if (entry.isDirectory()) entries.push({ name: entry.name, kind: 'folder' });
      else if (entry.isFile() && suiteOf(entry.name)) entries.push({ name: entry.name, kind: 'file' });
    }
    entries.sort((a, b) => (a.kind === b.kind ? 0 : a.kind === 'folder' ? -1 : 1) || a.name.localeCompare(b.name));
    return { path: slashed(relative), entries };
  }

  /**
   * Exports to a Session path or, for a download, to a Host-private file the caller streams and deletes. A path that
   * exists is replaced only with `overwrite` (through a temporary file in the same folder, renamed over it) and never
   * when it is this view's own file. `approved`: the owner chose this place in the Save As panel, which allows it
   * outside the workspace; a read-only Session still refuses.
   */
  async export(viewId, peer, { target, signal } = {}) {
    const view = this.view(viewId, peer);
    this.#ready(view, 'export');
    const { extension } = SUITES[view.suite];
    let file, replaced;
    if (target === 'download') {
      mkdirSync(this.exportRoot, { recursive: true, mode: 0o700 });
      file = nodePath.join(this.exportRoot, `${randomUUID()}${extension}`);
    } else {
      if (typeof target?.path !== 'string' || !target.path.toLowerCase().endsWith(extension)) fail('invalid-path', `Save a copy as a ${extension} file.`);
      await this.authorize(peer, view.scope, signal, { action: 'save', path: target.path });
      const absolute = nodePath.resolve(view.scope.workspaceRoot, target.path);
      file = nodePath.join(await realpath(nodePath.dirname(absolute)).catch(() => fail('not-found', 'That folder does not exist.')), nodePath.basename(absolute));
      const approved = target.approved === true && view.principal === 'owner';
      if (existsSync(file)) {
        const [existing, own] = await Promise.all([statPath(file), statPath(view.realPath).catch(() => undefined)]);
        if ((own && existing.dev === own.dev && existing.ino === own.ino) || await realpath(file) === view.realPath) fail('same-file', 'That is this document itself; use Save instead.');
        if (target.overwrite !== true) fail('exists', 'A file with that name already exists.');
        replaced = file;
        file = nodePath.join(nodePath.dirname(file), `.${nodePath.basename(file, extension)}.${randomUUID()}.tmp${extension}`);
        await this.#mayWrite(view, replaced, approved);
      }
      await this.#mayWrite(view, file, approved);
    }
    await this.whenIdle(view.id);
    this.#ready(view, 'export');
    const client = await this.#client();
    const exported = await client.request('export', { documentId: view.documentId, target: file }, { signal }).catch(error => ({ ok: false, error }));
    if (!exported.ok) {
      if (target === 'download' || replaced) rmSync(file, { force: true });
      if (exported.error) throw exported.error;
      throw kitFailure(exported, 'exporting');
    }
    if (replaced) {
      try { await this.renameFile(file, replaced); }
      catch (error) {
        await rm(file, { force: true });
        // Windows refuses to replace a file another program holds open (Word, Excel, a preview): the user closes it there.
        if (this.platform === 'win32' && (error.code === 'EPERM' || error.code === 'EBUSY')) fail('target-busy', 'Another program has that file open; close it there, then save again.');
        fail('unavailable', `The file could not be replaced: ${error.code ?? error.message}`);
      }
      file = replaced;
    }
    return { file, name: nodePath.basename(view.path), contentType: contentTypeOf(view.suite) };
  }

  /** A download: the export waits under a one-time token the same tab redeems once. */
  async exportDownload(viewId, peer, { signal } = {}) {
    const { file, name, contentType } = await this.export(viewId, peer, { target: 'download', signal });
    const token = randomUUID();
    this.downloads.set(token, { file, name, contentType, viewId, peerId: peer.id, expires: this.now() + 5 * 60_000 });
    return { token, name };
  }
  takeDownload(viewId, peer, token) {
    const download = this.downloads.get(token);
    if (!download || download.viewId !== viewId || download.peerId !== peer?.id) fail('not-found', 'The download is no longer available.');
    this.downloads.delete(token);
    return download;
  }

  /** The disk copy changed while edits were kept: use the disk (`import`) or keep editing (local only). */
  async resolveDiskChange(viewId, peer, { useDisk, signal } = {}) {
    const view = this.view(viewId, peer);
    if (!useDisk) { view.diskChanged = false; return { kept: true }; }
    if (!view.writer) fail('read-only', 'Only the editing tab can replace the document with the disk copy.');
    this.#ready(view, 'reload');
    await this.authorize(peer, view.scope, signal, { action: 'edit', path: view.path });
    await this.whenIdle(view.id);
    const client = await this.#client();
    const imported = await client.request('import', { documentId: view.documentId }, { signal });
    if (!imported.ok) throw kitFailure(imported, 'reading the disk copy');
    this.#inert.delete(view.documentId);
    view.stale = true; // the page reopens with the imported content and a new change sequence
    return { imported: true };
  }

  /**
   * The tab's file may have been rewritten by someone else: the Session's current version must be the one
   * this view opened or last saved, else `source-changed` (the tab reopens; Kit reloads or asks).
   */
  async checkDisk(viewId, peer, signal) {
    const view = this.view(viewId, peer);
    // A save in flight is our own write: compare with the version it leaves behind.
    await view.writing;
    const { version } = await this.stat(view.scope, view.path, signal);
    if (version === undefined || version !== view.version) fail('source-changed', 'The file changed on disk since this view opened or saved it.');
    return { version };
  }

  /** AI applies are serialized by the page; while one runs the document cannot be saved or exported. */
  aiBegin(viewId, peer) { const view = this.view(viewId, peer); if (!view.writer) fail('read-only', 'Only the editing tab can be edited by AI.'); view.aiActive++; return () => { view.aiActive = Math.max(0, view.aiActive - 1); }; }

  /** Media bytes for the page; the reference stays exactly as the suite wrote it. */
  async readMedia(viewId, peer, { reference, purpose = 'display', range }) {
    const view = this.view(viewId, peer);
    const client = await this.#client();
    const opened = await client.request('open-media-read', { documentId: view.documentId, reference, purpose, ...(range && { range }) });
    if (opened.code === 'media-range-unsatisfiable') fail('range-unsatisfiable', 'The requested byte range is outside the media.');
    if (!opened.ok) throw kitFailure(opened, 'reading media');
    return { body: client.side.openBody(opened.stream, opened.length), contentType: opened.contentType, length: opened.length,
      ...(opened.totalLength !== undefined && { offset: opened.offset, totalLength: opened.totalLength }) };
  }

  /** Stores inserted bytes in Kit; saving waits until a change that references them has reached Kit. */
  async writeMedia(viewId, peer, { bytes, extension, kind }) {
    const view = this.view(viewId, peer);
    if (!view.writer) fail('read-only', 'Only the editing tab can insert media.');
    if (!/^[a-z0-9]{1,8}$/i.test(extension ?? '')) fail('invalid-body', 'The file extension is not supported.');
    // Each suite names its own media kind (sheet none); the Host refuses rather than guessing for the page.
    const { kinds } = SUITES[view.suite];
    if (!kinds.includes(kind ?? undefined)) fail('invalid-body', kinds[0] === undefined ? `A ${view.suite} media write takes no kind.` : `A ${view.suite} media kind must be ${kinds.join(' or ')}.`);
    // The hold starts before any byte goes out (kit-usage/05-host.md, insert step 1).
    const pending = `pending:${randomUUID()}`;
    view.holds.add(pending);
    try {
      const client = await this.#client();
      const begun = await client.request('begin-media-write', { documentId: view.documentId, extension: extension.toLowerCase(), ...(kind && { kind }) });
      if (!begun.ok) throw kitFailure(begun, 'storing media');
      await client.side.writeBody(begun.stream, bytes);
      const finished = await client.request('finish-media-write', { documentId: view.documentId, stream: begun.stream });
      if (!finished.ok) throw kitFailure(finished, 'storing media');
      view.holds.add(finished.reference);
      return { reference: finished.reference };
    } finally { view.holds.delete(pending); }
  }
  cancelMedia(viewId, peer, reference) {
    const view = this.view(viewId, peer);
    if (!view.writer) fail('read-only', 'Only the editing tab can insert media.');
    view.holds.delete(reference);
    return { cancelled: true };
  }

  async fontCatalog() {
    const client = await this.#client();
    const catalog = await client.request('font-catalog');
    if (!catalog.ok) throw kitFailure(catalog, 'listing fonts');
    // The page parses Kit's own response shape (office-kit parseSystemFontCatalog), `ok` included.
    return { ok: true, pending: catalog.pending, schemaVersion: catalog.schemaVersion, catalogVersion: catalog.catalogVersion, fonts: catalog.fonts };
  }
  async readFont(resourceId) {
    const client = await this.#client();
    const font = await client.request('font-read', { resourceId });
    if (!font.ok) throw kitFailure(font, 'reading a font');
    return { body: client.side.openBody(font.stream, font.length), contentType: font.contentType, length: font.length };
  }

  /**
   * The tab is gone. Closing never discards: unsaved edits stay in Kit's work root for the next open.
   * `autosave` saves first when the owner may write the file without asking.
   */
  async close(viewId, peer, { autosave = true } = {}) {
    const view = this.views.get(viewId);
    if (!view || view.closed || (peer && view.peerId !== peer.id)) return { closed: true };
    let saved = false;
    try {
      if (autosave && view.writer && !view.stale && view.unsaved && view.principal === 'owner' && !view.fault && !view.unknown && !view.aiActive && !view.holds.size
        && await this.authorize(view.peer, view.scope, undefined, { action: 'save', path: view.path }).then(() => true, () => false)
        && await this.writeAccess(view.scope, view.realPath).catch(() => 'denied') === 'allowed') {
        await this.whenIdle(view.id);
        const client = await this.#client();
        saved = (await client.request('save', { documentId: view.documentId, target: view.realPath })).ok === true;
        if (saved) this.#inert.delete(view.documentId);
      } else if (this.#onlyInert(view)) {
        // Nothing was edited: Kit drops the editor's housekeeping by reading the file again, so its record is clean.
        // A change still on its way may be a real edit, and another view may open meanwhile: decide again after both.
        await this.whenIdle(view.id);
        const client = await this.#client();
        if (this.#onlyInert(view) && (await client.request('import', { documentId: view.documentId })).ok === true) this.#inert.delete(view.documentId);
      }
    } catch { /* the edits stay in Kit's record */ }
    view.closed = true;
    this.views.delete(view.id);
    return { closed: true, saved };
  }

  /**
   * The page could not mount an editor for this view: it closes without autosaving or discarding, so the
   * tab can fall back to a read-only preview of the disk copy while Kit keeps any earlier unsaved edits.
   */
  abandon(viewId, peer) {
    const view = this.view(viewId, peer);
    view.closed = true;
    this.views.delete(view.id);
    return { closed: true, unsaved: view.unsaved };
  }

  /** Views whose tab stopped renewing are closed (without discarding). */
  async expire() {
    for (const view of [...this.views.values()]) if (view.expires <= this.now()) await this.close(view.id, undefined, { autosave: true });
    for (const [token, download] of this.downloads) if (download.expires <= this.now()) { this.downloads.delete(token); rmSync(download.file, { force: true }); }
  }

  async dispose() {
    for (const view of [...this.views.values()]) await this.close(view.id, undefined, { autosave: true });
    for (const download of this.downloads.values()) rmSync(download.file, { force: true });
    this.downloads.clear();
    await this.kit.stop?.();
  }
}
