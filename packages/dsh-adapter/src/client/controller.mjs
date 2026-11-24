import { acceptFrameMessage, frameCommand } from '../../../core/src/browser/frame.mjs';
import { savedForSwitch } from './toolbar.mjs';
import { hostUrl, parseSource, unwrapRemote } from './address.mjs';

const NS = 'shimoPreview';
const RENEW_MAX_MS = 300_000;
const FLUSH_TIMEOUT_MS = 5_000;
const FLUSH_BEHIND_AGENT_MS = 120_000;
const AUTOSAVE_MS = 1_000;
const AUTOSAVE_RETRY_MS = 5_000;
// How often an editing Kit tab asks whether the Session may still write its file (the conversation's permission can change).
const ACCESS_POLL_MS = 3_000;
// A Host that does not answer in time leaves the access as it was: the save goes ahead and the Host still checks.
const ACCESS_WAIT_MS = 4_000;
const ACCESSES = new Set(['allowed', 'outside', 'denied']);
const RELEASE_WAIT_MS = 10_000;
// How long the toolbar's Save shows its tick after a save.
const SAVED_TICK_MS = 1_200;
// AI requests that write to the document: the toolbar's own undo/redo waits for them.
const AI_WRITES = new Set(['apply', 'undo', 'rewrite']);
// Failures that may pass on their own; anything else waits for the next hand edit.
const AUTOSAVE_RETRY = new Set(['officesdk/unavailable', 'officesdk/timeout', 'gateway/cancelled']);
const NEVER = new AbortController().signal;
const accountRefusal = () => Object.assign(new Error('The account is not connected: editing and saving wait until it is.'), { code: 'officesdk/account' });

/**
 * One document tab: its session, frame messages, renewals and the revision-adoption rules.
 * The body drives it; the header action reads the same instance.
 */
export class DocumentController {
  constructor({ rpc, events, base, address, signal, update, scheduler = globalThis, timeoutMs = 120_000, download, locale }) {
    Object.assign(this, { rpc, events, base, address, update, scheduler, timeoutMs, download, locale });
    this.address = address;
    this.source = parseSource(address);
    this.signal = signal;
    this.lease = undefined;
    // The Host's `process.platform`, from its first lease and kept after release: paths are spelled its way.
    this.platform = undefined;
    this.frame = undefined;
    this.state = { mode: 'view', dirty: false, diverged: false, loading: true };
    this.pending = new Map();
    this.listeners = new Set();
    this.receive = event => this.onMessage(event);
    events.addEventListener('message', this.receive);
    signal?.addEventListener('abort', () => this.stop(), { once: true });
  }

  call(method, args, signal = this.controller?.signal) {
    return this.rpc.call('/api', `${NS}/${method}`, { args: { workspaceFileScopeId: this.source.sessionId, ...args } }, signal).then(unwrapRemote);
  }
  set(patch) {
    // A switch held back by unsaved edits says so only until they are on disk.
    if (patch.dirty === false) patch = { ...patch, modeBlocked: undefined };
    this.state = { ...this.state, ...patch }; this.update(this.state); for (const listener of this.listeners) listener(this.state);
  }
  subscribe(listener) { this.listeners.add(listener); return () => this.listeners.delete(listener); }
  get origin() { return this.lease ? new URL(hostUrl(this.lease.entry, this.base)).origin : undefined; }

  /** A fresh open for this revision; any previous session is released first. */
  async open(revision, { mode = this.state.mode, loaded, failed } = {}) {
    // While the account holds edits back, a new or reopened view is a read-only one.
    if (this.accountHold && mode === 'edit') mode = 'view';
    this.closeSession();
    this.revision = revision;
    this.request = { loaded, failed };
    this.controller = new AbortController();
    const { signal } = this.controller;
    this.set({ mode, dirty: false, diverged: false, loading: true, ready: false, error: undefined, engine: undefined, url: undefined, fallback: undefined, handEdited: undefined, autosaveFailed: undefined, modeBlocked: undefined, history: undefined,
      saveState: undefined, saveFailed: undefined, saveFailedSeen: undefined, authorizePrompt: undefined, pendingOpen: undefined, changeAnchors:undefined });
    this.scheduler.clearTimeout(this.savedTick);
    // The mode a retry or reopen returns to; a preview is only ever a fallback from it.
    if (mode !== 'preview') this.editorMode = mode;
    this.timer = this.scheduler.setTimeout(() => this.fail(new Error('Document rendering timed out')), this.timeoutMs);
    // The universal package downloads its document engine on first use (spec 261110-universal-package §4.2 step 7):
    // while open waits, show how far it is and do not count the download against the rendering timeout.
    const watchEngine = () => {
      if (signal.aborted || this.lease || !this.state.loading) return;
      void this.call('engineStatus', {}, signal).then(status => {
        if (signal.aborted || this.lease || !this.state.loading) return;
        if (status?.state === 'downloading' || status?.state === 'waiting') {
          // `fallback`: the editor gave up on npm and comes from another source (only when the download profile lists one after npm).
          this.set({ engine: { received: status.received ?? 0, total: status.total ?? 0, ...status.fallback?.includes('npm') && { fallback: true }, ...status.platformStatus === 'test' && { testing: true } } });
          this.scheduler.clearTimeout(this.timer);
          this.timer = this.scheduler.setTimeout(() => this.fail(new Error('Document rendering timed out')), this.timeoutMs);
        } else if (this.state.engine) this.set({ engine: undefined });
      }, () => {}).finally(() => { if (!signal.aborted && !this.lease && this.state.loading) this.engineTimer = this.scheduler.setTimeout(watchEngine, 1000); });
    };
    this.scheduler.clearTimeout(this.engineTimer);
    this.engineTimer = this.scheduler.setTimeout(watchEngine, 800);
    try {
      // Never abort the open RPC itself: a lease the host grants after we stop listening would be
      // orphaned until it expires. Stop waiting when the tab gives up, and release the lease if it comes.
      const opening = this.call('open', { path: this.source.path, revision, mode }, NEVER);
      const release = lease => { if (lease?.id) void this.call('release', { id: lease.id, autosave: true }, NEVER).catch(() => {}); };
      const lease = await Promise.race([opening, new Promise((_, reject) => signal.addEventListener('abort', () => reject(signal.reason), { once: true }))])
        .catch(error => { if (!signal.aborted) throw error; opening.then(release, () => {}); });
      if (signal.aborted) { release(lease); return; }
      if (!lease || typeof lease.id !== 'string' || lease.revision !== revision || !['docx', 'xlsx', 'pptx'].includes(lease.format)) throw new Error('Invalid document lease');
      this.lease = lease;
      if (typeof lease.platform === 'string') this.platform = lease.platform;
      this.savedVersion = lease.version;
      this.editRevision = lease.editRevision ?? 0;
      const url = new URL(hostUrl(lease.entry, this.base));
      url.searchParams.set('revision', String(revision));
      url.searchParams.set('parentOrigin', new URL(this.base).origin);
      // The interface language the tab opened with: the Kit page's editor language, selection UI and AI overlay follow it.
      if (this.locale) url.searchParams.set('lang', this.locale);
      // A parked session comes back in edit mode with its unsaved changes, whatever mode was asked for.
      this.set({ url: url.href, capabilities: lease.capabilities, format: lease.format, mode: lease.mode ?? mode, dirty: lease.unsaved === true, saveNeedsApproval: lease.saveNeedsApproval === true || undefined,
        backend: lease.backend, diskChanged: lease.diskChanged === true || undefined, overwritePrompt: undefined,
        // PPT's toolbar undo/redo stay in place, unavailable, until the new frame announces its history.
        history: lease.backend === 'kit' && lease.format === 'pptx' ? { canUndo: false, canRedo: false } : undefined,
        // Kit reports whether this Session may write the file; outside its writable roots the user authorizes edits once.
        access: lease.access, authorized: lease.access === undefined || lease.access === 'allowed' || this.autosaveApproved === true });
      this.scheduleRenew(lease.expiresInMs);
      this.scheduleAccessPoll();
    } catch (error) {
      if (signal.aborted) return;
      // A host without write access still previews: the editor is only needed for AI edits.
      if (mode === 'edit' && ['officesdk/unsupported-operation', 'officesdk/read-only', 'officesdk/sandbox-denied'].includes(error?.code)) return this.open(revision, { mode: 'view', loaded, failed });
      this.fail(error);
    }
  }

  /**
   * Harness bumps the revision on every source change. Rule 1: the file is what the session last
   * saved (by us or by the Host for the AI), keep the editor; the Host checks the version. Rule 2: a
   * clean or read-only session reopens. Rule 3: unsaved edits survive as a diverged session.
   */
  async adopt(revision, currentVersion, { loaded, failed }) {
    await this.switched;
    if (revision === this.revision) return;
    // The last open failed (the retry button bumps the revision): open again rather than keep the error.
    if (!this.lease && this.state.error) return this.open(revision, { mode: this.resumeMode(), loaded, failed });
    if (!this.lease) return;
    // A Kit view shows Kit's record, not the disk. The Host checks the file against what this view opened or
    // saved: our own write rebinds; any other change reopens the view, and Kit reloads a clean record or reports
    // the disk change of an edited one (diskChanged → the prompt). What the editor still holds goes first.
    if (this.state.backend === 'kit') {
      if (await this.rebind(revision, false, loaded, failed, { quiet: true })) return;
      if (!this.lease) return;
      await this.flushEditor().catch(() => {});
      return this.open(revision, { mode: this.state.mode, loaded, failed });
    }
    if (this.state.mode === 'edit') {
      if (await this.rebind(revision, false, loaded, failed, { quiet: currentVersion !== this.savedVersion })) return;
      if (!this.lease) return;
      if (this.state.dirty) {
        if (await this.rebind(revision, true, loaded, failed)) this.set({ diverged: true });
        return;
      }
    }
    return this.open(revision, { loaded, failed });
  }

  /** True once the session follows `revision`; `quiet` lets a source change fall through to the next rule. */
  async rebind(revision, diverged, loaded, failed, { quiet = false } = {}) {
    this.request = { loaded, failed };
    try {
      const result = await this.call('rebind', { id: this.lease.id, revision, diverged });
      this.post(frameCommand({ id: this.lease.id, revision: this.revision, command: 'rebind', newRevision: revision }));
      this.revision = revision;
      if (result?.version) { this.savedVersion = result.version; this.lease.version = result.version; }
      loaded?.(this.savedVersion);
      return true;
    } catch (error) {
      if (quiet && error?.code === 'officesdk/source-changed') return false;
      // A Kit view that went stale (the document service restarted) reopens like a changed file.
      if (quiet && this.state.backend === 'kit' && error?.details?.rebuild) return false;
      this.fail(error);
      return false;
    }
  }

  post(message) { this.frame?.contentWindow?.postMessage(message, this.origin); }

  /** Calls `listener` on each press inside the editor page; returns the unsubscribe. */
  onPress(listener) { (this.pressListeners ??= new Set()).add(listener); return () => this.pressListeners.delete(listener); }

  onMessage(event) {
    if (!this.lease) return;
    const message = acceptFrameMessage(event, { frame: this.frame, origin: this.origin, id: this.lease.id, revision: this.revision });
    if (!message) return;
    if (message.event === 'ready') { this.scheduler.clearTimeout(this.timer); this.set({ loading: false, ready: true }); this.request?.loaded?.(this.lease.version); this.settleReady(); this.onReady?.(); }
    else if (message.event === 'agent-result') {
      const failure = message.ok ? undefined : Object.assign(new Error(message.message || 'The editor could not do that'), { code: `officesdk/${message.code ?? 'editor-failed'}` });
      this.settle(`agent:${message.requestId}`, message.result, failure);
    }
    else if (message.event === 'failed') {
      if (message.rebuild === true && this.state.backend === 'kit') return void this.reopen();
      if (this.state.backend === 'kit') return void this.fallBackToPreview(message);
      this.fail(new Error(message.message || 'Document rendering failed'));
    }
    else if (message.event === 'dirty') {
      this.editRevision = message.editRevision ?? this.editRevision;
      // A hand edit ends the AI undo history: the editor's own undo is the way back now.
      this.set({ dirty: message.dirty !== false, unstaged: typeof message.unstaged === 'string' ? message.unstaged : undefined, ...message.manual === true && { edits: [], handEdited: true } });
      // Hand edits in a Kit editor are written back once the user pauses (spec §13.8); AI edits wait for a save.
      if (message.manual === true && message.recovered !== true) this.scheduleAutosave();
    }
    // Selection rewrite (spec 261109-selection-ai-edit): the page's request goes to the conversation (selection-bridge.mjs);
    // its later fate (cancelled, kept, reverted) is the Host's to know.
    else if (message.event === 'open-website') this.onOpenWebsite?.();
    else if (message.event === 'selection-request') { if (this.state.mode === 'edit' && /^[A-Za-z0-9_-]{1,64}$/.test(message.requestId ?? '')) void this.onSelectionRequest?.(message); }
    else if (message.event === 'document-ui' && this.documentUi && this.state.mode === 'edit') {
      if (message.permission === 'always' || message.permission === 'ask') this.setAiPermission(message.permission);
      if ('chatSelection' in message) this.chatSelection=message.chatSelection ?? undefined;
      if (Array.isArray(message.restoredChanges)) {
        this.documentUi.changes=message.restoredChanges;
        this.set({changes:message.restoredChanges});
      }
      if (typeof message.reviewSaveFailed==='boolean') this.set({reviewSaveFailed:message.reviewSaveFailed});
      if (message.revealChange === null || (typeof message.revealChange === 'string' && this.documentUi.changes?.some(r=>r.id===message.revealChange))) {
        this.set({revealedChange:{id:message.revealChange}});
      }
      if (Number.isFinite(message.dockBottom)) this.set({ dockBottom: message.dockBottom });
      if (message.changeAnchors?.height>0 && Array.isArray(message.changeAnchors.items)) this.set({changeAnchors:message.changeAnchors});
      if (message.draftState) {
        const { draft, references, pending, candidates, changes = [], next } = message.draftState;
        if (typeof draft === 'string' && [references, pending, candidates].every(Array.isArray) && Number.isSafeInteger(next)) {
          Object.assign(this.documentUi, { draft, references, pending, candidates, next, changes });
          this.set({ changes });
        }
      }
    }
    else if (message.event === 'selection-update') { if (/^sel-[0-9a-f]{8}$/.test(message.selection ?? '') && typeof message.state === 'string') void this.call('selectionUpdate', { id: message.selection, selectionState: message.state }).catch(() => {}); }
    // A press inside the editor page: the header closes its open menus (the page is cross-origin, so no outside click reaches it).
    else if (message.event === 'press') { for (const listener of this.pressListeners ?? []) listener(); }
    else if (message.event === 'history') this.set({ history: { canUndo: message.canUndo === true, canRedo: message.canRedo === true } });
    else if (message.event === 'save-request') { if (this.state.mode === 'edit' && this.state.capabilities?.save) this.onSaveRequest?.(); }
    else if (message.event === 'saved') {
      this.editRevision = message.editRevision ?? this.editRevision;
      if (message.version !== undefined) { this.savedVersion = message.version; this.lease.version = message.version; this.set({ dirty: false }); }
      this.settle('saved', message);
    } else if (message.event === 'conflict') { this.set({ conflict: { draftId: message.draftId } }); this.settle('saved', undefined, Object.assign(new Error('The document changed outside this editor'), { code: 'officesdk/source-changed', details: { draftId: message.draftId } })); }
  }

  settle(kind, value, error) {
    const entry = this.pending.get(kind);
    if (!entry) return;
    this.pending.delete(kind);
    error ? entry.reject(error) : entry.resolve(value);
  }

  /** Resolves once the frame of the current session reports ready; a failure or tab close rejects. */
  whenReady(signal) {
    if (this.state.ready && this.lease) return Promise.resolve();
    if (this.state.error) return Promise.reject(Object.assign(new Error(this.state.error), { code: this.state.errorCode }));
    return new Promise((resolve, reject) => {
      const waiter = { resolve, reject };
      (this.readyWaiters ??= new Set()).add(waiter);
      signal?.addEventListener('abort', () => { this.readyWaiters.delete(waiter); reject(signal.reason ?? new Error('Cancelled')); }, { once: true });
    });
  }
  settleReady(error) {
    for (const waiter of this.readyWaiters ?? []) error ? waiter.reject(error) : waiter.resolve();
    this.readyWaiters?.clear();
  }

  /**
   * The header's pen/eye: reopens the document in `mode`. A Kit session parked with unsaved edits always
   * reopens editing, so leaving edit first writes them, and the tab stays editing (`modeBlocked`) while any
   * edit is not on disk. A failed save throws for the header to show. Returns whether the mode changed.
   * An AI request or a source change that arrives meanwhile waits until the new view is ready (`switched`).
   */
  switchMode(mode) {
    // A tab whose last open failed has no session; the buttons then open it again in the mode asked for.
    if (this.switching || (this.state.mode === mode && !this.state.error) || !['edit', 'view'].includes(mode)) return Promise.resolve(false);
    this.switching = true;
    const running = this.runSwitch(mode);
    // A source change waiting on the switch (its own save bumps the revision) rebinds a loaded frame only.
    const switched = running.then(changed => changed && this.whenReady(), () => {}).catch(() => {})
      .finally(() => { if (this.switched === switched) this.switched = undefined; });
    this.switched = switched;
    return running;
  }

  async runSwitch(mode) {
    this.set({ switching: true, modeBlocked: undefined });
    try {
      if (this.lease && this.state.mode === 'edit' && mode === 'view') {
        this.scheduler.clearTimeout(this.autosaveTimer);
        await this.flushEditor();
        const writable = this.state.authorized === true && this.state.access !== 'denied';
        if (this.state.dirty && writable) await this.save({ approved: this.autosaveApproved === true });
        if (this.state.dirty || this.state.unstaged) { this.set({ modeBlocked: true }); return false; }
      }
      // Kit closes the old view before it opens the file again: an open racing that close is aborted. The old
      // frame goes first, so nothing is typed into a view that is closing.
      this.set({ loading: true, url: undefined });
      let timer;
      await Promise.race([this.closeSession(), new Promise(resolve => { timer = this.scheduler.setTimeout(resolve, RELEASE_WAIT_MS); })]);
      this.scheduler.clearTimeout(timer);
      // The tab closed while Kit released the view: nothing is left to open it for.
      if (this.stopped) return false;
      await this.open(this.revision, { mode, ...this.request });
      return true;
    } finally {
      this.switching = false;
      this.set({ switching: false });
    }
  }

  /** Switches the open document to `mode` (keeping its revision) and waits for the new frame. */
  async ensureMode(mode, signal) {
    // An AI edit needs editing, which the account hold does not allow: refuse before reopening anything.
    if (mode === 'edit' && this.accountHold) throw accountRefusal();
    if (this.switched) {
      // A cancelled AI request stops waiting for the header's switch.
      await Promise.race([this.switched, new Promise((_, reject) => {
        if (signal?.aborted) reject(signal.reason);
        signal?.addEventListener('abort', () => reject(signal.reason ?? new Error('Cancelled')), { once: true });
      })]);
    }
    if (this.state.mode !== mode || !this.lease) await this.open(this.revision, { mode, ...this.request });
    return this.whenReady(signal);
  }

  /** One request to the frame's editor agent; the frame answers with agent-result for the same id. */
  agent(op, args = {}, signal) {
    if (!this.lease) return Promise.reject(Object.assign(new Error('The document is not open'), { code: 'officesdk/no-editor' }));
    const requestId = `a${Date.now().toString(36)}${Math.random().toString(36).slice(2, 10)}`;
    const { id } = this.lease, revision = this.revision;
    const writes = AI_WRITES.has(op) || (op==='changeDetail' && args.action==='remove');
    if (writes && this.accountHold) return Promise.reject(accountRefusal());
    if (writes) this.set({ aiBusy: (this.aiWrites = (this.aiWrites ?? 0) + 1) > 0 });
    const done = () => { if (writes) this.set({ aiBusy: (this.aiWrites -= 1) > 0 }); };
    return new Promise((resolve, reject) => {
      const kind = `agent:${requestId}`;
      const onAbort = () => {
        this.post(frameCommand({ id, revision: this.revision, command: 'agent', requestId: `c${requestId.slice(1)}`, op: 'cancel', args: { target: requestId } }));
      };
      this.pending.set(kind, { resolve: value => { signal?.removeEventListener('abort', onAbort); done(); resolve(value); }, reject: error => { signal?.removeEventListener('abort', onAbort); done(); reject(error); } });
      signal?.addEventListener('abort', onAbort, { once: true });
      this.post(frameCommand({ id, revision, command: 'agent', requestId, op, args }));
    });
  }

  /** Shows a proposed change in the tab and waits for Apply (true) or Discard (false). */
  propose(proposal, signal) {
    if (signal?.aborted) return Promise.reject(signal.reason ?? new Error('Cancelled'));
    this.decision?.resolve(false);
    if (this.state.aiPermission === 'always' && !proposal.danger) return Promise.resolve(true);
    return new Promise((resolve, reject) => {
      const clear = () => {
        signal?.removeEventListener('abort', abort);
        this.decision = undefined;
        this.set({ proposal: undefined });
      };
      const finish = value => { clear(); resolve(value); };
      const abort = () => { clear(); reject(signal.reason ?? new Error('Cancelled')); };
      this.decision = { resolve: finish };
      this.set({ proposal });
      signal?.addEventListener('abort', abort, { once: true });
    });
  }
  decide(approved) { this.decision?.resolve(approved === true); }
  syncDocumentUi() {
    if (this.documentUi && this.state.ready && this.state.mode === 'edit') {
      void this.agent('documentUi', { state: this.documentUi, permission: this.state.aiPermission ?? 'ask', focusMode: Boolean(this.state.focusMode) }).catch(() => {});
    }
  }
  setAiPermission(mode) {
    if (mode !== 'ask' && mode !== 'always') return;
    this.onAiPermission?.(mode);
    this.set({ aiPermission: mode });
    if (mode === 'always' && !this.state.proposal?.danger) this.decide(true);
    if (this.state.ready) void this.agent('documentUi', { permission: mode }).catch(() => {});
  }

  /**
   * Writes the staged edits to the file. Outside the workspace the Host refuses with `outside`;
   * the panel asks the user and calls again with `approved`.
   */
  async save({ approved = false, overwrite = false, auto = false } = {}) {
    if (!this.lease || this.state.mode !== 'edit') throw new Error('Only an editing session can save');
    if (this.accountHold) throw accountRefusal();
    // One save at a time: a hand save waits for a running autosave, then writes what is new.
    while (this.saving) await this.saving.catch(() => {});
    const running = this.writeOnce({ approved, overwrite, auto });
    this.saving = running;
    try { return await running; } finally { if (this.saving === running) this.saving = undefined; }
  }

  /**
   * The toolbar's Save and Cmd/Ctrl+S: what the editor holds is taken first; a document with nothing new is not
   * written. A running autosave is waited for and counts when nothing changed since. Clicks while one runs share it.
   */
  saveNow(options = {}) {
    if (this.accountHold) return Promise.reject(accountRefusal());
    this.handSave ??= (async () => {
      this.scheduler.clearTimeout(this.savedTick);
      this.set({ saveState: 'saving', saveFailed: undefined, saveFailedSeen: undefined });
      try {
        // A dirty document's save takes the editor's input itself. A clean-looking one (also after an autosave that
        // left an IME composition or an edited cell alone) is asked for its input first.
        while (this.saving) await this.saving.catch(() => {});
        if (!this.state.dirty && !this.state.unstaged) await this.flushEditor();
        if (this.state.dirty || this.state.unstaged) await this.save(options);
        this.set({ saveState: 'saved' });
        this.savedTick = this.scheduler.setTimeout(() => { if (this.state.saveState === 'saved') this.set({ saveState: undefined }); }, SAVED_TICK_MS);
      } catch (error) {
        // The Session turned read only since the last check: say so, not the Host's refusal.
        if (this.state.backend === 'kit' && error?.code === 'officesdk/sandbox-denied' && !error.details?.outside) {
          this.readOnly();
          this.set({ saveState: 'failed', saveFailed: this.messages?.readOnlyHint ?? 'The Session is read only.', saveFailedSeen: undefined });
        } else this.set({ saveState: 'failed', saveFailed: error?.message || String(error) });
        throw error;
      } finally { this.handSave = undefined; }
    })();
    return this.handSave;
  }

  /**
   * The account holds edits back (preview of AUTH-01): nothing is written — no autosave, save, authorization, AI edit
   * or autosave on release — and new views open read-only. Lifting it saves pending edits at the next pause.
   */
  setAccountHold(held) {
    const next = held === true;
    if (next === Boolean(this.accountHold)) return;
    this.accountHold = next;
    this.set({ accountHold: next || undefined });
    if (next) this.scheduler.clearTimeout(this.autosaveTimer);
    else if (this.state.dirty && this.state.authorized) this.scheduleAutosave();
  }

  /** A save that is current ends an earlier failure's notice, whoever saved (autosave included). */
  cleared() { return this.state.saveState === 'failed' ? { saveState: undefined, saveFailed: undefined, saveFailedSeen: undefined } : {}; }

  /**
   * Holds autosave while Save As is open or a switch saves for the last time; a Session that cannot write is a
   * different stop (`access: 'denied'`) that lasts until the Session may write again. Resuming saves pending edits at the next pause.
   */
  pauseAutosave(paused) {
    this.autosavePaused = paused === true;
    if (this.autosavePaused) this.scheduler.clearTimeout(this.autosaveTimer);
    else if (this.state.dirty) this.scheduleAutosave();
  }

  /**
   * OPEN-01's last save before the tab switches files: the editor takes no input (the tab covers it) and autosave
   * waits; saved means written with nothing new arriving meanwhile. Three tries, then the tab stays and says so.
   * Prompts a save raises (authorization, overwrite) reject here and are the caller's.
   */
  async saveForSwitch(options = {}) {
    if (!this.lease || this.state.mode !== 'edit') return true;
    this.pauseAutosave(true);
    this.set({ locked: true });
    try {
      for (let attempt = 0; attempt < 3; attempt++) {
        const before = this.editRevision;
        await this.saveNow(options);
        if (savedForSwitch(this.state, before, this.editRevision)) return true;
      }
      this.set({ saveState: 'failed', saveFailed: this.messages?.keptEditing ?? 'New edits kept arriving while saving.', saveFailedSeen: undefined });
      return false;
    } finally {
      this.set({ locked: undefined });
      this.pauseAutosave(false);
    }
  }

  /** A picked file or folder as the Host sees it (canonical path, real path; a folder's write access). */
  probe(path, kind = 'file') { return this.call('probe', { path, kind }); }
  browse(path = '') { return this.call('browse', { path }); }
  /** This tab's own file, by its real path: picking it again does nothing. */
  ownRealPath() {
    this.ownReal ??= this.probe(this.source.path).then(result => result.realPath, () => { this.ownReal = undefined; return undefined; });
    return this.ownReal;
  }

  /** The toolbar's undo/redo of the editor's own history (PPT). */
  editHistory(action) { return this.agent('history', { action }); }
  /** Opens the editor's print dialog; resolves once the frame took the request. */
  print() { return this.agent('print'); }

  scheduleAutosave(delay = AUTOSAVE_MS) {
    if (this.state.backend !== 'kit' || this.state.mode !== 'edit' || this.autosavePaused || this.accountHold) return;
    this.scheduler.clearTimeout(this.autosaveTimer);
    this.autosaveTimer = this.scheduler.setTimeout(() => { void this.autosave(); }, delay);
  }

  /**
   * Writes a Kit view's hand edits back without asking once editing is authorized: always for a file the
   * Session may write, and after authorize() for one outside its writable roots. A read-only Session holds it
   * (the tab says read only, nothing is reported as failed) until the Session may write again; a busy document
   * tries again at the next pause; a file another program rewrote is left to the disk-change flow (the revision
   * bump reopens the view), never overwritten here.
   */
  async autosave() {
    if (!this.autosaveReady()) return;
    // The conversation may have turned read only since the last poll: nothing is handed over or written then. While
    // the Host answered, Save As, a switch's last save, the account hold or a reopen may have begun: check again.
    const lease = this.lease;
    if (await this.refreshAccess() === 'denied' || this.lease !== lease || !this.autosaveReady()) return;
    try { await this.save({ auto: true, approved: this.autosaveApproved === true }); }
    catch (error) {
      // A Host that did not report the file's access finds out here: the tab then asks for authorization.
      if (error?.code === 'officesdk/sandbox-denied' && error.details?.outside && !this.autosaveApproved) this.set({ access: 'outside', authorized: false });
      else if (error?.code === 'officesdk/sandbox-denied') this.readOnly();
      else if (error?.code === 'officesdk/busy') this.scheduleAutosave();
      else {
        // A failure while the Session turned read only (the hand-over timing out, say) is not a failed save; whichever
        // question saw it (this one or a poll), the tab's state says so. A failure of a replaced view is not this one's.
        await this.refreshAccess();
        if (this.state.access === 'denied' || this.lease !== lease) return;
        // Edits that did not reach the file must not look saved: say so, and retry failures that may pass.
        this.set({ autosaveFailed: error?.message || String(error) });
        if (!error?.code || AUTOSAVE_RETRY.has(error.code)) this.scheduleAutosave(AUTOSAVE_RETRY_MS);
      }
    }
  }

  autosaveReady() {
    return Boolean(this.lease) && this.state.backend === 'kit' && this.state.mode === 'edit' && this.state.dirty && !this.autosavePaused && !this.accountHold && this.state.authorized === true;
  }

  /** The Session may not write this file (now): the tab says read only, as if it had opened so, and reports no autosave failure. */
  readOnly() {
    this.scheduler.clearTimeout(this.autosaveTimer);
    this.set({ access: 'denied', authorized: false, autosaveFailed: undefined });
  }

  /**
   * Asks the Host whether the Session may still write this editing Kit view's file and follows the answer: read
   * only, writable again (pending edits autosave at the next pause; an outside file needs the authorization it had),
   * or outside. An answer for a lease that was replaced meanwhile, an unknown answer or a failed question change nothing.
   */
  async refreshAccess() {
    const lease = this.lease;
    if (!lease || this.state.backend !== 'kit' || this.state.mode !== 'edit' || this.accessUnsupported) return undefined;
    // A poll and an autosave may ask at once: only the latest question's answer counts.
    const asked = this.accessAsked = (this.accessAsked ?? 0) + 1;
    let access, timer;
    try {
      const late = new Promise(resolve => { timer = this.scheduler.setTimeout(() => resolve(undefined), ACCESS_WAIT_MS); });
      ({ access } = await Promise.race([this.call('access', { id: lease.id }), late]) ?? {});
    } catch (error) {
      // A Host from before this question (DSH says method-not-found): keep the access the tab opened with.
      if (/(^|\/)method-not-found$/.test(error?.code ?? '')) this.accessUnsupported = true;
      return undefined;
    } finally { this.scheduler.clearTimeout(timer); }
    if (asked !== this.accessAsked || this.lease !== lease || this.state.mode !== 'edit' || !ACCESSES.has(access)) return undefined;
    if (access === this.state.access) return access;
    if (access === 'denied') { this.readOnly(); return access; }
    const authorized = access === 'allowed' || this.autosaveApproved === true;
    const wasReadOnly = this.state.access === 'denied';
    this.set({ access, authorized, ...(wasReadOnly && this.state.saveState === 'failed' && this.cleared()) });
    if (authorized && this.state.dirty) this.scheduleAutosave();
    return access;
  }

  scheduleAccessPoll() {
    this.scheduler.clearTimeout(this.accessTimer);
    if (!this.lease || this.state.backend !== 'kit' || this.state.mode !== 'edit' || this.accessUnsupported) return;
    const lease = this.lease;
    this.accessTimer = this.scheduler.setTimeout(async () => {
      await this.refreshAccess();
      if (this.lease === lease) this.scheduleAccessPoll();
    }, ACCESS_POLL_MS);
  }

  /** The user authorizes editing the original file: what the editor holds is written now, then autosaved. */
  async authorize() {
    if (this.accountHold) throw accountRefusal();
    this.autosaveApproved = true;
    this.set({ authorized: true });
    if (this.lease && this.state.backend === 'kit' && this.state.mode === 'edit' && this.state.dirty) await this.save({ approved: true });
  }

  async writeOnce({ approved, overwrite, auto }) {
    if (!this.lease || this.state.mode !== 'edit') throw new Error('Only an editing session can save');
    // An autosave takes only what already reached the editor's queue; it never ends the user's input.
    await this.flushEditor({ settle: !auto });
    const editRevision = this.editRevision;
    const lease = this.lease;
    const result = await this.call('save', { id: lease.id, editRevision, approved, overwrite }, NEVER);
    // The view may have been reopened while the save was on its way (a file change arrived first).
    if (this.lease !== lease) return result;
    // Kit's save reports the file's new version; the save is current when nothing was edited meanwhile.
    if (this.state.backend === 'kit') {
      if (result.version !== undefined) { this.savedVersion = result.version; this.lease.version = result.version; }
      if (result.saved && this.editRevision === editRevision) this.set({ dirty: false, saveOutside: undefined, saveNeedsApproval: undefined, overwritePrompt: undefined, ...this.cleared() });
      if (result.saved) this.set({ autosaveFailed: undefined });
    }
    else this.markSaved(result.version, result.editRevision);
    return result;
  }

  /**
   * The Kit editor could not mount (spec §13.4): the page already closed its view without saving, so the
   * tab shows Core's read-only preview of the disk copy. Edits kept in Kit's record are announced, not lost.
   */
  async fallBackToPreview(message) {
    const unsaved = message.unsaved === true || this.state.dirty === true;
    // The page already abandoned its view; if that request was lost, this release still must not save.
    this.abandoning = true;
    await this.open(this.revision, { mode: 'preview', ...this.request });
    // Without a Core preview (a Kit-only package) the tab stays on the editor, so a retry opens Kit again.
    this.set(this.state.error ? { mode: this.resumeMode() } : { mode: 'view', fallback: unsaved ? 'unsaved' : 'preview' });
  }

  /** The same document again from the Host (a restarted document service or a disk import). */
  reopen() { return this.open(this.revision, { mode: this.resumeMode(), ...this.request }); }

  /** The tab's own mode (edit or view), never the preview it may have fallen back to. */
  resumeMode() { return this.state.mode === 'preview' || this.state.fallback ? this.editorMode ?? 'view' : this.state.mode; }

  /** The disk copy changed while edits were kept: use it (the editor reloads) or keep editing. */
  async resolveDiskChange(useDisk) {
    if (!this.lease) return;
    await this.call('resolveDiskChange', { id: this.lease.id, useDisk: useDisk === true });
    if (useDisk) return this.reopen();
    this.set({ diskChanged: undefined });
  }

  /** The Host saved this session (here or for the AI); the next revision bump is our own write. */
  markSaved(version, editRevision) {
    if (!this.lease) return;
    // A Kit save (here or the AI's) reports the file's version; only the change count says whether it is current.
    if (this.state.backend === 'kit') {
      if (version !== undefined) { this.savedVersion = version; this.lease.version = version; }
      if (editRevision === undefined || editRevision === this.editRevision) this.set({ ...this.cleared(),  dirty: false, saveOutside: undefined, saveNeedsApproval: undefined, overwritePrompt: undefined });
      return;
    }
    if (version === undefined) return;
    this.savedVersion = version;
    this.lease.version = version;
    if (editRevision === undefined || editRevision === this.editRevision) this.set({ dirty: false, saveOutside: undefined, saveNeedsApproval: undefined });
  }

  /**
   * Hand edits in a Kit editor (any format) or Core's PPTX editor reach the Host a moment after they are
   * made: save and export first ask the editor to hand over what it still holds, so they never write older
   * bytes. An editor that cannot (Core's DOCX/XLSX, view mode) is skipped; one that does not answer in time
   * fails the save rather than losing the edit.
   */
  async flushEditor({ settle = true } = {}) {
    if (!this.lease || (this.lease.format !== 'pptx' && this.state.backend !== 'kit') || this.state.mode !== 'edit' || !this.state.ready) return;
    let timer;
    // Behind a running AI edit the hand-over waits for the edit to finish.
    const waitMs = [...this.pending.keys()].some(kind => kind.startsWith('agent:')) ? FLUSH_BEHIND_AGENT_MS : FLUSH_TIMEOUT_MS;
    const late = new Promise((_, reject) => { timer = this.scheduler.setTimeout(() => reject(Object.assign(new Error('The editor did not hand over its latest changes; try again.'), { code: 'officesdk/editor-failed' })), waitMs); });
    try {
      const result = await Promise.race([this.agent('flush', settle ? {} : { settle: false }), late]);
      if (Number.isSafeInteger(result?.editRevision) && result.editRevision > this.editRevision) this.editRevision = result.editRevision;
    } catch (error) {
      if (error?.code !== 'officesdk/unsupported-operation') throw error;
    } finally { this.scheduler.clearTimeout(timer); }
  }

  /** Downloads start in the parent page: the frame has no allow-downloads. */
  async exportDownload() {
    if (!this.lease) throw new Error('No document is open');
    await this.flushEditor();
    const result = await this.call('exportTo', { id: this.lease.id, editRevision: this.editRevision, target: 'download' });
    this.download?.(hostUrl(result.url, this.base), result.name);
    return result;
  }
  /** Save As: a copy of what the editor holds; `overwrite` replaces an existing file, `approved` the owner's place. */
  async exportCopy(path, { overwrite = false, approved = false } = {}) {
    // A copy is a save too (docs/01 B04): the account hold applies.
    if (this.accountHold) throw accountRefusal();
    await this.flushEditor();
    return this.call('exportTo', { id: this.lease.id, editRevision: this.editRevision, target: { path, overwrite, approved } });
  }

  listDrafts() { return this.call('listDrafts', {}, undefined); }
  discardDraft(draftId) { return this.call('discardDraft', { draftId }, undefined); }
  async downloadDraft(draftId) {
    const result = await this.call('exportDraft', { draftId, target: 'download' }, undefined);
    this.download?.(hostUrl(result.url, this.base), result.name);
    return result;
  }

  scheduleRenew(expiresInMs) {
    const interval = Number.isFinite(expiresInMs) && expiresInMs > 0 ? Math.max(1, Math.min(RENEW_MAX_MS, Math.floor(expiresInMs / 3))) : RENEW_MAX_MS;
    this.renewTimer = this.scheduler.setTimeout(async () => {
      try {
        const result = await this.call('renew', { id: this.lease.id });
        if (this.lease) this.scheduleRenew(result?.expiresInMs ?? expiresInMs);
      } catch (error) {
        // A Kit view goes stale when the document service restarts: reopen; its unsaved edits are in Kit.
        if (error?.details?.rebuild && this.state.backend === 'kit') return this.reopen();
        this.fail(error);
      }
    }, interval);
  }

  fail(error) {
    this.scheduler.clearTimeout(this.timer);
    this.scheduler.clearTimeout(this.renewTimer);
    this.scheduler.clearTimeout(this.accessTimer);
    const message = error instanceof Error ? error.message : String(error);
    this.set({ loading: false, error: message, errorCode: error?.code });
    this.request?.failed?.();
    for (const kind of [...this.pending.keys()]) this.settle(kind, undefined, error);
    this.settleReady(error);
    this.closeSession();
  }

  closeSession() {
    // Agent requests belong to one frame; a replaced or released session cannot answer them.
    const gone = Object.assign(new Error('The document session was replaced'), { code: 'officesdk/no-editor' });
    for (const kind of [...this.pending.keys()]) if (kind.startsWith('agent:')) this.settle(kind, undefined, gone);
    this.controller?.abort();
    this.scheduler.clearTimeout(this.timer);
    this.scheduler.clearTimeout(this.renewTimer);
    this.scheduler.clearTimeout(this.accessTimer);
    this.scheduler.clearTimeout(this.autosaveTimer);
    let released;
    if (this.lease) { const { id } = this.lease; this.lease = undefined; released = this.call('release', { id, autosave: !this.abandoning && !this.accountHold }, undefined).catch(() => {}); }
    this.abandoning = false;
    return released;
  }

  stop() {
    this.autosavePaused = false;
    this.scheduler.clearTimeout(this.savedTick);
    if (this.stopped) return;
    this.stopped = true;
    this.events.removeEventListener('message', this.receive);
    const closed = Object.assign(new Error('The document tab closed'), { code: 'officesdk/no-editor' });
    for (const kind of [...this.pending.keys()]) this.settle(kind, undefined, closed);
    this.settleReady(closed);
    this.decision?.resolve(false);
    this.closeSession();
  }
}
