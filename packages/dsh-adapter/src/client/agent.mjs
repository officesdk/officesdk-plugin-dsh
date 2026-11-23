import { describePptxSteps, summarizeDeck, validatePptxOps } from '../../../core/src/agent/pptx.mjs';
import { fileAddress, unwrapRemote } from './address.mjs';
import { runDocumentCommand } from './document-agent.mjs';
import { runSpreadsheetCommand } from './spreadsheet-agent.mjs';
import { callKey, editContext, editKey, readOutline, settleContext } from './tool-views.mjs';

const NS = 'shimoPreview';
/** Formats the AI edits inside the open editor; the others are not offered yet. */
export const AGENT_FORMATS = Object.freeze({ pptx: 'edit', docx: 'edit', xlsx: 'edit' });

/** The Session file address the document preview opens for `path`: the one tabs use (`fileAddress`). */
export { fileAddress as sessionAddress } from './address.mjs';

/**
 * Opening a Session's file from its chat row: the same address the AI's tab uses (so no second editor for the same
 * file), and only for the Session the right Sidebar shows (the one the bridge serves); `subscribeDeck` follows it.
 */
export function deckOpener(sidebarOf, platform) {
  return {
    // `platform`: the Host's, as a value or a getter, the same one AgentBridge spells the AI tab's address with.
    openDeck: (sessionId, path) => sidebarOf()?.openResource?.(fileAddress(sessionId, path, typeof platform === 'function' ? platform() : platform)),
    canOpenDeck: sessionId => { const sidebar = sidebarOf(); return typeof sidebar?.openResource === 'function' && sidebar.mounted?.getSnapshot?.() === sessionId; },
    subscribeDeck: listener => sidebarOf()?.mounted?.subscribe?.(listener) ?? (() => {}),
  };
}

const formatOf = path => /\.([A-Za-z0-9]+)$/.exec(path)?.[1]?.toLowerCase();
const failure = (code, message) => Object.assign(new Error(message), { code: `officesdk/${code}` });

/**
 * Serves the Host's AI document commands for the Session on screen. Each command opens (or focuses)
 * the document tab, works through that tab's controller, and answers with a result or a coded error.
 * `platform`: the Host's, as a value or a getter (a lease says it); a command that carries its own wins.
 */
export class AgentBridge {
  constructor({ rpc, openResource, controllers, mounted, views, platform, scheduler = globalThis, retryMs = 2_000, aliveMs = 15_000, generatingMs = 60_000 }) {
    Object.assign(this, { rpc, openResource, controllers, mounted, views, platform, scheduler, retryMs, aliveMs, generatingMs });
    this.running = new Map();
    this.loop = undefined;
  }

  call(sessionId, method, args, signal) {
    return this.rpc.call('/api', `${NS}/${method}`, { args: { workspaceFileScopeId: sessionId, ...args } }, signal).then(unwrapRemote);
  }

  /** Follows the on-screen Session: one long-poll loop at a time, restarted when the Session changes. */
  start() {
    const follow = () => {
      const sessionId = this.mounted.getSnapshot();
      if (this.loop?.sessionId === sessionId) return;
      this.loop?.controller.abort();
      this.loop = sessionId ? { sessionId, controller: new AbortController() } : undefined;
      if (this.loop) void this.poll(this.loop);
    };
    const unsubscribe = this.mounted.subscribe(follow);
    follow();
    return () => { unsubscribe(); this.loop?.controller.abort(); this.loop = undefined; for (const running of this.running.values()) running.abort(); };
  }

  async poll({ sessionId, controller }) {
    while (!controller.signal.aborted) {
      let command;
      try { command = await this.call(sessionId, 'agentNext', {}, controller.signal); }
      catch {
        if (controller.signal.aborted) return;
        await new Promise(resolve => this.scheduler.setTimeout(resolve, this.retryMs));
        continue;
      }
      if (command) void this.serve(sessionId, command);
    }
  }

  async serve(sessionId, command) {
    if (command.kind === 'cancel') { this.running.get(command.target)?.abort(); return; }
    // One-way: the Host saved this file for the AI, so the tab must not treat the new revision as foreign.
    if (command.kind === 'saved') {
      const controller = this.controllers.byAddress.get(this.address(sessionId, command.path, command));
      controller?.markSaved(command.version, command.editRevision);
      // A save marks the end of a generation: an idle "generating · Stop" goes now instead of after its spell.
      const idle = controller?.state.generating;
      if (idle && !idle.running && !idle.stopped) { this.scheduler.clearTimeout(idle.timer); controller.set({ generating: undefined }); }
      return;
    }
    const abort = new AbortController();
    this.running.set(command.requestId, abort);
    // The Host gives up on a command whose Client goes quiet; report while the work (or the user) is pending.
    let beat;
    const alive = () => { beat = this.scheduler.setTimeout(() => { void this.call(sessionId, 'agentAlive', { requestId: command.requestId }).catch(() => {}); alive(); }, this.aliveMs); };
    alive();
    let reply;
    try { reply = { ok: true, result: await this.run(sessionId, command, abort.signal) }; }
    catch (error) {
      const code = typeof error?.code === 'string' && error.code.startsWith('officesdk/') ? error.code.slice('officesdk/'.length) : abort.signal.aborted ? 'cancelled' : 'editor-failed';
      reply = { ok: false, code, message: error instanceof Error ? error.message : String(error) };
      // The chat row may not see this code in the tool's error text (see tool-views failureOf).
      if (command.callId && this.views) this.views.record([callKey(command.callId)], { ...this.views.get(callKey(command.callId)), failure: code });
    } finally { this.scheduler.clearTimeout(beat); this.running.delete(command.requestId); }
    await this.call(sessionId, 'agentReply', { requestId: command.requestId, ...reply }).catch(() => {});
  }

  address(sessionId, path, command) {
    // Every Host command names its platform; remembered so a chat row opened with no tab left still spells it that way.
    if (command?.platform) this.lastPlatform = command.platform;
    const platform = command?.platform ?? (typeof this.platform === 'function' ? this.platform() : this.platform);
    return fileAddress(sessionId, path, platform);
  }

  /** The tab for `path`, open in the mode the AI needs and rendered. */
  async document(sessionId, path, signal, command) {
    const format = formatOf(path);
    const mode = AGENT_FORMATS[format];
    if (!mode) throw failure('unsupported-operation', `AI editing supports ${Object.keys(AGENT_FORMATS).join(', ')} files for now.`);
    const address = this.address(sessionId, path, command);
    this.openResource(address);
    const controller = await this.controllers.wait(address, signal);
    await controller.ensureMode(mode, signal);
    return { controller, format };
  }

  async run(sessionId, command, signal) {
    const { controller, format } = await this.document(sessionId, command.path, signal, command);
    // selection_rewrite: the user's own request, so no approval card; the page checks the selection still holds.
    if (command.kind === 'rewrite') {
      const { selection, text, values } = command;
      return controller.agent('rewrite', { ...selection, ...(text !== undefined && { text }), ...(values !== undefined && { values }) }, signal);
    }
    // Word has its own operations, card rows and no AI undo (client/document-agent.mjs); below is the presentation path.
    if (format === 'docx') return runDocumentCommand(this, sessionId, command, controller, signal);
    if (format === 'xlsx') return runSpreadsheetCommand(this, sessionId, command, controller, signal);
    if (command.kind === 'read') {
      const deck = await controller.agent('inspect', {}, signal);
      const result = { format, path: command.path, version: controller.savedVersion, unsaved: controller.state.dirty === true, editRevision: controller.editRevision, deck: summarizeDeck(deck) };
      // The chat row reads this when the result it shows was cut short.
      if (command.callId) this.views?.record([callKey(command.callId)], { outline: readOutline(result.deck), unsaved: result.unsaved });
      return result;
    }
    if (command.kind === 'undo') {
      const result = await controller.agent('undo', { editId: command.editId }, signal);
      controller.set({ edits: (controller.state.edits ?? []).filter(edit => edit.editId !== command.editId) });
      return { ...result, unsaved: true };
    }
    // Before the Host saves a Kit document for the AI, the editor hands over what it still holds.
    if (command.kind === 'flush') { await controller.flushEditor(); return { editRevision: controller.editRevision }; }
    if (command.kind !== 'edit') throw failure('unsupported-operation', 'Unknown document command.');
    const deck = await controller.agent('inspect', {}, signal);
    let plan;
    try { plan = validatePptxOps(command.ops, deck); }
    catch (error) { throw failure(error.code ?? 'invalid-edit', error.message); }
    // A file the model is generating (the Host decides): no card, the tab says so and offers Stop instead. After the
    // user pressed Stop in this tab the card is back even if the Host has not heard of it yet.
    const auto = command.auto === true && !controller.state.generating?.stopped;
    // For the chat row: the slides this call names, at their pages before the edit, until the edit settles them.
    const record = this.recorder(command, deck);
    if (auto) this.generating(sessionId, command, controller);
    else {
      record?.({ awaiting: true });
      let approved;
      try { approved = await controller.propose({ summary: command.summary, rows: describePptxSteps(plan.steps, deck), danger: plan.danger, reasons: plan.reasons }, signal); }
      finally { record?.({ awaiting: false }); }
      if (!approved) return { declined: true };
    }
    try {
      const result = await controller.agent('apply', { ops: command.ops, editId: command.editId }, signal);
      controller.set({ edits: [...(controller.state.edits ?? []), { editId: command.editId, summary: command.summary }].slice(-20) });
      record?.settle(controller, result);
      return result;
    } finally { if (auto) this.generating(sessionId, command, controller, { idle: true }); }
  }

  /**
   * Records an edit for its chat row (by call id) and for an undo naming it (by edit id): pages before the edit at
   * once, then — without holding up the result — the pages as they are after it. Images stay out of the record.
   */
  recorder(command, deck) {
    const keys = [command.callId && callKey(command.callId), command.editId && editKey(command.editId)];
    if (!this.views || !keys.some(Boolean)) return undefined;
    const ops = (Array.isArray(command.ops) ? command.ops : []).map(op => {
      if (!op || typeof op !== 'object') return op;
      const { imageBase64: _bytes, ...rest } = op;
      return rest;
    });
    let base = { ...editContext(ops, deck), summary: command.summary, path: command.path, ops };
    const record = extra => { base = { ...base, ...extra }; this.views.record(keys, base); };
    record({ awaiting: false });
    record.settle = (controller, result) => {
      void controller.agent('inspect', {}).then(after => {
        const slideIds = Array.isArray(result?.slideIds) ? result.slideIds : [];
        this.views.record(keys, settleContext(base, editContext(ops, after, { slideIds, refs: result?.refs }), slideIds));
      }, () => {});
    };
    return record;
  }

  /**
   * Shows "AI is generating · Stop" in the tab while a batch runs and for `generatingMs` after it, so the user can
   * stop between batches too (a batch of several slides takes the model a while to write; a save ends it at once). Stop cancels the running batch (a PPT batch is undone;
   * a Word batch keeps the steps it finished) and tells the Host to ask for every batch from now on.
   */
  generating(sessionId, command, controller, { idle = false } = {}) {
    const current = controller.state.generating;
    if (current) this.scheduler.clearTimeout(current.timer);
    if (current?.stopped) return;
    const token = {};
    const stop = () => {
      this.scheduler.clearTimeout(token.timer);
      controller.set({ generating: { stopped: true } });
      for (const [requestId, running] of this.running) if (requestId === command.requestId || controller.state.generatingRequest === requestId) running.abort();
      void this.call(sessionId, 'agentStop', { path: command.path }).catch(() => {});
      this.scheduler.setTimeout(() => { if (controller.state.generating?.stopped) controller.set({ generating: undefined }); }, 4_000);
    };
    Object.assign(token, { running: !idle, summary: command.summary, stop });
    if (idle) token.timer = this.scheduler.setTimeout(() => { if (controller.state.generating === token) controller.set({ generating: undefined }); }, this.generatingMs);
    controller.set({ generating: token, generatingRequest: idle ? undefined : command.requestId });
  }
}

/** Controllers by document address, with waiters for tabs that are still mounting. */
export class ControllerRegistry {
  constructor({ scheduler = globalThis, timeoutMs = 30_000 } = {}) {
    Object.assign(this, { scheduler, timeoutMs });
    this.byAddress = new Map();
    this.waiters = new Map();
  }
  add(address, controller) {
    this.byAddress.set(address, controller);
    for (const waiter of this.waiters.get(address) ?? []) waiter(controller);
    this.waiters.delete(address);
    return () => { if (this.byAddress.get(address) === controller) this.byAddress.delete(address); };
  }
  wait(address, signal) {
    const current = this.byAddress.get(address);
    if (current) return Promise.resolve(current);
    return new Promise((resolve, reject) => {
      const list = this.waiters.get(address) ?? new Set();
      const done = controller => { this.scheduler.clearTimeout(timer); resolve(controller); };
      const timer = this.scheduler.setTimeout(() => { list.delete(done); reject(failure('no-editor', 'The document did not open in the side panel.')); }, this.timeoutMs);
      list.add(done);
      this.waiters.set(address, list);
      signal?.addEventListener('abort', () => { list.delete(done); this.scheduler.clearTimeout(timer); reject(signal.reason ?? failure('cancelled', 'Cancelled.')); }, { once: true });
    });
  }
}
