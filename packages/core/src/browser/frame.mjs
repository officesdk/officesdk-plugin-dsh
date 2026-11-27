/** Frame protocol v1: iframe events and parent commands, each bound to one session id and revision. */
export const FRAME_VERSION = 1;
/**
 * `selection-request` / `selection-update`: the user's selection rewrite request and its later fate (cancelled, kept,
 * reverted). The envelope owns `id` (the view): a selection id travels as `selection`.
 */
export const EVENTS = Object.freeze(['ready', 'failed', 'dirty', 'saved', 'conflict', 'agent-result', 'save-request', 'history', 'selection-request', 'selection-update', 'document-ui', 'open-website', 'press']);
export const COMMANDS = Object.freeze(['save', 'flush', 'rebind', 'agent']);
/**
 * Operations an `agent` command may ask of an editor; adapters implement the subset they support. `undo` takes back
 * an AI edit; `history` is the toolbar's undo/redo of the editor's own history; `print` opens the print dialog.
 * Selection rewrite: `composer` tells the page whether the conversation can take a message, `selectionStatus` answers a
 * selection-request, `rewrite` carries the model's selection_rewrite.
 */
export const AGENT_OPS = Object.freeze(['inspect', 'preview', 'apply', 'undo', 'cancel', 'flush', 'snapshot', 'history', 'print', 'slideshow', 'composer', 'submitChatRegion', 'readSelection', 'documentUi', 'selectionStatus', 'rewrite', 'changeDetail']);
const REQUEST_ID = /^[A-Za-z0-9_-]{1,64}$/;

/** Parent side: only the session's own frame, origin, id and current revision are accepted. */
export function acceptFrameMessage(event, { frame, origin, id, revision }) {
  if (!frame || event.source !== frame.contentWindow || event.origin !== origin) return undefined;
  const message = event.data;
  if (!message || typeof message !== 'object' || message.v !== FRAME_VERSION || message.id !== id || message.revision !== revision) return undefined;
  const prefix = typeof message.type === 'string' && ['officesdk:'].find(candidate => message.type.startsWith(candidate));
  if (!prefix) return undefined;
  const name = message.type.slice(prefix.length);
  if (name === 'agent-result' && !REQUEST_ID.test(message.requestId ?? '')) return undefined;
  return EVENTS.includes(name) ? { ...message, event: name } : undefined;
}

export function frameCommand({ id, revision, command, newRevision, requestId, op, args }) {
  if (!COMMANDS.includes(command)) throw new Error('Unknown frame command');
  if (command === 'agent' && (!REQUEST_ID.test(requestId ?? '') || !AGENT_OPS.includes(op))) throw new Error('Invalid agent command');
  return { type: 'officesdk:command', v: FRAME_VERSION, id, revision, command,
    ...(command === 'rebind' && { newRevision }), ...(command === 'agent' && { requestId, op, args: args ?? {} }) };
}

/** Iframe side: posts events to the parent and applies rebinds before any handler sees a command. */
export function createFrameChannel({ window, id, revision, parentOrigin, signal }) {
  let current = revision;
  const handlers = new Set();
  const target = parentOrigin && parentOrigin !== 'null' ? parentOrigin : window.location.origin;
  const notify = (event, payload = {}) => {
    if (!EVENTS.includes(event)) throw new Error('Unknown frame event');
    if (window.parent !== window) window.parent.postMessage({ ...payload, type: `officesdk:${event}`, v: FRAME_VERSION, id, revision: current }, target);
  };
  const receive = event => {
    if (event.source !== window.parent || event.origin !== target) return;
    const message = event.data;
    if (!message || message.type !== 'officesdk:command' || message.v !== FRAME_VERSION || message.id !== id || message.revision !== current || !COMMANDS.includes(message.command)) return;
    if (message.command === 'rebind') {
      if (!Number.isSafeInteger(message.newRevision) || message.newRevision < 0) return;
      current = message.newRevision;
    }
    if (message.command === 'agent' && (!REQUEST_ID.test(message.requestId ?? '') || !AGENT_OPS.includes(message.op))) return;
    for (const handler of handlers) handler(message);
  };
  window.addEventListener('message', receive);
  signal?.addEventListener('abort', () => window.removeEventListener('message', receive), { once: true });
  return { notify, onCommand(handler) { handlers.add(handler); return () => handlers.delete(handler); }, get revision() { return current; } };
}

/**
 * Iframe side of `agent` commands: one handler per op, one result event per request.
 * Errors cross as a code and fixed text only; an editor without the op answers `unsupported-operation`.
 */
export function serveAgent(channel, ops, { signal } = {}) {
  const running = new Map();
  return channel.onCommand(async message => {
    if (message.command !== 'agent') return;
    if (message.op === 'cancel') { running.get(message.args?.target)?.abort(); return; }
    const reply = payload => channel.notify('agent-result', { requestId: message.requestId, ...payload });
    const handler = ops[message.op];
    if (!handler) return reply({ ok: false, code: 'unsupported-operation', message: 'This editor cannot do that.' });
    const controller = new AbortController();
    running.set(message.requestId, controller);
    const abort = () => controller.abort();
    signal?.addEventListener('abort', abort, { once: true });
    try { reply({ ok: true, result: await handler(message.args ?? {}, controller.signal) }); }
    catch (error) { reply({ ok: false, code: typeof error?.code === 'string' ? error.code : 'editor-failed', message: error instanceof Error ? error.message : String(error) }); }
    finally { running.delete(message.requestId); signal?.removeEventListener('abort', abort); }
  });
}

/**
 * The frame may open the print dialog (`allow-modals`), which would also let the editor raise alert/confirm/prompt
 * and leave-page prompts: they keep the answers the sandbox gave before, without a dialog.
 */
export function quietModals(window) {
  const quiet = (name, answer) => (...args) => { console.info(`officesdk: ${name} suppressed`, ...args); return answer; };
  for (const [name, answer] of [['alert', undefined], ['confirm', false], ['prompt', null]]) {
    try { window[name] = quiet(name, answer); } catch {}
  }
  try { Object.defineProperty(window, 'onbeforeunload', { configurable: true, get: () => null, set() {} }); } catch {}
  try { window.addEventListener('beforeunload', event => event.stopImmediatePropagation(), { capture: true }); } catch {}
}
