import { runningOf } from './style-flows.mjs';

/**
 * Selection rewrite, the DSH side (spec 261109-selection-ai-edit §3.3). The editor page asks for a rewrite of the user's
 * selection; the request has to become the user's message in that document's conversation. Plugins reach the input box
 * only from the conversation's own slots (`inputActions`), so a hidden component in `conversation.input.left` registers
 * each mounted input box here by Session: its actions, draft (text and attachments) and whether the Session is running.
 * The document tab then (1) tells its page whether a message can be sent now and (2) sends the request: it checks the
 * input box first, records the selection on the Host, writes the Host's message into the draft and submits it; a failed
 * submit cancels the selection so none is left behind.
 */
export const SELECTION_COMPOSER_ID = 'officesdk-selection-composer';

/** Whether the conversation can take the request as a message now, and why not. */
export function composerState(entry) {
  if (!entry || typeof entry.actions?.setDraft !== 'function' || typeof entry.actions?.submit !== 'function') return { canSend: false, reason: 'no-composer' };
  if (entry.running === true) return { canSend: false, reason: 'running' };
  if ((typeof entry.draft === 'string' && entry.draft.trim()) || entry.attachments > 0) return { canSend: false, reason: 'draft' };
  return { canSend: true };
}

/** The mounted input boxes by Session (the newest mount wins; split panes of one Session share its draft). */
export function createComposerRegistry() {
  const entries = new Map();
  const listeners = new Set();
  const changed = sessionId => { for (const listener of listeners) listener(sessionId); };
  return {
    get: sessionId => entries.get(sessionId)?.at(-1),
    state: sessionId => composerState(entries.get(sessionId)?.at(-1)),
    /** Registers or refreshes one mount; returns its removal. */
    set(sessionId, token, entry) {
      const list = entries.get(sessionId) ?? [];
      const at = list.findIndex(item => item.token === token);
      const before = JSON.stringify(composerState(list.at(-1)));
      if (at >= 0) list[at] = { ...entry, token }; else list.push({ ...entry, token });
      entries.set(sessionId, list);
      if (JSON.stringify(composerState(list.at(-1))) !== before) changed(sessionId);
    },
    remove(sessionId, token) {
      const list = entries.get(sessionId);
      const at = list?.findIndex(item => item.token === token) ?? -1;
      if (at < 0) return;
      list.splice(at, 1);
      if (!list.length) entries.delete(sessionId);
      changed(sessionId);
    },
    subscribe(listener) { listeners.add(listener); return () => listeners.delete(listener); },
  };
}

/** The hidden slot component: keeps this input box's entry current. */
export function createComposerProbe({ React, registry }) {
  const { useEffect, useRef } = React;
  return function SelectionComposer({ sessionId, inputActions, useInput, useSessionStatus, session, input: zoneInput }) {
    const token = useRef(undefined);
    token.current ??= {};
    const input = useInput?.(value => value) ?? zoneInput;
    const running = runningOf({ sessionId, useSessionStatus, session });
    const draft = typeof input?.draft === 'string' ? input.draft : '';
    const attachments = Array.isArray(input?.attachmentIds) ? input.attachmentIds.length : 0;
    useEffect(() => {
      if (!sessionId) return undefined;
      registry.set(sessionId, token.current, { actions: inputActions, draft, attachments, running });
      return undefined;
    });
    useEffect(() => () => { if (sessionId) registry.remove(sessionId, token.current); }, [sessionId]);
    return null;
  };
}

export function registerSelectionComposer(ctx, React, { registry }) {
  const SelectionComposer = createComposerProbe({ React, registry });
  ctx.effect(() => ctx.slots.inject('conversation.input.left', () => ctx.slots.register({ name: 'conversation.input.left', id: SELECTION_COMPOSER_ID, order: 101 }, SelectionComposer)));
  return SelectionComposer;
}

/**
 * Serves one `selection-request` from a document tab's page: answers it with `selectionStatus` ({ ok, id, drafted? }
 * or { ok: false, reason }). `draftOnly`: only put the message into the draft (beside what is there), unsent.
 */
export async function handleSelectionRequest({ controller, message, registry, lang = 'zh', defer = run => setTimeout(run, 0) }) {
  const reply = payload => controller.agent('selectionStatus', { requestId: message.requestId, ...payload }).catch(() => {});
  const sessionId = controller.source?.sessionId;
  const entry = registry.get(sessionId);
  const state = composerState(entry);
  const draftOnly = message.draftOnly === true;
  if (!entry || state.reason === 'no-composer') return reply({ ok: false, reason: 'no-composer' });
  if (!draftOnly && !state.canSend) return reply({ ok: false, reason: state.reason });
  const { kind, anchor, original, cells, instruction } = message;
  let created;
  try {
    created = await controller.call('selectionCreate', { path: controller.source.path, selectionRequest: { kind, anchor, original, ...(cells && { cells }), instruction }, selectionLanguage: lang });
    if (!created?.id || typeof created.message !== 'string') throw new Error('No selection was recorded.');
  } catch { return reply({ ok: false, reason: 'failed' }); }
  const draft = typeof entry.draft === 'string' ? entry.draft : '';
  const cancel = () => { void controller.call('selectionUpdate', { id: created.id, selectionState: 'cancelled' }).catch(() => {}); };
  let submitted;
  try {
    if (draftOnly) {
      entry.actions.setDraft(draft.trim() ? `${draft.replace(/\s+$/, '')}\n\n${created.message}` : created.message);
      return reply({ ok: true, id: created.id, drafted: true });
    }
    entry.actions.setDraft(created.message);
    // The editor takes the text before the send, as the style picker's answer does.
    await new Promise(resolve => defer(resolve));
    submitted = entry.actions.submit();
  } catch {
    cancel();
    try { entry.actions.setDraft(draft); } catch {}
    return reply({ ok: false, reason: 'failed' });
  }
  // DSH's submit settles only when the turn it starts is over: the page hears at once, so the user can still cancel
  // while the model works. A send that fails later cancels the selection (the model never got it), and the page takes
  // its waiting mark down.
  void Promise.resolve(submitted).catch(() => {
    cancel();
    void controller.agent('selectionStatus', { dropped: created.id }).catch(() => {});
  });
  return reply({ ok: true, id: created.id });
}

/** Tells every ready tab of `sessionId` whether its conversation can take a message now. */
export function pushComposerState(controllers, registry, sessionId) {
  const state = registry.state(sessionId);
  for (const controller of controllers) {
    if (controller.source?.sessionId !== sessionId || !controller.state?.ready || controller.state.mode !== 'edit') continue;
    void controller.agent('composer', state).catch(() => {});
  }
}
