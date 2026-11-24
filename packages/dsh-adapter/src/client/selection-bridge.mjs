import { runningOf } from './style-flows.mjs';
import { documentQuoteSummary } from './document-context.mjs';

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
  const entriesChanged = new Set();
  const changed = sessionId => { for (const listener of listeners) listener(sessionId); };
  return {
    get: sessionId => entries.get(sessionId)?.at(-1),
    state: sessionId => composerState(entries.get(sessionId)?.at(-1)),
    /** Registers or refreshes one mount; returns its removal. */
    set(sessionId, token, entry) {
      const list = entries.get(sessionId) ?? [];
      const at = list.findIndex(item => item.token === token);
      const prev=list.at(-1);
      if (prev?.token===token && Object.keys(entry).every(key=>entry[key]===prev[key])) return;
      if (at >= 0) list[at] = { ...entry, token }; else list.push({ ...entry, token });
      entries.set(sessionId, list);
      for (const listener of entriesChanged) listener(sessionId);
      if (JSON.stringify(composerState(prev)) !== JSON.stringify(composerState(list.at(-1)))) changed(sessionId);
    },
    remove(sessionId, token) {
      const list = entries.get(sessionId);
      const at = list?.findIndex(item => item.token === token) ?? -1;
      if (at < 0) return;
      list.splice(at, 1);
      if (!list.length) entries.delete(sessionId);
      for (const listener of entriesChanged) listener(sessionId);
      changed(sessionId);
    },
    subscribeEntries(listener) { entriesChanged.add(listener); return () => entriesChanged.delete(listener); },
    subscribe(listener) { listeners.add(listener); return () => listeners.delete(listener); },
  };
}

/** The hidden slot component: keeps this input box's entry current. */
export function createComposerProbe({ React, registry }) {
  const { useEffect, useRef } = React;
  return function SelectionComposer({ sessionId, inputActions, useInput, useSessionStatus, useProjection, useChat, session, input: zoneInput }) {
    const token = useRef(undefined), marker=useRef(null);
    token.current ??= {};
    const input = useInput?.(value => value) ?? zoneInput;
    const running = runningOf({ sessionId, useSessionStatus, session });
    const draft = typeof input?.draft === 'string' ? input.draft : '';
    const title = useProjection?.('title') ?? '';
    const snapshot = useChat?.(value=>value);
    const chat = snapshot?.legacy;
    const attachments = Array.isArray(input?.attachmentIds) ? input.attachmentIds.length : 0;
    useEffect(() => {
      if (!sessionId || marker.current?.closest('.office-chat-decisions')) return undefined;
      registry.set(sessionId, token.current, { actions: inputActions, draft, attachments, running, title, chat, snapshot });
      return undefined;
    });
    useEffect(() => () => { if (sessionId) registry.remove(sessionId, token.current); }, [sessionId]);
    return React.createElement('span',{ref:marker,hidden:true});
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
  const reportFailure = error => controller.set?.({parallelError:error?.message || (typeof error==='string' && error) || (lang==='en'?'Could not send. Please try again.':'发送失败，请重试。')});
  if (controller.state?.parallelError) controller.set?.({parallelError:undefined});
  if (!entry || state.reason === 'no-composer') return reply({ ok: false, reason: 'no-composer' });
  if (message.kind === 'region-batch') {
    const parallel = message.parallel === true || state.reason === 'running';
    const ownsDraft = current => message.fromChat === true && !current?.attachments && current?.draft?.trim() === message.instruction;
    if (!state.canSend && !parallel && !ownsDraft(entry)) return reply({ ok: false, reason:state.reason });
    const instruction = String(message.instruction ?? '').trim();
    const refs = message.references;
    if (!instruction || instruction.length>500 || !Array.isArray(refs) || refs.length!==1) return reply({ok:false, reason:'failed'});
    const ids = [], parts = [];
    let wrote = false;
    const cleanup = () => Promise.allSettled(ids.map(item=>controller.call('selectionUpdate',{id:item.id,selectionState:'cancelled'})));
    const path = controller.source.path;
    try {
      for (const ref of refs) {
        if (ref.kind==='document-context') { parts.push(String(ref.original).slice(0,2000)); continue; }
        const created = await controller.call('selectionCreate',{path, selectionRequest:{
          kind:ref.kind,anchor:ref.anchor,original:ref.original,trackRange:true,...(ref.cells && {cells:ref.cells}),instruction
        },selectionLanguage:lang});
        if (!created?.id || !created.message) throw new Error('Missing selection');
        ids.push({number:ref.number,id:created.id});
        parts.push(created.message);
      }
      if ((parallel || registry.get(sessionId)?.running) && ids.length===1) {
        const id=ids[0].id;
        controller.set?.({parallelError:undefined});
        await reply({ok:true,ids,independent:true});
        try {
          const result=await controller.call('selectionParallel',{id});
          if (result?.error) {
            reportFailure(result.error);
            await cleanup();
            await reply({dropped:id});
          }
        } catch (error) {
          await cleanup();
          const missingRoute = /\/shimoPreview\/selectionParallel.*HTTP 404\b/i.test(error.message ?? '');
          const detail = missingRoute
            ? (lang === 'en' ? 'Restart DSH to load parallel editing, then select this passage and send again.' : '请重启 DSH 以加载并行编辑，再选中这段文字重新发送。')
            : error.message;
          controller.set?.({parallelError:detail});
          await reply({dropped:id});
        }
        return;
      }
      if (!composerState(registry.get(sessionId)).canSend && !ownsDraft(registry.get(sessionId))) { await cleanup(); return reply({ok:false,reason:'draft'}); }
      const target = lang==='en' ? 'Editing target: '+path : '编辑目标：'+path;
      const quote = [target,...parts].join('\n\n');
      const label = refs.map(r=>Array.isArray(r.original)?r.original.join(' · '):r.original).filter(Boolean).join(' · ') || path;
      entry.actions.setDraft(instruction);
      if (!entry.actions.insertQuote?.(quote,documentQuoteSummary(quote,path,label))) entry.actions.setDraft([quote,instruction].join('\n\n'));
      wrote = true;
      await new Promise(resolve=>defer(resolve));
      const submitted=entry.actions.submit();
      void Promise.resolve(submitted).catch(async error => {
        reportFailure(error);
        await cleanup();
        for(const item of ids) void controller.agent('selectionStatus',{requestId:message.requestId,dropped:item.id}).catch(()=>{});
      });
      return reply({ok:true,ids});
    } catch (error) {
      reportFailure(error);
      await cleanup();
      if (wrote) entry.actions.setDraft(entry.draft ?? '');
      return reply({ok:false,reason:'failed'});
    }
  }
  if (message.kind === 'document-context') {
    const instruction = typeof message.instruction === 'string' && message.instruction.length <= 500 ? message.instruction.trim() : '';
    if (!instruction) return reply({ ok: false, reason: 'failed' });
    if (state.reason === 'running') return reply({ ok: false, reason: 'running' });
    const path = controller.source.path;
    const name = lang === 'en' ? `Editing target: ${path}` : `编辑目标：${path}`;
    const selected = typeof message.original === 'string' ? message.original.slice(0, 2_000).trim() : '';
    const target = selected ? `${name}\n${selected}` : name;
    try {
      if (!state.canSend || draftOnly) {
        const position = entry.actions.captureInsertion?.();
        const prompt = `\n${target}\n${instruction}`;
        if (!position || !entry.actions.insertText?.(prompt, position)) return reply({ ok: false, reason: 'draft' });
        return reply({ ok: true, drafted: true });
      }
      entry.actions.setDraft(instruction);
      if (!entry.actions.insertQuote?.(target, documentQuoteSummary(target, path, path))) entry.actions.setDraft(`${target}\n${instruction}`);
      await new Promise(resolve => defer(resolve));
      void Promise.resolve(entry.actions.submit()).catch(reportFailure);
      return reply({ ok: true });
    } catch { return reply({ ok: false, reason: 'failed' }); }
  }
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
  void Promise.resolve(submitted).catch(error => {
    reportFailure(error);
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
    const chatEdit=state.reason==='running'?controller.pendingChatEdit:undefined;
    if(chatEdit)controller.pendingChatEdit=undefined;
    void controller.agent('composer', chatEdit?{...state,chatEdit}:state).catch(() => {});
  }
}
