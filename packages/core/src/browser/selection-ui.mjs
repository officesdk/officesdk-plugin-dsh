/**
 * The selection rewrite on a Kit page (spec 261109-selection-ai-edit §2, §5, §6): after the user selects text (Word) or a
 * block of cells (Sheet), a small "AI rewrite" button appears next to it; it opens a one-line box for what to change.
 * Sending hands the request to the parent (the DSH tab), which puts it into the conversation as the user's message;
 * the selection then shows as waiting until the model's selection_rewrite arrives, and after the rewrite a Keep / Undo
 * bar offers to take it back. Everything drawn is the plugin's own elements on the page body with inline styles: no
 * CSS is injected and the SDK's DOM is only read.
 *
 * createSelectionUi draws; createSelectionFeature decides (and is tested with a stand-in for the drawing).
 */
export const SELECTION_TEXTS = {
  zh: {
    chip: 'AI 改写', chipTitle: '让 AI 只改选中的内容（⌘J）', placeholder: '告诉 AI 怎么改，例如：改得更正式', send: '发送',
    suggestions: ['更正式', '更简洁', '改错别字', '翻成英文'], pending: '已发给 AI', cancel: '取消这次修改', done: 'AI 已改写', keep: '保留', revert: '撤回',
    reverting: '正在撤回…', reverted: '已撤回', revertFailed: '内容已被改动，无法撤回', sending: '正在发送…', toDraft: '写进对话输入框', drafted: '已写进对话输入框，请在对话里发送',
    reasons: {
      running: 'AI 正在回复，完成后再试', draft: '对话输入框里还有内容，请先发送或清空', 'no-composer': '请先打开这个文档所在的对话', failed: '发送失败，请重试',
      table: '暂不支持改写表格里的文字', 'too-long': '选中的内容太长（最多 2000 字、20 段）', multiple: '只支持一块连续单元格', 'not-cells': '只支持一块连续单元格（不支持整行、整列）',
      'too-many': '最多选 200 个单元格', merged: '暂不支持含合并单元格的区域', busy: 'AI 正在编辑，稍后再试', 'read-only': '只读打开时不能改写',
    },
  },
  en: {
    chip: 'AI rewrite', chipTitle: 'Let AI change only the selection (⌘J)', placeholder: 'Tell AI what to change, e.g. make it more formal', send: 'Send',
    suggestions: ['More formal', 'More concise', 'Fix typos', 'Translate to Chinese'], pending: 'Sent to AI', cancel: 'Cancel this change', done: 'AI rewrote this', keep: 'Keep', revert: 'Undo',
    reverting: 'Undoing…', reverted: 'Undone', revertFailed: 'The text changed since; cannot undo', sending: 'Sending…', toDraft: 'Put it in the chat box', drafted: 'Put in the chat box; send it there',
    reasons: {
      running: 'AI is replying; try again when it is done', draft: 'The chat box has text in it; send or clear it first', 'no-composer': 'Open the conversation of this document first', failed: 'Could not send; try again',
      table: 'Text inside tables cannot be rewritten yet', 'too-long': 'The selection is too long (at most 2000 characters, 20 paragraphs)', multiple: 'Select one block of cells', 'not-cells': 'Select one block of cells (not whole rows or columns)',
      'too-many': 'Select at most 200 cells', merged: 'Ranges with merged cells are not supported yet', busy: 'AI is editing; try again shortly', 'read-only': 'Open for editing to rewrite',
    },
  },
};
/** ⌘J / Ctrl+J while a selection can be rewritten (Writer leaves it unbound; measured in tooling/suite/e2e-selection.mjs). */
export const isShortcut = event => (event.metaKey || event.ctrlKey) && !event.altKey && String(event.key).toLowerCase() === 'j';
/** Enter sends unless an input method is still composing. */
export const isSend = event => event.key === 'Enter' && !event.shiftKey && !event.isComposing && event.keyCode !== 229;
const MARK = 'data-shimo-selection';
const FONT = '-apple-system, BlinkMacSystemFont, "PingFang SC", "Microsoft YaHei", system-ui, sans-serif';
const Z = '2147483000';
const BLUE = '#2f6bff';

/** Where a box `size` goes next to `point` without leaving the window. */
export function placeNear(point, size, view) {
  const width = view?.innerWidth ?? 1024, height = view?.innerHeight ?? 768;
  let left = point.x + 8, top = point.y + 8;
  if (left + size.width > width - 8) left = Math.max(8, width - size.width - 8);
  if (top + size.height > height - 8) top = Math.max(8, point.y - size.height - 8);
  return { left: Math.round(left), top: Math.round(top) };
}

export function createSelectionUi(document, { locale = 'zh' } = {}) {
  const view = document.defaultView;
  const texts = SELECTION_TEXTS[locale] ?? SELECTION_TEXTS.zh;
  const make = (tag, style = {}, attrs = {}) => {
    const node = document.createElement(tag);
    node.setAttribute(MARK, attrs.mark ?? '');
    for (const [name, value] of Object.entries(attrs)) if (name !== 'mark') node.setAttribute(name, value);
    Object.assign(node.style, { boxSizing: 'border-box', font: `13px/20px ${FONT}`, ...style });
    return node;
  };
  const show = (node, on) => { node.hidden = !on; node.style.display = on ? (node.dataset.display ?? 'flex') : 'none'; };
  const button = (label, style = {}, attrs = {}) => {
    const node = make('button', { border: '1px solid #d0d5dd', borderRadius: '14px', background: '#fff', color: '#1f2329', padding: '2px 10px', cursor: 'pointer', whiteSpace: 'nowrap', ...style }, { type: 'button', ...attrs });
    node.textContent = label;
    return node;
  };

  const chip = button(`✦ ${texts.chip}`, { position: 'fixed', zIndex: Z, display: 'none', background: BLUE, color: '#fff', borderColor: BLUE, boxShadow: '0 2px 8px rgba(0,0,0,.18)', fontWeight: '500' }, { mark: 'chip', title: texts.chipTitle });
  chip.dataset.display = 'block';

  const panel = make('div', { position: 'fixed', zIndex: Z, display: 'none', flexDirection: 'column', gap: '6px', width: '340px', padding: '8px', background: '#fff', border: '1px solid #d0d5dd', borderRadius: '10px', boxShadow: '0 8px 24px rgba(0,0,0,.16)' }, { mark: 'panel', role: 'dialog' });
  const row = make('div', { display: 'flex', gap: '6px' });
  const input = make('input', { flex: '1', minWidth: '0', height: '30px', padding: '4px 8px', border: `1px solid ${BLUE}`, borderRadius: '6px', outline: 'none' }, { mark: 'input', type: 'text', placeholder: texts.placeholder, maxlength: '500', 'aria-label': texts.chip });
  const send = button(texts.send, { background: BLUE, color: '#fff', borderColor: BLUE, borderRadius: '6px' }, { mark: 'send' });
  row.append(input, send);
  const ideas = make('div', { display: 'flex', gap: '6px', flexWrap: 'wrap' });
  const suggestionButtons = texts.suggestions.map(text => button(text, { fontSize: '12px' }, { mark: 'suggestion' }));
  ideas.append(...suggestionButtons);
  const hint = make('div', { display: 'none', gap: '8px', alignItems: 'center', color: '#b42318', fontSize: '12px' }, { mark: 'hint', role: 'status' });
  const hintText = make('span');
  const toDraft = button(texts.toDraft, { fontSize: '12px', display: 'none' }, { mark: 'to-draft' });
  hint.append(hintText, toDraft);
  panel.append(row, ideas, hint);

  const frame = make('div', { position: 'fixed', zIndex: Z, display: 'none', pointerEvents: 'none', border: `1.5px dashed ${BLUE}`, borderRadius: '4px', background: 'rgba(47,107,255,.06)' }, { mark: 'pending' });
  frame.dataset.display = 'block';
  const tag = make('div', { position: 'fixed', zIndex: Z, display: 'none', alignItems: 'center', gap: '4px', padding: '0 4px 0 8px', height: '24px', fontSize: '12px', color: '#fff', background: BLUE, borderRadius: '12px' }, { mark: 'pending-tag' });
  const tagText = make('span');
  tagText.textContent = texts.pending;
  const tagCancel = button('×', { border: 'none', background: 'transparent', color: '#fff', padding: '0 4px', fontSize: '14px' }, { mark: 'pending-cancel', title: texts.cancel, 'aria-label': texts.cancel });
  tag.append(tagText, tagCancel);

  const bar = make('div', { position: 'fixed', zIndex: Z, display: 'none', alignItems: 'center', gap: '6px', padding: '4px 6px 4px 10px', fontSize: '12px', background: '#fff', border: `1px solid ${BLUE}`, borderRadius: '14px', boxShadow: '0 2px 8px rgba(0,0,0,.14)' }, { mark: 'result', role: 'status' });
  const barText = make('span');
  const keep = button(`✓ ${texts.keep}`, { fontSize: '12px' }, { mark: 'keep' });
  const revert = button(`↶ ${texts.revert}`, { fontSize: '12px' }, { mark: 'revert' });
  bar.append(barText, keep, revert);
  document.body.append(chip, panel, frame, tag, bar);

  const handlers = {};
  const on = (node, type, run) => node.addEventListener(type, event => { event.stopPropagation(); run(event); });
  on(chip, 'click', () => handlers.open?.());
  on(send, 'click', () => handlers.submit?.(input.value));
  for (const item of suggestionButtons) on(item, 'click', () => handlers.submit?.(item.textContent));
  on(toDraft, 'click', () => handlers.submit?.(input.value, { draftOnly: true }));
  on(input, 'keydown', event => {
    if (isSend(event)) { event.preventDefault(); handlers.submit?.(input.value); }
    else if (event.key === 'Escape') { event.preventDefault(); handlers.close?.(); }
  });
  on(tagCancel, 'click', () => handlers.cancelPending?.());
  on(keep, 'click', () => handlers.keep?.());
  on(revert, 'click', () => handlers.revert?.());
  // Presses on our own elements must not reach the editor (it would move its selection away).
  for (const node of [chip, panel, tag, bar]) for (const type of ['pointerdown', 'mousedown', 'pointerup', 'mouseup']) node.addEventListener(type, event => event.stopPropagation());
  const place = (node, point) => { const rect = node.getBoundingClientRect(); const at = placeNear(point, { width: rect.width || 120, height: rect.height || 30 }, view); node.style.left = `${at.left}px`; node.style.top = `${at.top}px`; };

  return {
    owns: target => Boolean(target?.closest?.(`[${MARK}]`)),
    handlers,
    showChip(point, { disabled = false, reason } = {}) {
      show(chip, true);
      chip.style.opacity = disabled ? '0.55' : '1';
      chip.title = disabled && reason ? (texts.reasons[reason] ?? texts.chipTitle) : texts.chipTitle;
      chip.setAttribute('aria-disabled', String(disabled));
      place(chip, point);
    },
    hideChip() { show(chip, false); },
    get chipShown() { return chip.style.display !== 'none'; },
    openPanel(point) {
      show(chip, false);
      show(panel, true);
      input.value = '';
      this.hint();
      place(panel, point);
      input.focus();
    },
    closePanel() { show(panel, false); },
    get panelOpen() { return panel.style.display !== 'none'; },
    busy(on) { input.disabled = on; send.disabled = on; for (const item of suggestionButtons) item.disabled = on; if (on) this.hint(texts.sending, { tone: 'info' }); },
    hint(text, { tone = 'error', draft = false } = {}) {
      show(hint, Boolean(text));
      hint.style.color = tone === 'error' ? '#b42318' : '#475467';
      hintText.textContent = text ?? '';
      toDraft.style.display = draft ? 'inline-block' : 'none';
    },
    showPending(box) {
      if (!box) { show(frame, false); show(tag, false); return; }
      show(frame, true);
      Object.assign(frame.style, { left: `${box.left - 3}px`, top: `${box.top - 3}px`, width: `${box.width + 6}px`, height: `${box.height + 6}px` });
      show(tag, true);
      Object.assign(tag.style, { left: `${box.left - 3}px`, top: `${Math.max(4, box.top - 30)}px` });
    },
    showResult(point, text = texts.done) {
      barText.textContent = text;
      show(bar, true);
      keep.style.display = revert.style.display = text === texts.done ? 'inline-block' : 'none';
      place(bar, point);
    },
    hideResult() { show(bar, false); },
    get resultShown() { return bar.style.display !== 'none'; },
    texts,
    dispose() { for (const node of [chip, panel, frame, tag, bar]) node.remove(); },
  };
}

/**
 * The logic around the drawing. `agent` is the suite's AI agent (readSelection / rewriteSelection / revertSelection /
 * applying / locate); `notify(event, payload)` posts to the parent tab. Returns the handlers kit-page serves:
 * `composer(state)`, `status(reply)` and `rewrite(args, signal)`.
 */
export function createSelectionFeature({ ui, agent, notify, window, document, editable = () => true, debounceMs = 220, statusTimeoutMs = 15_000, schedule = (run, ms) => window.setTimeout(run, ms), unschedule = id => window.clearTimeout(id) }) {
  const texts = ui.texts;
  let current, timer, lastPoint, composer = { canSend: true }, pending, result, sending, counter = 0;
  const waiting = new Map();
  const reasonText = reason => texts.reasons[reason] ?? texts.reasons.failed;
  const anchorPoint = selection => {
    const box = selection?.box ?? agent.locate?.(selection);
    if (lastPoint && Date.now() - lastPoint.at < 1500) return lastPoint;
    return box ? { x: box.left + box.width, y: box.top + box.height } : lastPoint ?? { x: 40, y: 40 };
  };
  const refresh = async () => {
    timer = undefined;
    if (!editable() || agent.applying || ui.panelOpen || sending) return;
    let selection;
    try { selection = await agent.readSelection(); } catch { selection = undefined; }
    if (!editable() || agent.applying || ui.panelOpen || sending) return;
    if (!selection) { current = undefined; ui.hideChip(); return; }
    const point = anchorPoint(selection);
    if (selection.unsupported) ui.showChip(point, { disabled: true, reason: selection.unsupported });
    else ui.showChip(point, { disabled: composer.canSend === false && composer.reason === 'running', reason: composer.reason });
    current = selection.unsupported ? { unsupported: selection.unsupported, point } : { ...selection, point };
  };
  const later = () => { if (timer !== undefined) unschedule(timer); timer = schedule(() => { void refresh(); }, debounceMs); };
  // The editor may hold the keyboard (the Sheet takes focus back from anything else): it lets go while the box is open.
  const openPanel = point => { void agent.focusEditor?.(false); ui.openPanel(point); };
  const closePanel = () => { if (!ui.panelOpen) return; ui.closePanel(); void agent.focusEditor?.(true); };
  const open = () => {
    if (!current) return;
    if (current.unsupported) { openPanel(current.point); ui.hint(reasonText(current.unsupported)); return; }
    if (composer.canSend === false && composer.reason === 'running') { ui.showChip(current.point, { disabled: true, reason: 'running' }); return; }
    openPanel(current.point);
  };
  const close = () => { closePanel(); later(); };
  const submit = async (value, { draftOnly = false } = {}) => {
    const instruction = String(value ?? '').trim();
    if (!instruction || !current || current.unsupported || sending) return;
    const requestId = `q${Date.now().toString(36)}${(counter++).toString(36)}`;
    const { point, box, view, ...request } = current;
    sending = { requestId, selection: current };
    ui.busy(true);
    let reply;
    try {
      reply = await new Promise(resolve => {
        const timeout = schedule(() => { waiting.delete(requestId); resolve({ ok: false, reason: 'failed' }); }, statusTimeoutMs);
        waiting.set(requestId, value => { unschedule(timeout); resolve(value); });
        notify('selection-request', { requestId, ...request, instruction, ...(draftOnly && { draftOnly: true }) });
      });
    } finally { ui.busy(false); }
    const selection = sending.selection;
    sending = undefined;
    if (!reply?.ok) { ui.hint(reasonText(reply?.reason), { draft: reply?.reason === 'draft' && reply?.drafted !== false }); return; }
    closePanel();
    ui.hideChip();
    if (reply.drafted) { ui.showResult(selection.point, texts.drafted); schedule(() => ui.hideResult(), 4000); }
    pending = { id: reply.id, selection };
    ui.showPending(agent.locate?.(selection) ?? selection.box);
  };
  const cancelPending = () => {
    if (!pending) return;
    notify('selection-update', { selection: pending.id, state: 'cancelled' });
    pending = undefined;
    ui.showPending(undefined);
  };
  const keep = () => {
    if (!result) return;
    notify('selection-update', { selection: result.id, state: 'kept' });
    result = undefined;
    ui.hideResult();
  };
  const revert = async () => {
    if (!result || agent.applying) return;
    const done = result;
    ui.showResult(done.point, texts.reverting);
    try {
      await agent.revertSelection(done.revert);
      result = undefined;
      notify('selection-update', { selection: done.id, state: 'reverted' });
      ui.showResult(done.point, texts.reverted);
    } catch (error) {
      result = undefined;
      console.warn('officesdk: selection undo refused', error?.code ?? '', error?.message ?? error);
      ui.showResult(done.point, texts.revertFailed);
    }
    schedule(() => { if (!result) ui.hideResult(); }, 2500);
  };
  Object.assign(ui.handlers, { open, close, submit, cancelPending, keep, revert });

  const pointer = event => {
    if (ui.owns(event.target)) return;
    if (event.type === 'pointerup') lastPoint = { x: event.clientX, y: event.clientY, at: Date.now() };
    if (event.type === 'pointerdown') { if (!ui.panelOpen) ui.hideChip(); return; }
    if (ui.panelOpen) closePanel();
    later();
  };
  const key = event => {
    if (ui.owns(event.target)) return;
    if (event.type === 'keydown') {
      if (isShortcut(event) && current && !ui.panelOpen) { event.preventDefault(); event.stopImmediatePropagation?.(); open(); return; }
      // Typing in the document: the Keep / Undo bar counts as kept, the button goes until the next selection.
      if (!['Shift', 'Meta', 'Control', 'Alt'].includes(event.key)) { if (result) keep(); if (!event.shiftKey) ui.hideChip(); }
      return;
    }
    if (event.shiftKey || event.key === 'Shift' || /^Arrow|^(Home|End)$/.test(event.key)) later();
  };
  const wheel = () => { if (!ui.panelOpen) ui.hideChip(); };
  const listeners = [['pointerdown', pointer], ['pointerup', pointer], ['keydown', key], ['keyup', key], ['wheel', wheel]];
  for (const [type, run] of listeners) document.addEventListener(type, run, { capture: true, passive: type !== 'keydown' });

  return {
    /** The parent's word on whether the conversation can take a message now. */
    composer(state) { composer = { canSend: state?.canSend !== false, ...(state?.reason && { reason: state.reason }) }; if (ui.chipShown && current) later(); return { ok: true }; },
    /** The parent's answer to a selection-request. */
    status(reply) {
      // A send that failed after the page heard it went: the selection waits for nothing now.
      if (typeof reply?.dropped === 'string') { if (pending?.id === reply.dropped) { pending = undefined; ui.showPending(undefined); } return { ok: true }; }
      const resolve = waiting.get(reply?.requestId); waiting.delete(reply?.requestId); resolve?.(reply); return { ok: Boolean(resolve) };
    },
    /** The model's selection_rewrite, delivered by the parent: rewrite, then offer Keep / Undo. */
    async rewrite(args, signal) {
      ui.hideChip();
      closePanel();
      if (pending?.id === args.id) { pending = undefined; ui.showPending(undefined); }
      if (result) keep();
      const outcome = await agent.rewriteSelection(args, signal);
      const point = (() => { const box = agent.locate?.(outcome.located); return box ? { x: box.left + box.width, y: box.top + box.height } : lastPoint ?? { x: 40, y: 40 }; })();
      if (outcome.revert) { result = { id: args.id, revert: outcome.revert, point }; ui.showResult(point); }
      return outcome.reply;
    },
    /** The selection waiting for the model, if any (for tests and the page's own checks). */
    get pending() { return pending; },
    get result() { return result; },
    refresh,
    dispose() { for (const [type, run] of listeners) document.removeEventListener(type, run, { capture: true }); if (timer !== undefined) unschedule(timer); ui.dispose(); },
  };
}
