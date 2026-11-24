import { recordChange, restoreChangeHistory, sameChangeAnchor } from './change-history.mjs';
import {createChangeLinks} from './change-links.mjs';
import { PROMPT_CSS } from './prompt-style.mjs';

export const SELECTION_TEXTS = {
  zh: {
    chip: '局部编辑', chipTitle: '让 AI 只改选中的内容（⌘J）', placeholder: '输入修改要求…', send: '发送', chat: '编辑文档', close: '收起',
    suggestions: ['更正式', '更简洁', '改错别字', '翻成英文'], pending: '已发给 AI', cancel: '取消这次修改', done: 'AI 已改写', keep: '保留', revert: '撤回',
    reverting: '正在撤回…', reverted: '已撤回', revertFailed: '内容已被改动，无法撤回', sending: '正在发送…', toDraft: '写进对话输入框', drafted: '已写进对话输入框，请在对话里发送',
    reasons: {
      'no-target': '此处没有可继续修改的文字', running: 'AI 正在回复，完成后再试', draft: '对话输入框里还有内容，请先发送或清空', 'no-composer': '请先打开这个文档所在的对话', failed: '发送失败，请重试',
      table: '暂不支持改写表格里的文字', 'too-long': '选中的内容太长（最多 2000 字、20 段）', multiple: '只支持一块连续单元格', 'not-cells': '只支持一块连续单元格（不支持整行、整列）',
      'too-many': '最多选 200 个单元格', merged: '暂不支持含合并单元格的区域', busy: 'AI 正在编辑，稍后再试', 'read-only': '只读打开时不能改写',
    },
  },
  en: {
    chip: 'Local edit', chipTitle: 'Let AI change only the selection (⌘J)', placeholder: 'Enter editing instructions…', send: 'Send', chat: 'Edit document', close: 'Collapse',
    suggestions: ['More formal', 'More concise', 'Fix typos', 'Translate to Chinese'], pending: 'Sent to AI', cancel: 'Cancel this change', done: 'AI rewrote this', keep: 'Keep', revert: 'Undo',
    reverting: 'Undoing…', reverted: 'Undone', revertFailed: 'The text changed since; cannot undo', sending: 'Sending…', toDraft: 'Put it in the chat box', drafted: 'Put in the chat box; send it there',
    reasons: {
      'no-target': 'There is no text here to edit', running: 'AI is replying; try again when it is done', draft: 'The chat box has text in it; send or clear it first', 'no-composer': 'Open the conversation of this document first', failed: 'Could not send; try again',
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
const SEND_ICON = '<svg width="23" height="23" viewBox="0 0 24 24" aria-hidden="true" fill="none" stroke="currentColor" stroke-width="1.65" stroke-linecap="round" stroke-linejoin="round"><path d="M12 19V5m-7 7 7-7 7 7"/></svg>';

const CSS = `
:root[data-office-submitted-selection] rect[data-paint-key^="overlay:selection:"]{visibility:hidden}
.office-region{box-sizing:border-box;position:fixed;z-index:2147483000;color:#202124;font:13px/1.45 system-ui,sans-serif}
.office-region[hidden],.office-region [hidden]{display:none!important}
.office-region button,button.office-region{font:inherit;color:inherit;cursor:pointer}
.office-chip,.office-review button{background:#fff;border:1px solid #dadde0;box-shadow:0 2px 6px #20232712}
button.office-chip{display:inline-flex;align-items:center;justify-content:center;gap:10px;height:28px;padding:0 12px 0 10px;border-radius:9999px;corner-shape:round;color:#292d33;white-space:nowrap;font:400 13px/18px system-ui,sans-serif}
.office-chip .office-selection-sparkles{display:block;width:15px;height:16px;flex:none;overflow:visible;color:#a1a3a6}
.office-selection-sparkles path{transform-box:fill-box;transform-origin:center}
.office-selection-sparkles path:nth-child(1){--spark-delay:240ms;--spark-scale:1.5;--spark-rotate:20deg}
.office-selection-sparkles path:nth-child(2){--spark-delay:120ms;--spark-scale:1.4;--spark-rotate:10deg}
.office-selection-sparkles path:nth-child(3){--spark-delay:0ms;--spark-scale:1.25;--spark-rotate:0deg}
@keyframes office-spark-pop{50%{transform:scale(var(--spark-scale)) rotate(var(--spark-rotate))}}
@media(prefers-reduced-motion:no-preference){.office-chip:not(:disabled):is(:hover,:focus-visible) .office-selection-sparkles path{animation:office-spark-pop .5s ease-in-out var(--spark-delay) 1}}
.office-panel{width:min(232px,calc(100vw - 12px));transform-origin:center bottom;animation:office-open .17s ease-out}
@keyframes office-open{from{transform:scale(.72,.93);opacity:.4}to{transform:none;opacity:1}}
@media(prefers-reduced-motion:reduce){.office-panel{animation:none}}
[data-office-over-change] #editor-root :is(.writer-document-preview__scroll,.writer-document-preview__scroll *){cursor:pointer!important}
:root{--office-writer-selection-blue:#1a73e8}
@media screen{[data-office-change-ink]{fill:var(--office-change-paint)!important}}
.office-change-links{position:fixed;inset:0;z-index:3;pointer-events:none;overflow:hidden}
.office-change-links i{position:absolute;box-sizing:border-box;border-bottom:1px solid var(--office-writer-selection-blue);pointer-events:none}
.office-writer-change-line{stroke:var(--office-writer-selection-blue);stroke-width:1px;vector-effect:non-scaling-stroke;pointer-events:none}
@media print{.office-change-links,.office-writer-change-line{display:none}}
${PROMPT_CSS}
.office-candidate{cursor:pointer;background:#fff;color:inherit;white-space:pre-wrap;overflow-wrap:anywhere;border:0;border-bottom:1px solid #afbac8;padding:0;text-align:start;box-shadow:0 1px 4px #0001}
.office-candidate b{display:inline-grid;place-items:center;position:relative;vertical-align:super;background:#202124;color:white;border-radius:6px;min-width:18px;height:18px;font:11px/18px system-ui;margin-inline-start:3px}
.office-review{display:flex;gap:0;background:white;box-shadow:0 2px 8px #0002;border-radius:6px;transform:translateY(-100%)}
.office-review button{box-shadow:none}
.office-region-error:not(:empty){width:fit-content;margin:3px auto;padding:2px 8px;border-radius:6px;background:white;color:#9f302b;font-size:12px}
`;

export function createSelectionUi(document, { locale = 'zh' } = {}) {
  const en = locale === 'en';
  const texts = SELECTION_TEXTS[locale] ?? SELECTION_TEXTS.zh;
  const handlers = {};
  const node = (tag, css, parent = document.body) => {
    const el = document.createElement(tag);
    el.className = css;
    parent.append(el);
    return el;
  };
  const button = (css, label, parent) => {
    const el = node('button', css, parent);
    el.type = 'button';
    el.textContent = label;
    return el;
  };
  const style = node('style', '', document.head);
  style.textContent = CSS;
  const chip = button('office-region office-chip', texts.chip);
  chip.setAttribute(MARK,'chip');
  chip.hidden = true;
  chip.insertAdjacentHTML('afterbegin','<svg class="office-selection-sparkles" viewBox="0 0 24 26" fill="currentColor" aria-hidden="true"><path d="M11.5 1c.5 2.1 1.2 2.9 3.3 3.4-2.1.5-2.8 1.3-3.3 3.4-.5-2.1-1.3-2.9-3.4-3.4 2.1-.5 2.9-1.3 3.4-3.4Z"/><path d="M5.7 5.8c.7 3.5 1.9 4.9 5.5 5.6-3.6.7-4.8 2.1-5.5 5.7-.7-3.6-2-5-5.5-5.7 3.5-.7 4.8-2.1 5.5-5.6Z"/><path d="M15.1 8.6c1 5.4 3 7.5 8.4 8.4-5.4 1-7.4 3-8.4 8.5-1-5.5-3-7.5-8.5-8.5 5.5-.9 7.5-3 8.5-8.4Z"/></svg>');
  const panel = node('section','office-region office-panel');
  panel.setAttribute('role','dialog');
  panel.setAttribute(MARK,'panel');
  panel.hidden = true;
  const foot = node('footer','office-prompt',panel);
  const input = node('textarea','',foot);
  input.setAttribute(MARK,'input');
  input.setAttribute('aria-label',texts.chip);
  input.rows = 1;
  input.wrap = 'off';
  input.maxLength = 500;
  input.placeholder = texts.placeholder;
  const send = button('office-send office-prompt-send',texts.send,foot);
  send.setAttribute(MARK,'send');
  send.setAttribute('aria-label',texts.send);
  send.innerHTML = SEND_ICON;
  const hint = node('div','office-region-error',panel);
  hint.setAttribute('role','status');
  const review = node('div','office-region office-review');
  review.hidden = true;
  let selectedId;
  for (const [action, label] of [['restore',en?'Restore':'恢复'],['apply',en?'Apply':'应用'],['always',en?'Always apply':'始终应用']]) {
    button('',label,review).onclick = () => handlers.decide(selectedId,action);
  }
  const candidateNodes = new Map();
  input.oninput = () => { send.disabled = !input.value.trim(); handlers.draft(input.value); };
  input.onkeydown = event => {
    if (isSend(event)) { event.preventDefault(); handlers.submit(input.value); }
    else if (event.key === 'Escape') handlers.close();
  };
  send.onclick = () => handlers.submit(input.value);
  chip.onmousedown = event => event.preventDefault();
  chip.onclick = () => handlers.open();

  const removeNodes = [style,chip,panel,review];
  function updateReview(el, id) {
    selectedId = id;
    const rect = el.getBoundingClientRect();
    review.hidden = false;
    review.style.left = Math.max(3,Math.min(rect.left,document.defaultView.innerWidth-review.offsetWidth-3))+'px';
    review.style.top = Math.max(review.offsetHeight+3,rect.top-8)+'px';
  }
  let anchor, placeAnchor;
  const place = (element, box, gap = 8) => {
    if (!box) return;
    const w = document.defaultView;
    const left = box.left + (box.width - element.offsetWidth) / 2;
    const above = box.top - gap - element.offsetHeight;
    element.style.left = Math.max(4,Math.min(left,w.innerWidth-element.offsetWidth-4))+'px';
    element.style.top = Math.max(4,Math.min(above >= 4 ? above : box.top+box.height+gap,w.innerHeight-element.offsetHeight-4))+'px';
  };
  return {
    handlers, texts, owns: target => !!target?.closest?.('.office-region'),
    showChip(point,{ disabled, reason, box } = {}) {
      chip.hidden = false;
      chip.disabled = !!disabled;
      chip.title = texts.reasons[reason] || texts.chipTitle;
      place(chip,box ?? {left:point.x,width:0,top:point.y,height:0},4);
    },
    hideChip() { chip.hidden = true; },
    muteNativeSelection(on) { document.documentElement.toggleAttribute('data-office-submitted-selection',on); },
    openPanel(ref, locate) { anchor = ref; placeAnchor = locate; panel.hidden = false; place(panel,locate?.(ref)); input.focus(); },
    closePanel() { panel.hidden = true; anchor = undefined; },
    get panelOpen() { return !panel.hidden; },
    busy(busy) { input.disabled = busy; send.disabled = busy || !input.value.trim(); },
    hint: message => { hint.textContent = message ?? ''; },
    render(state, locate) {
      if (input.value !== state.draft) input.value = state.draft;
      send.disabled = input.disabled || !input.value.trim();
      this.position(state,locate);
    },
    position(state,locate) {
      if (anchor && !panel.hidden) place(panel,(placeAnchor ?? locate)(anchor));
      const alive = new Set();
      for (const ref of state.candidates) {
        const key = ref.id ?? ref.number;
        alive.add(key);
        let el = candidateNodes.get(key);
        if (!el) {
          el = button('office-region office-candidate','');
          candidateNodes.set(key,el);
          node('span','',el).textContent = ref.preview;
          node('b','',el).textContent = ref.number;
          el.onclick = () => updateReview(el,ref.id);
        }
        const box = locate(ref);
        const win = document.defaultView;
        el.hidden = !box || box.top+box.height<0 || box.top>win.innerHeight || box.left+box.width<0 || box.left>win.innerWidth;
        if (el.hidden) continue;
        el.style.left = box.left+'px'; el.style.top = box.top+'px';
        el.style.maxWidth = Math.max(45,win.innerWidth-box.left-16)+'px';
        if(ref.typography) Object.assign(el.style,ref.typography);
        if (selectedId === ref.id) updateReview(el,ref.id);
      }
      for (const [key,el] of candidateNodes) if (!alive.has(key)) { el.remove(); candidateNodes.delete(key); }
      if (!state.candidates.some(ref => ref.id===selectedId && !candidateNodes.get(ref.id ?? ref.number)?.hidden)) review.hidden = true;
    },
    dispose() { for (const el of [...removeNodes,...candidateNodes.values()]) el.remove(); },
  };
}

export function createSelectionFeature({ ui, agent, notify, window, document, reviews, editable = () => true, allowSelection = true, debounceMs = 120 }) {
  let state = { draft:'', references:[], pending:[], candidates:[], changes:reviews?.current ?? [], next:1 };
  let mode = 'always', current, timer, tick, epoch = 0, disposed = false, deciding = false;
  const locate = ref => agent.locate?.(ref) ?? ref.box;
  // Local instructions live only as long as the open selection prompt.
  let promptDraft = '';
  const paint = () => ui.render({...state,draft:promptDraft},locate);
  const links=agent.changeLocations && createChangeLinks(document,window,records=>agent.changeLocations(records),
    id=>notify('document-ui',{revealChange:id}),changeAnchors=>notify('document-ui',{changeAnchors}));
  const save = () => {
    reviews?.stage(state.changes);
    notify('document-ui',{ draftState:structuredClone(state) });
    void links?.sync(state.changes);
    paint();
  };
  let activeRef, dragPointer, muted, running;
  let chatEpoch=0, chatCycle=0, nextPreviewNumber=0, chatTurn=false, chatPreview, lastChatSelection;
  const declaredPreviews = new Set();
  const working = new Map(), submissions = new Map();
  const sameRange = (a,b) => !!a && a.kind === b?.kind && sameChangeAnchor(a.anchor,b.anchor) &&
    (a.kind === 'xlsx' || a.original === b.original);
  const mute = (ref,hideNative=!!ref) => { muted=ref; ui.muteNativeSelection?.(hideNative); };
  function beginWork(ref) {
    if (mode!=='always' || !ref.anchor || !agent.beginSelectionWork) return;
    const job = { ref, handle:agent.beginSelectionWork(ref) };
    working.set(ref.number,job);
    return job;
  }
  function endWork(job) {
    if (!job || working.get(job.ref.number)!==job) return;
    working.delete(job.ref.number);
    declaredPreviews.delete(job);
    if (chatPreview === job) chatPreview = undefined;
    job.handle.dispose();
  }
  async function updateWaitingRanges() {
    const list=[...working.values()].filter(job=>job.ref.kind==='docx' && !job.writing && job.handle.point);
    if (!list.length || !agent.changeLocations) return;
    const records=list.map(({ref})=>({id:ref.number,kind:ref.kind,revert:{anchor:ref.anchor,original:ref.original},located:ref}));
    try {
      const locators=await agent.changeLocations(records,true);
      for(const job of list) {
        if(!job.writing && working.get(job.ref.number)===job) job.handle.point(locators.get(job.ref.number) ?? (()=>undefined));
      }
    } catch { /* Retry on the next committed change. */ }
  }
  function clearChatPreview() {
    ++chatEpoch;
    if(chatPreview && !chatPreview.writing)endWork(chatPreview);
    chatPreview=undefined;
  }
  async function previewChat({instruction,selection}) {
    if (declaredPreviews.size) return;
    clearChatPreview();
    const token=chatEpoch;
    let ref;
    try {ref=await agent.chatSelection?.(instruction,selection ?? current);} catch {return;}
    if(token!==chatEpoch || disposed || !running || !ref || ref.unsupported || !editable() || agent.applying || [...working.values()].some(job=>sameRange(job.ref,ref)))return;
    chatPreview=beginWork({...ref,number:--nextPreviewNumber});
    if(chatPreview){++epoch;hide();mute(ref);}
  }
  function resumePending() {
    for (const ref of state.pending) {
      if ((!running && !ref.independent) || !ref.requestId || working.has(ref.number)) continue;
      const job=beginWork(ref);
      if (job) { job.id=ref.id; job.independent=!!ref.independent; mute(ref); }
    }
  }
  const close = () => { activeRef = undefined; promptDraft = ''; ui.closePanel(); ui.hint(); void agent.focusEditor?.(true); };
  const hide = () => { current = undefined; ui.hideChip(); };
  const showChip = () => {
    if (sameRange(current,muted) || [...working.values()].some(job => sameRange(job.ref,current))) { ui.hideChip(); return; }
    const box = current && (agent.locateActiveSelection?.(current) ?? locate(current));
    if (!box || box.top+box.height < 0 || box.top > window.innerHeight) { ui.hideChip(); return; }
    ui.showChip({ x:box.left+box.width/2, y:box.top },{ box, disabled:!!current.unsupported, reason:current.unsupported });
  };
  const refresh = async () => {
    const token = ++epoch;
    if (dragPointer !== undefined || (ui.panelOpen && ui.owns(document.activeElement))) return;
    if (!allowSelection || !editable() || agent.applying) return hide();
    let selection;
    try { selection = await agent.readSelection(); } catch {}
    if (token !== epoch || disposed) return;
    current = selection;
    const ref=selection?.anchor && !selection.unsupported ? {kind:selection.kind,original:selection.original,anchor:selection.anchor,...(selection.cells && {cells:selection.cells})} : null;
    if (JSON.stringify(ref)!==JSON.stringify(lastChatSelection)) {
      lastChatSelection=ref;notify('document-ui',{chatSelection:ref});
    }
    showChip();
  };
  const later = () => {
    window.clearTimeout(timer);
    if (dragPointer === undefined) timer = window.setTimeout(refresh,debounceMs);
  };
  function open() {
    if (!current || current.unsupported || sameRange(current,muted) || [...working.values()].some(job => sameRange(job.ref,current))) return;
    promptDraft = '';
    activeRef = { ...current, number:state.next++ };
    hide();
    void agent.focusEditor?.(false);
    ui.hint();
    ui.busy(false);
    ui.openPanel(activeRef,agent.locateActiveSelection ?? locate);
    save();
  }
  function submit(input) {
    const instruction = String(input).trim();
    if (!instruction || !activeRef || [...submissions.values()].some(s=>s.ref===activeRef)) return;
    dispatch(instruction,activeRef);
  }
  function dispatch(instruction,ref,onResult) {
    const requestId = 'q'+Date.now().toString(36)+'-'+ref.number;
    const independent = !!(running || submissions.size || state.pending.length || working.size);
    const job = beginWork(ref);
    if (job) job.independent=independent;
    submissions.set(requestId,{ref,job,onResult});
    const {number,kind,anchor,original,cells} = ref;
    if (!onResult) { ui.busy(true); ui.hint(ui.texts.sending); }
    if (job) {
      ++epoch; hide(); mute(ref); ui.closePanel();
    }
    notify('selection-request',{kind:'region-batch',requestId,instruction,...(independent && {parallel:true}),...(onResult && {fromChat:true}),references:[{number,kind,anchor,original,cells}]});
  }
  async function decide(id,action) {
    const candidate = state.candidates.find(item=>item.id===id);
    if (!candidate || deciding) return;
    deciding = true;
    try {
      if (action !== 'restore') {
        await agent.rewriteSelection(candidate.args);
        if (action==='always') { mode='always'; notify('document-ui',{permission:'always'}); }
      }
      notify('selection-update',{selection:id,state:action==='restore'?'reverted':'kept'});
      state.candidates = state.candidates.filter(item=>item!==candidate);
      save();
    } catch (error) {
      void agent.focusEditor?.(false);
      ui.openPanel(activeRef ?? candidate,locate);
      ui.hint(error.message || ui.texts.revertFailed);
    }
    finally { deciding=false; }
  }
  Object.assign(ui.handlers,{open,close,submit,decide,
    draft(value) { if (activeRef) promptDraft=value; },
  });
  function beginPointer(event) {
    if (event.isPrimary === false || ui.owns(event.target)) return;
    if (event.target?.closest?.('#editor-root')) mute(undefined);
    dragPointer = event.pointerId;
    ++epoch;
    window.clearTimeout(timer);
    hide();
    if (ui.panelOpen || activeRef) close();
  }
  function endPointer(event) {
    if (dragPointer === undefined || (event.pointerId !== undefined && event.pointerId !== dragPointer)) return;
    dragPointer = undefined;
    later();
  }
  // Recover when a mouse was released outside the editor iframe.
  function pointerMove(event) { if (event.buttons === 0) endPointer(event); }
  function key(event) {
    if (ui.owns(event.target)) return;
    if (event.type==='keydown' && event.target?.closest?.('.writer-next-workbench__viewport') && (event.shiftKey && /^(Arrow|Home|End)/.test(event.key) ||
        (event.metaKey || event.ctrlKey) && event.key.toLowerCase()==='a')) mute(undefined);
    if (event.type==='keydown' && current && isShortcut(event)) { event.preventDefault(); event.stopImmediatePropagation(); open(); }
    else if (event.type==='keyup') later();
  }
  const position = () => {
    if (tick !== undefined) return;
    tick=window.requestAnimationFrame(()=>{ tick=undefined; showChip(); ui.position(state,locate); });
  };
  const events=[['pointerdown',beginPointer],['pointermove',pointerMove],['pointerup',endPointer],['pointercancel',endPointer],['keydown',key],['keyup',key],['scroll',position],['selectionchange',later]];
  for(const [type,fn] of events) document.addEventListener(type,fn,true);
  window.addEventListener('blur',endPointer);
  window.addEventListener('resize',position);
  if (reviews) {
    notify('document-ui',{restoredChanges:state.changes});
    void links?.sync(state.changes);
  }
  paint();
  return {
    refresh, layoutChanged:()=>links?.schedule(), contentChanged:()=>Promise.all([links?.sync(state.changes,true),updateWaitingRanges()]), composer(data) {
      if (data?.reason==='no-composer') return {ok:true};
      const wasRunning = running;
      running = data?.reason==='running';
      if(running && data.chatEdit){chatTurn=true;void previewChat(data.chatEdit);}
      resumePending();
      if (!running && (wasRunning || chatTurn)) {
        ++chatCycle;chatTurn=false;clearChatPreview();
        for (const job of working.values()) {
          if (job.writing || job.independent) continue;
          job.ended = true;
          endWork(job);
          if (job.id) {
            state.pending=state.pending.filter(r=>r.id!==job.id);
            notify('selection-update',{selection:job.id,state:'cancelled'});
          }
        }
        save();
      }
      return {ok:true};
    },
    async inspectEditTarget(target, signal) {
      const cycle = chatCycle;
      const doc = await agent.inspect();
      const ref = agent.modelEditRange(doc, target);
      if (disposed || signal?.aborted || cycle !== chatCycle) throw new Error('The editing turn ended.');
      if (!editable() || mode !== 'always') return doc;
      const existing = [...working.values()].find(job => sameRange(job.ref, ref));
      if (existing && existing !== chatPreview && !declaredPreviews.has(existing)) throw new Error(ui.texts.reasons.busy);
      // A tool names the actual passage. A slower guess from the submitted message must never replace it.
      ++chatEpoch;
      chatTurn = true;
      if (chatPreview && existing !== chatPreview) clearChatPreview();
      const job = existing ?? beginWork({ ...ref, number: --nextPreviewNumber });
      if (job) {
        declaredPreviews.add(job);
        chatPreview = undefined;
        ++epoch;hide();mute(ref);
      }
      return doc;
    },
    async submitChatRegion({instruction,selection}) {
      if (!editable() || !instruction?.trim()) return {handled:false};
      const ref = await agent.chatSelection?.(instruction,selection);
      if (!ref || ref.unsupported || (ref.original?.length ?? 0)>2000 || (ref.kind==='docx' && ref.anchor.paragraphs[1]-ref.anchor.paragraphs[0]>=20)) return {handled:false};
      if ([...working.values()].some(job=>sameRange(job.ref,ref))) throw new Error(ui.texts.reasons.busy);
      const reply=await new Promise(resolve=>dispatch(instruction,{...ref,number:state.next++},resolve));
      if (!reply.ok) throw new Error(ui.texts.reasons[reply.reason] || ui.texts.reasons.failed);
      return {handled:true};
    },
    takeChatPreview() {
      if(!chatTurn)return;
      ++chatEpoch;
      return {
        claim(ref) {
          const job = [...declaredPreviews, chatPreview].find(item => item && !item.writing && sameRange(item.ref, ref));
          if (chatPreview && job !== chatPreview) clearChatPreview();
          if (!job) return;
          job.writing = true;
          return { ...job.handle, dispose: () => endWork(job) };
        },
        get changes() { return state.changes; },
        record(change,previousId) { state.changes=recordChange(state.changes,change,previousId);save(); },
      };
    },
    hydrate(data) {
      if (data?.state) {
        state={...structuredClone(data.state),draft:'',references:[]};
        const count=state.changes?.length ?? 0;
        state.changes=reviews?.current ?? restoreChangeHistory((state.changes ?? []).filter(card=>!card.accepted));
        if (state.changes.length !== count) save();
      }
      if (data?.permission) mode=data.permission;
      if (data?.state) { resumePending(); void links?.sync(state.changes); }
      paint();return {ok:true};
    },
    status(reply) {
      if (reply.dropped) {
        const job=submissions.get(reply.requestId)?.job ?? [...working.values()].find(job=>job.id===reply.dropped);
        if (job && !job.writing) { job.ended=true; endWork(job); }
        state.pending=state.pending.filter(ref=>ref.id!==reply.dropped);
        ui.hint(ui.texts.reasons.failed);save();return {ok:true};
      }
      const sent=submissions.get(reply.requestId);
      if (!sent) return {ok:false};
      submissions.delete(reply.requestId);
      const {ref,job,onResult}=sent;
      if (!onResult && (!activeRef || activeRef===ref)) ui.busy(false);
      if (!reply.ok) {
        endWork(job);
        if (muted===ref) mute(undefined);
        if (activeRef === ref) {
          if (job) ui.openPanel(ref,agent.locateActiveSelection ?? locate);
          ui.hint(ui.texts.reasons[reply.reason] || ui.texts.reasons.failed);
        }
      } else {
        const id = reply.ids?.find(item => item.number === ref.number)?.id;
        if (job) { job.id=id; job.independent=!!reply.independent; }
        if (id && !job?.ended && !job?.writing && !state.candidates.some(c => c.id === id)) state.pending.push({...ref,id,requestId:reply.requestId,...(reply.independent && {independent:true})});
        else if (id && job?.ended && !job.writing) notify('selection-update',{selection:id,state:'cancelled'});
        if (activeRef === ref) close();
        save();
      }
      onResult?.(reply);return {ok:true};
    },
    async changeDetail({ id,action,instruction },signal) {
      const index=state.changes.findIndex(item=>item.id===id);
      if (index<0 || deciding || agent.applying) throw new Error(ui.texts.reasons.busy);
      const record=state.changes[index];
      if (action==='select') {
        const source=links?.locationFor(id) ?? id;
        const ref=await agent.selectChange(source===id?record:state.changes.find(r=>r.id===source) ?? record);
        if (activeRef) close();
        ++epoch; hide(); mute(ref,false);
        return {selected:true};
      }
      if (action==='accept') {
        state.changes=state.changes.filter(r=>r.id!==id);
        save();return {accepted:true};
      }
      if (action==='followup') {
        const text=String(instruction ?? '').trim();
        if (!text || text.length>500) throw new Error(ui.texts.reasons.failed);
        if (!editable()) throw new Error(ui.texts.reasons['read-only']);
        deciding=true;
        try {
          // Resolve again on send: typing in the card can outlast a selection change or a document edit.
          const ref=await agent.selectChange(record);
          if (ref?.unsupported) throw new Error(ui.texts.reasons[ref.unsupported]);
          if (!ref || (ref.kind==='docx' && !ref.original?.trim())) throw new Error(ui.texts.reasons['no-target']);
          if (disposed || signal?.aborted) throw new Error(ui.texts.reasons.busy);
          const reply=await new Promise(resolve=>dispatch(text,{...ref,number:state.next++},resolve));
          if (!reply.ok) throw new Error(ui.texts.reasons[reply.reason] || ui.texts.reasons.failed);
          return {sent:true};
        } finally { deciding=false; }
      }
      if (action!=='remove') throw new Error('Unknown change action');
      deciding=true;
      try {
        await agent.revertSelection(record.revert,signal,record.located);
        state.changes=state.changes.filter(r=>r.id!==id);
        notify('selection-update',{selection:record.latestId ?? id,state:'reverted'});
        save();return {removed:true};
      } finally { deciding=false; }
    },
    async rewrite(args,signal) {
      let job=[...working.values()].find(item=>item.id===args.id) ??
        [...working.values()].find(item=>!item.id && sameRange(item.ref,args));
      if (!job?.independent && !state.pending.some(r=>r.id===args.id)) {
        ++chatEpoch;
        if(chatPreview && !sameRange(chatPreview.ref,args))clearChatPreview();
      }
      const ref=state.pending.find(item=>item.id===args.id || (job && item.number===job.ref.number)) ?? job?.ref;
      if (mode==='always') {
        if (!job && chatTurn && agent.describeSelection && agent.beginSelectionWork) {
          const actual=await agent.describeSelection(args);
          job=beginWork({...actual,number:state.next++});
        }
        if (job?.writing) throw new Error(ui.texts.reasons.busy);
        if (job) job.writing=true;
        let outcome;
        try { outcome=await agent.rewriteSelection(args,signal,job?.handle,ref,state.changes); }
        finally {
          endWork(job);
          if(chatPreview===job)chatPreview=undefined;
          state.pending=state.pending.filter(item=>item.id!==args.id && (!job || item.number!==job.ref.number));
          if (!outcome) save();
        }
        if (outcome.reply.applied && outcome.revert) {
          const display = value => Array.isArray(value) ? value.map(row=>Array.isArray(row)?row.join(' · '):row).join('\n') : String(value ?? '');
          state.changes=recordChange(state.changes,{id:args.id,kind:args.kind,appliedAt:Date.now(),before:display(outcome.reply.before),after:display(outcome.reply.after),
            revert:outcome.revert,located:outcome.located},outcome.previousId);
        }
        save();return outcome.reply;
      }
      // Even if the tab was remounted or the request came from another composer, resolve its actual target.
      const located=ref ?? await agent.describeSelection?.(args) ?? args;
      const candidate={...located,args,id:args.id,number:ref?.number ?? state.next++,preview:args.text ?? args.values?.map(row=>row.join(' · ')).join('\n') ?? ''};
      const box=locate(candidate);
      if (box && args.kind==='docx') {
        const glyph=document.elementsFromPoint?.(box.left+1,box.top+box.height/2)?.find(el=>el.closest?.('text[data-writer-document-content]'));
        if(glyph) { const css=window.getComputedStyle(glyph);candidate.typography={fontFamily:css.fontFamily,fontSize:Math.max(9,Math.min(70,box.height/Math.max(1,args.anchor.paragraphs[1]-args.anchor.paragraphs[0]+1)*0.85))+'px',fontWeight:css.fontWeight}; }
      }
      state.pending=state.pending.filter(item=>item.id!==args.id);
      state.candidates.push(candidate);
      save();
      return {applied:false,pendingReview:true,after:candidate.preview};
    },
    dispose() { disposed=true; links?.dispose(); for(const sent of submissions.values())sent.onResult?.({ok:false,reason:'failed'}); submissions.clear(); ++epoch;for(const job of working.values()) endWork(job);mute(undefined);window.clearTimeout(timer);if(tick!==undefined)window.cancelAnimationFrame(tick);for(const [type,fn] of events)document.removeEventListener(type,fn,true);window.removeEventListener('resize',position);window.removeEventListener('blur',endPointer);ui.dispose(); },
  };
}
