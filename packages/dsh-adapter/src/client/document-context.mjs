import { PARALLEL_EDIT_COMMAND, parallelInputDefinition, parallelMessage } from '../parallel-message.mjs';
import {isChatEdit,usesChatSelection} from '../../../core/src/agent/chat-edit-target.mjs';
import { registerEditQuote } from './edit-quote.mjs';
import { baseName } from './paths.mjs';

const SLOT = 'conversation.input.dock';
const NS = 'officesdkDocumentContext';
const SEND_LABELS = new Set(['发送消息', '排队发送', '插话发送', 'Send message', 'Queue message', 'Steer']);

export function documentTarget(sidebar, controllers, sessionId) {
  if (!sessionId || sidebar?.mounted?.getSnapshot() !== sessionId || !sidebar.isExpanded?.()) return undefined;
  const tab = sidebar.active?.();
  const controller = controllers.get(tab?.id);
  return controller?.source?.sessionId === sessionId ? controller : undefined;
}

export function documentQuote(selection) {
  if (selection?.unsupported) return '';
  if (typeof selection?.original === 'string') return selection.original;
  if (Array.isArray(selection?.original)) return selection.original.join('\n');
  return '';
}

export function documentQuoteSummary(text, title, label = text) {
  return { label: label.replace(/\s+/g, ' ').trim().slice(0, 120), title, body: text };
}

// Both composer surfaces use the same session. Stage the target here and start the visual
// only when DSH reports that this submission has actually entered a running turn.
export function stageChatEdit(controller,instruction) {
  const text=instruction?.trim();
  controller.pendingChatEdit=isChatEdit(text)?{instruction:text,selection:usesChatSelection(text)?controller.chatSelection:undefined}:undefined;
  return controller.pendingChatEdit;
}

const chatDispatches=new WeakMap();
export function tryParallelChat(controller,registry,id,lang='zh') {
  const entry=registry.get(id), instruction=entry?.draft;
  if (!entry || (!entry.running && !/\.(?:xlsx|pptx)$/i.test(controller.source?.path ?? '')) || entry.attachments || !isChatEdit(instruction)) return;
  const pending=chatDispatches.get(controller);
  if (pending?.instruction===instruction) return pending.promise;
  const job={instruction};
  const selection=usesChatSelection(instruction)?controller.chatSelection:undefined;
  job.promise=controller.agent('submitChatRegion',{instruction,selection}).then(result=>{
    const current=registry.get(id);
    if (current?.draft!==instruction) return;
    if (result?.handled) current.actions.setDraft('');
    else {
      if (!current.running) {
        const intent=stageChatEdit(controller,instruction), path=controller.source?.path;
        if (path) {
          const body=(lang==='en'?'Editing target: ':'编辑目标：')+path+(intent?.selection?.original ? '\n'+intent.selection.original : '');
          current.actions.insertQuote?.(body,documentQuoteSummary(body,path,intent?.selection?.original || path));
        }
      }
      return current.actions.submit();
    }
  }).finally(()=>{if(chatDispatches.get(controller)===job)chatDispatches.delete(controller);});
  chatDispatches.set(controller,job);
  return job.promise;
}

export function registerDocumentContext(ctx, React, { sidebarOf, controllers, composers, document }) {
  const { createElement: h, useEffect, useRef, useState } = React;
  const readTarget = sessionId => documentTarget(sidebarOf(), controllers, sessionId);

  function DocumentContext({ session, sessionId, input, t }) {
    const id = sessionId ?? session?.sessionId ?? session?.id;
    const root = useRef(null);
    const [target, setTarget] = useState();
    const [error,setError]=useState('');
    useEffect(() => {
      const line = root.current;
      const seat = line?.closest('[data-composer-seat]');
      if (!seat) return;
      let tick;
      const align = () => {
        const editable = seat.querySelector('[data-lexical-editor="true"],[contenteditable="true"],textarea');
        if (!editable) return;
        const css = document.defaultView.getComputedStyle(editable);
        const inset = editable.getBoundingClientRect().left - line.getBoundingClientRect().left + parseFloat(css.paddingLeft || 0) + parseFloat(css.borderLeftWidth || 0);
        if (Number.isFinite(inset)) line.style.paddingLeft = Math.max(0, inset) + 'px';
      };
      const watch = new document.defaultView.ResizeObserver(() => {
        cancelAnimationFrame(tick);
        tick = requestAnimationFrame(align);
      });
      watch.observe(seat);
      align();
      return () => { watch.disconnect(); cancelAnimationFrame(tick); };
    }, [id, target?.controller, input?.phase]);

    useEffect(() => {
      if (!id) return undefined;
      let frame;
      let alive = true;
      let current;
      let unsubscribe;
      const win = document.defaultView;
      const update = () => {
        frame = undefined;
        if (!alive) return;
        const controller = readTarget(id);
        if (current !== controller) {
          unsubscribe?.();
          current = controller;
          unsubscribe = controller?.subscribe(schedule);
        }
        const snapshot = controller ? { controller, mode: controller.state.mode, error:controller.state.parallelError } : undefined;
        setTarget(previous => previous?.controller === snapshot?.controller && previous?.mode === snapshot?.mode && previous?.error === snapshot?.error ? previous : snapshot);
      };
      function schedule() {
        if (frame === undefined) frame = win.requestAnimationFrame(update);
      }
      const observer = new win.MutationObserver(schedule);
      observer.observe(document.body, { childList: true, subtree: true, attributes: true, attributeFilter: ['aria-selected', 'data-state', 'hidden'] });
      const offMounted = sidebarOf()?.mounted?.subscribe?.(schedule);
      update();
      return () => {
        alive = false;
        observer.disconnect();
        if (frame !== undefined) win.cancelAnimationFrame(frame);
        unsubscribe?.();
        offMounted?.();
      };
    }, [id]);

    useEffect(() => {
      const seat = root.current?.closest('[data-composer-seat]');
      if (!seat || seat.closest?.('.office-chat-decisions') || target?.mode !== 'edit' || input?.phase !== 'plain') return undefined;
      const rememberTarget = event => {
        const takeover=event.target.closest?.('[data-slot="conversation.composer"]');
        if (takeover && !event.target.closest('[data-chain-overlay-fallback="conversation.composer"]')) return;
        const trigger = event.type === 'click'
          ? event.target.closest?.('[data-composer-card] button')
          : undefined;
        if (event.type === 'click' && (!SEND_LABELS.has(trigger?.getAttribute('aria-label')) || trigger.disabled)) return;
        if (event.type === 'keydown' && (event.key !== 'Enter' || event.repeat || event.isComposing || event.keyCode === 229 || event.shiftKey || !event.target.closest?.('[data-lexical-editor],textarea'))) return;
        const controller = readTarget(id);
        const entry = composers.get(id);
        const draft = entry?.draft ?? '';
        if (controller !== target.controller || (!draft.trim() && !entry?.attachments) || draft.trimStart().startsWith('/')) return;
        const parallel=tryParallelChat(controller,composers,id,t?.('lang'));
        if (parallel) {
          event.preventDefault(); event.stopImmediatePropagation();
          setError('');void parallel.catch(err=>setError(err.message));return;
        }
        if (!entry?.running) stageChatEdit(controller,draft);
        const path = controller.source.path;
        const quote = t?.('lang') === 'en' ? `Editing target: ${path}` : `编辑目标：${path}`;
        if (!draft.includes(quote)) {
          const selected=controller.pendingChatEdit?.selection?.original;
          const body=selected?quote+'\n'+selected:quote;
          entry?.actions?.insertQuote?.(body,documentQuoteSummary(body,path,selected || path));
        }
      };
      seat.addEventListener('click', rememberTarget, true);
      seat.addEventListener('keydown', rememberTarget, true);
      return () => {
        seat.removeEventListener('click', rememberTarget, true);
        seat.removeEventListener('keydown', rememberTarget, true);
      };
    }, [id, target, input?.phase, t]);

    if (!target || input?.command || input?.phase === 'claimed') return null;
    const { controller, mode } = target;
    const name = baseName(controller.source.path, controller.platform);
    const focus = () => {
      if (readTarget(id) !== controller) return;
      sidebarOf()?.focus?.(sidebarOf()?.active?.()?.id);
      controller.frame?.focus?.();
    };
    const viewing = mode !== 'edit';
    return h(React.Fragment,null,
      (error || target.error) && h('div',{role:'alert'},error || target.error),
      h('div', { ref: root, className: 'sbp-document-context' },
      h('div', { className: 'sbp-document-target', 'aria-live': 'polite' }, t?.(viewing ? 'viewing' : 'editing'), ' ',
        h('button', { type: 'button', className: 'sbp-document-name', title: name, onClick: focus }, name.replace(/\.[^.]+$/, '') || name),
        ' ', t?.('region'))));
  }

  function ParallelMessage({ node, renderFactorySlot }) {
    const message = parallelMessage(node);
    if (!message) return null;
    return h('div', {className:'office-sent-edit'},
      renderFactorySlot('conversation.chat.user.content', {
        content: {text:message.text, attachments:[], extraBlocks:[]}, facts:{}, interactive:false
      }),
      message.error && h('small', {role:'alert'}, message.error));
  }
  const messageSlot = 'conversation.chat.commandview';
  let projected = false;
  ctx.inject?.(['uiConversation'], scope => {
    registerEditQuote(scope, React);
    scope.uiConversation.events.register(parallelInputDefinition);
    const slot = 'conversation.chat.node';
    scope.slots.inject(slot, () => scope.slots.register({ name: slot, key: parallelInputDefinition.kind },
      props => h(ParallelMessage, { ...props, node: props.node.data })));
    scope.effect(() => {
      projected = true;
      return () => { projected = false; };
    });
  });
  ctx.effect(() => ctx.slots.inject(messageSlot, () => ctx.slots.register({ name: messageSlot, key: PARALLEL_EDIT_COMMAND },
    props => projected ? null : h(ParallelMessage, props))));
  ctx.effect(() => ctx.locale.register(NS, {
    zh: { lang: 'zh', editing: '正在编辑', viewing: '正在查看', region: '文档', quote: '引用当前文档的选区', selectFirst: '请先在右侧选中可引用的内容' },
    en: { lang: 'en', editing: 'Editing', viewing: 'Viewing', region: 'document', quote: 'Quote the current selection', selectFirst: 'Select content in the document first' },
  }));
  ctx.effect(() => ctx.slots.inject(SLOT, () => ctx.slots.register({ name: SLOT, id: 'officesdk-document-context', order: 10, locale: NS }, DocumentContext)));
}
