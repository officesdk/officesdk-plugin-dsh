import { focusOfficeTurn, deckModels, createDeckCard, turnActivity } from './deck-chat.mjs';
import { parallelMessage } from '../parallel-message.mjs';
import {groupChangeCards,positionChangeCards,bindCardSwipe,slideCardStrip,stripOffset} from './change-rail-layout.mjs';
import {stageChatEdit,documentQuoteSummary,tryParallelChat} from './document-context.mjs';
import {officeIcon} from '../../../core/src/browser/office-brand.mjs';

export function focusChatRows(chat, {since,lastSeq}, en=false, running=false, snapshot) {
  const rows = [];
  const decks=new Map();
  for (const turn of snapshot?.timeline?.turnOrder ?? []) {
    const location=snapshot.timeline.turns.get(turn);
    if (location?.end?.time < since) continue;
    const office=focusOfficeTurn(location);
    if (!office) continue;
    const {suite,time,seq}=office;
    const beganHere=time>=since && seq>lastSeq;
    const models=deckModels(snapshot,turn).filter(model=>beganHere || model.lastActivity>=since ||
      (location.end?.time>=since && location.end?.seq>lastSeq));

    const activities=turnActivity(snapshot,turn,en).filter(a=>a.time>=since);
    if(location.status!=='closed' && (beganHere || activities.length)){
      const model=models[0] ?? {turn,closed:false,start:time,suite};
      decks.set(turn,{seq,models:[model],activities});
    }else if(models.length)decks.set(turn,{seq,models,activities});
  }
  const add = (role, blocks, key) => {
    const text = (blocks ?? []).filter(b => (b.kind ?? b.type) === 'text').map(b => b.text).join('\n').trim();
    if (text) rows.push({ role, text, key });
  };
  for (const node of chat?.nodes ?? []) {
    // Sequence excludes nodes already visible on entry. Event time also excludes
    // older history that DSH hydrates later, including updated parallel commands.
    if (node.seq<=lastSeq || !(node.time>=since)) continue;
    if (node.kind === 'user' || node.kind === 'steering') add('user', node.content, node.seq);
    const edit = parallelMessage(node);
    if (edit) {
      rows.push({role:'user', text:edit.text, key:node.seq});
      if (edit.error) rows.push({role:'error', text:edit.error, key:'error-'+node.seq});
      else rows.push({role:'assistant',key:'work-'+node.seq,text:node.outcome?.kind==='success'
        ? (en?'Edit complete.':'这处修改已完成。') : (en?'Editing…':'正在修改…')});
    }
    if ((node.kind === 'turn-error' || node.kind === 'turn-max-tokens') && (chat.turnTimings?.get(node.turn)?.startTime ?? node.time)>=since) {
      const reason=node.kind==='turn-max-tokens' ? (en?'The response reached its length limit.':'回复已达到长度上限。')
        : node.message || node.code || (en?'The request failed.':'请求失败，请重试。');
      rows.push({role:'error',text:reason,key:node.seq});
    }
    if (node.kind === 'assistant' && !decks.get(node.turn)?.models.every(m=>!m.closed || (m.savedPath && m.completed===m.count)) && (node.timing?.stepStartTime ?? chat.turnTimings?.get(node.turn)?.startTime ?? node.time)>=since)
      add('assistant', node.blocks, node.seq);
  }
  if (chat?.partial && !decks.has(chat.partial.turn) && chat.turnTimings?.get(chat.partial.turn)?.startTime>=since) add('assistant', chat.partial.blocks, 'live');
  if (running && chat?.turnTimings) {
    const turn=[...chat.turnTimings].reverse().find(([,timing])=>timing.startTime>=since && timing.endTime==null);
    if (turn && !decks.has(turn[0]) && !rows.some(row=>row.key==='live')) {
      const tool=chat.runningCalls?.some(call=>call.turn===turn[0]);
      rows.push({role:'status',key:'progress-'+turn[0],text:tool?(en?'Running a tool…':'正在调用工具…'):(en?'Processing…':'正在处理…')});
    }
  }
  for (const {seq,models,activities} of decks.values()) {
    const next=rows.findIndex(r=>typeof r.key==='number' && r.key>seq);
    rows.splice(next<0?rows.length:next,0,...models.map(model=>({role:'deck',key:'deck-'+model.turn+(model.closed?'-'+model.path:''),model,activities})));
  }
  return rows;
}

export function createEditSurfaces(React) {
  const { createElement: h, useEffect, useRef, useState, useSyncExternalStore } = React;
  // Toolbar, editor and rail have different React owners. Keep each decision's original
  // callbacks while presenting it once in the focused document's chat.
  const decisionStores=new WeakMap();
  function decisionsFor(controller) {
    if (decisionStores.has(controller)) return decisionStores.get(controller);
    const items=new Map(), listeners=new Set();
    let snapshot=[];
    const refresh=()=>{snapshot=[...items.values()];for(const fn of listeners)fn();};
    const store={get:()=>snapshot,subscribe:fn=>{listeners.add(fn);return ()=>listeners.delete(fn);},
      set(id,node){if(items.get(id)!==node){items.set(id,node);refresh();}},
      remove(id){if(items.delete(id))refresh();}};
    decisionStores.set(controller,store);
    return store;
  }
  function FocusDecision({controller,children}) {
    const store=decisionsFor(controller),id=React.useId();
    React.useLayoutEffect(()=>{store.set(id,children);},[store,id,children]);
    React.useLayoutEffect(()=>()=>store.remove(id),[store,id]);
    return null;
  }
  const rails = new Map();
  function railFor({sessionId,path}) {
    // Absolute paths identify the same document in different chats. A relative path belongs to its workspace session.
    const key = 'officesdk:change-rail:' + JSON.stringify(/^(?:\/|[A-Za-z]:[\\/]|\\\\)/.test(path) ? path : [sessionId,path]);
    if (rails.has(key)) return rails.get(key);
    let collapsed = false;
    try { collapsed = globalThis.localStorage?.getItem(key) === 'closed'; } catch {}
    let state = {collapsed, read:new Set()};
    const listeners = new Set();
    const update = next => { state=next; for (const fn of listeners) fn(); };
    const rail = {
      snapshot: () => state,
      subscribe(fn) { listeners.add(fn); return () => listeners.delete(fn); },
      toggle() {
        const collapsed = !state.collapsed;
        try { globalThis.localStorage?.setItem(key,collapsed?'closed':'open'); } catch {}
        update({...state,collapsed});
      },
      markRead(id) {
        if (!state.read.has(id)) update({...state,read:new Set([...state.read,id])});
      },
    };
    rails.set(key,rail);
    return rail;
  }
  function CardGroup({group,active,show,content,en}) {
    const root=useRef(null), stepRef=useRef(null), motion=useRef(null), [front,setFront]=useState();
    const cards=group.records;
    const current=cards.find(r=>r.id===active) ?? cards.find(r=>r.id===front) ?? cards[0];
    const index=cards.indexOf(current), stacked=cards.length>1;
    function step(delta) {
      if (!delta) return;
      const next=cards[(index+delta+cards.length)%cards.length].id;
      const strip=root.current.querySelector('.office-stack-strip');
      if (strip && Math.abs(delta)===1) motion.current={next,offset:stripOffset(strip)+delta*strip.getBoundingClientRect().width};
      setFront(next);show(next);
    }
    React.useLayoutEffect(()=>{
      stepRef.current=step;
      if (motion.current?.next===current.id) {
        slideCardStrip(root.current.querySelector('.office-stack-strip'),motion.current.offset);
        motion.current=null;
      }
      if (current.id===active && front!==active) setFront(active);
    });
    useEffect(()=>{
      if (stacked) return bindCardSwipe(root.current,delta=>stepRef.current(delta));
    },[stacked]);
    return h('div',{ref:root,className:'office-change-group','data-stack':stacked,
      onKeyDown:e=>{
        if (!stacked || e.altKey || e.ctrlKey || e.metaKey || e.target.closest('form,.office-undo-confirm')) return;
        if (e.key==='ArrowLeft' || e.key==='ArrowRight') {
          const delta=e.key==='ArrowRight'?1:-1;
          e.preventDefault();step(delta);
          root.current.querySelectorAll('nav button')[(index+delta+cards.length)%cards.length]?.focus({preventScroll:true});
        }
      }},
      h('div',{className:'office-stack-face'},stacked ? h('div',{className:'office-stack-strip'},
        ...[-1,0,1].map(side=>side===0 ? content(current) : h('div',{key:'neighbor'+side,className:'office-stack-neighbor',
          'data-side':side,inert:true,'aria-hidden':true},content(cards[(index+side+cards.length)%cards.length],true)))) : content(current)),
      stacked && h('nav',{'aria-label':en?'Changes on these lines':'此处的改动'},
        ...cards.map((card,i)=>h('button',{type:'button',key:card.id,tabIndex:i===index?0:-1,
          'aria-current':i===index?'true':undefined,'aria-label':en?`Change ${i+1} of ${cards.length}`:`第 ${i+1} 条，共 ${cards.length} 条`,
          onClick:()=>step(i-index)}))));
  }
  function ChangeRail({ controller, records, en, revealed, anchors }) {
    const rail=railFor(controller.source);
    const {collapsed,read}=useSyncExternalStore(rail.subscribe,rail.snapshot);
    const listId=React.useId(), list=useRef(null), lock=useRef(false);
    const selectionVersion=useRef(0);
    const [error,setError]=useState(''), [active,setActive]=useState(), [expanded,setExpanded]=useState();
    const [confirm,setConfirm]=useState(), [working,setWorking]=useState(false), [draft,setDraft]=useState();
    const [keyboard,setKeyboard]=useState(false);
    const groups=React.useMemo(()=>groupChangeCards(records ?? [],anchors?.items),[records,anchors]);
    React.useLayoutEffect(()=>{
      const el=list.current;
      if (collapsed || !el || !groups.length) return;
      const place=follow=>{
        const frame=controller.frame?.getBoundingClientRect();
        positionChangeCards(el,groups,(frame?.top ?? 0)-el.getBoundingClientRect().top,
          frame?.height && anchors?.height ? frame.height/anchors.height : 1,follow);
      };
      place(true);
      const observer=new ResizeObserver(()=>place(false));
      observer.observe(el);
      for(const node of el.firstElementChild.children)observer.observe(node);
      return ()=>observer.disconnect();
    },[collapsed,groups,anchors,controller]);
    const title=en?'Change details':'修改详情';
    const toggle=en?(collapsed?'Expand change details':'Collapse change details'):(collapsed?'展开修改详情':'收起修改详情');
    const removeLabel=en?'Undo this change':'撤销修改', acceptLabel=en?'Accept change':'采用修改';
    const label=en?'Enter editing instructions…':'输入修改要求…';
    useEffect(()=>{
      const doc=controller.frame?.ownerDocument;
      const outside=e=>{
        if(e.target.closest?.('.office-change-group,.office-rail-toggle')) return;
        selectionVersion.current++;
        setActive(undefined);
      };
      doc?.addEventListener('click',outside,true);
      return ()=>doc?.removeEventListener('click',outside,true);
    },[controller]);
    useEffect(()=>{
      if (!revealed) return;
      const {id}=revealed;
      selectionVersion.current++;
      if (id===null) {
        setActive(undefined);
        if(controller.state?.revealedChange===revealed)controller.set({revealedChange:undefined});
        return;
      }
      if (rail.snapshot().collapsed) rail.toggle();
      rail.markRead(records.find(r=>r.id===id)?.latestId ?? id);setActive(id);setExpanded(id);setConfirm(undefined);
      const frame=requestAnimationFrame(()=>{
        const el=list.current, card=Array.from(el?.querySelectorAll('article') ?? []).find(node=>node.dataset.changeId===id);
        if (card) {
          const bounds=el.getBoundingClientRect(), rect=card.getBoundingClientRect();
          if (rect.top<bounds.top || rect.height>bounds.height) el.scrollTop+=rect.top-bounds.top;
          else if (rect.bottom>bounds.bottom) el.scrollTop+=rect.bottom-bounds.bottom;
        }
        if(controller.state?.revealedChange===revealed)controller.set({revealedChange:undefined});
      });
      return ()=>cancelAnimationFrame(frame);
    },[rail,revealed]);
    async function run(id,action) {
      if (lock.current) return;
      const instruction=draft?.id===id?draft.text.trim():'';
      if (action==='followup' && !instruction) return;
      const version=selectionVersion.current;
      lock.current=true;setWorking(true);setError('');setConfirm(undefined);
      try {
        await controller.agent('changeDetail',action==='followup'?{id,action,instruction}:{id,action});
        rail.markRead(records.find(r=>r.id===id)?.latestId ?? id);
        if (action==='select') {
          if (version!==selectionVersion.current) return;
          setActive(id);setExpanded(old=>old===id?undefined:id);
          setDraft(old=>old?.id===id?old:undefined);
        } else {
          setDraft(undefined);
          if (action==='remove' || action==='accept') {setActive(undefined);setExpanded(undefined);}
        }
      } catch(e){setError(e.message);}
      finally {lock.current=false;setWorking(false);}
    }
    function icon(path,size=16) {
      return h('svg',{width:size,height:size,viewBox:'0 0 16 16',fill:'none',stroke:'currentColor',strokeWidth:1.5,'aria-hidden':true},h('path',{d:path,strokeLinecap:'round',strokeLinejoin:'round'}));
    }
    const dockDecisions=Boolean(controller.state?.focusMode && controller.state?.ready && controller.state.mode==='edit');
    function content(r,preview=false) {
      const open=expanded===r.id, asking=!preview && confirm===r.id, text=draft?.id===r.id?draft.text:'';
      const date=Number.isFinite(r.appliedAt)?new Date(r.appliedAt):null;
      const applied=en?'Accepted':'已采用';
      const undo=asking && h('div',{className:'office-undo-confirm',role:'group','aria-label':removeLabel,
          onKeyDown:e=>{if(e.key==='Escape'){setConfirm(undefined);list.current?.querySelector('article[data-change-id="'+CSS.escape(r.id)+'"] .office-change-remove')?.focus();}}},
          h('button',{type:'button',autoFocus:true,onClick:()=>setConfirm(undefined)},en?'Cancel':'取消'),
          h('button',{type:'button',onClick:()=>run(r.id,'remove')},en?'Undo':'确认撤销'));
      return h(preview?'div':'article',{className:'office-change',key:r.id,'data-change-id':preview?undefined:r.id,'data-active':active===r.id,'data-expanded':open},
        h('button',{className:'office-change-body',type:'button',disabled:working,'aria-expanded':open,
          title:open?(en?'Collapse and locate':'收起并定位正文'):(en?'Expand and locate':'展开并定位正文'),onClick:()=>run(r.id,'select')},
          h('span',{className:'office-change-meta'},h('img',{src:officeIcon,alt:'OfficeSDK',width:14,height:14}),
            date && h('time',{dateTime:date.toISOString(),title:date.toLocaleString(en?'en-US':'zh-CN')},
              date.toLocaleTimeString(en?'en-US':'zh-CN',{hour:'2-digit',minute:'2-digit',hour12:false}))),
          h('del',{className:'office-diff-text'},r.before),
          h('span',{className:'office-replacement'},icon('M2 2v3a4 4 0 0 0 4 4h7m-3-3 3 3-3 3',15),
            h('span',{className:'office-diff-text'},r.after || (en?'Deleted':'已删除')))),
        h('div',{className:'office-change-actions'},
          h('button',{type:'button',className:'office-change-remove',disabled:working,'aria-expanded':asking,title:removeLabel,'aria-label':removeLabel,onClick:()=>{setConfirm(r.id);setError('');}},icon('m4 4 8 8m0-8-8 8',14)),
          h('button',{type:'button',className:'office-change-check',disabled:working || r.accepted,'aria-pressed':!!r.accepted,
            title:r.accepted?applied:acceptLabel,'aria-label':r.accepted?applied:acceptLabel,onClick:()=>run(r.id,'accept')},icon('m3 8 3.2 3.2L13 4.5'))),
        asking ? (dockDecisions ? h(FocusDecision,{controller},undo) : undo) :
        h('form',{className:'office-card-prompt office-prompt',onSubmit:e=>{e.preventDefault();void run(r.id,'followup');}},
          h('textarea',{rows:1,wrap:'off',maxLength:500,value:text,readOnly:working,placeholder:label,'aria-label':label,
            onFocus:()=>setDraft(old=>old?.id===r.id?old:undefined),onChange:e=>setDraft({id:r.id,text:e.target.value}),
            onKeyDown:e=>{
              if(e.nativeEvent.isComposing || e.nativeEvent.keyCode===229)return;
              if(e.key==='Enter' && !e.shiftKey){e.preventDefault();void run(r.id,'followup');}
              if(e.key==='Escape' && !working){e.preventDefault();setDraft(undefined);e.currentTarget.closest('article').querySelector('.office-change-body').focus();}
            }}),
          h('button',{type:'submit',className:'office-prompt-send',disabled:working || !text.trim(),'aria-label':en?'Send':'发送修改要求'},
            h('svg',{width:18,height:18,viewBox:'0 0 20 20',fill:'none','aria-hidden':true},
              h('path',{d:'M10 16V4m-5 5 5-5 5 5',stroke:'currentColor',strokeWidth:1.65,strokeLinecap:'round',strokeLinejoin:'round'})))));
    }
    const saveFailed=controller.state?.reviewSaveFailed;
    if (!records?.length && !saveFailed) return null;
    const unread=records?.some(record=>!read.has(record.latestId ?? record.id));
    return h('aside',{className:'office-change-rail','data-collapsed':collapsed,'data-keyboard':keyboard,'aria-label':title,
      onPointerDownCapture:()=>setKeyboard(false),onKeyDownCapture:e=>{if(e.key==='Tab')setKeyboard(true);}},
      h('header',null,!collapsed && h('strong',null,title),
        h('button',{type:'button',className:'office-rail-toggle',onClick:rail.toggle,title:toggle,'aria-controls':listId,'aria-expanded':!collapsed,
          'aria-label':toggle+(unread?(en?', new changes':'，有新增修改'):'')},
          icon(collapsed?'M10 3.5 5.5 8l4.5 4.5':'M6 3.5 10.5 8 6 12.5'),
          unread && h('span',{className:'office-rail-dot','aria-hidden':true}))),
      h('div',{id:listId,ref:list,className:'office-change-list',hidden:collapsed},
        h('div',{className:'office-change-track'},...groups.map(group=>h(CardGroup,{key:group.key,group,active,content,en,
          show:id=>{setActive(id);setExpanded(undefined);setConfirm(undefined);rail.markRead(records.find(r=>r.id===id)?.latestId ?? id);}})))),
      saveFailed && h('div',{role:'alert'},en?'Change history could not be saved. ':'修改详情未能保存。',
        h('button',{type:'button',onClick:()=>controller.agent('documentUi',{retryReviews:true}).catch(e=>setError(e.message))},en?'Retry':'重试')),
      error && h('div',{role:'alert'},error));
  }
  const embeddedSlots={views:()=>null,widthControls:()=>null};
  const DeckCard=createDeckCard(React);
  function FocusChat({ registry, controller, openDeck, period, frameFullscreen, bottom, en, renderFactorySlot, pendingInteraction, proposal }) {
    const sessionId=controller.source.sessionId;
    const entry=useSyncExternalStore(fn=>registry.subscribeEntries(id=>{ if (id===sessionId) fn(); }),()=>registry.get(sessionId));
    const [open,setOpen]=useState(false), [error,setError]=useState('');
    const log=useRef(null), input=useRef(null), submitting=useRef(false);
    const rows=React.useMemo(()=>focusChatRows(entry?.chat,period,en,entry?.running,entry?.snapshot),[entry?.chat,period,en,entry?.running,entry?.snapshot]);
    const panel=useRef(null);
    React.useLayoutEffect(()=>{
      // Fullscreen puts the editor iframe above its siblings. A manual popover
      // keeps the same composer visible above it without moving focus.
      if (frameFullscreen) panel.current?.showPopover?.();
    },[frameFullscreen]);
    // Use the host's session-scoped composer election for questions, approvals and future request types.
    // Both presentations share its request identity, draft store and one-shot settlement.
    const waiting=Boolean(pendingInteraction && renderFactorySlot);
    const decisionStore=decisionsFor(controller);
    const otherDecisions=useSyncExternalStore(decisionStore.subscribe,decisionStore.get);
    const needsAnswer=waiting || Boolean(proposal) || otherDecisions.length>0;
    const expanded=open || needsAnswer;
    const ready=Boolean(entry?.actions && (entry.draft?.trim() || entry.attachments));
    useEffect(()=>{ if(expanded && log.current) log.current.scrollTop=log.current.scrollHeight; },[expanded,rows]);
    const send=async()=>{
      if(!ready) return;
      const parallel=tryParallelChat(controller,registry,sessionId,en?'en':'zh');
      if(parallel){setOpen(true);setError('');try {await parallel;} catch(e){setError(e.message);}return;}
      if(submitting.current && !entry.running) return;
      submitting.current=true;setOpen(true);setError('');
      const intent=!entry.running && stageChatEdit(controller,entry.draft);
      try {
        if(intent){
          const path=controller.source.path, quote=(en?'Editing target: ':'编辑目标：')+path;
          const selected=intent.selection?.original, body=selected?quote+'\n'+selected:quote;
          if(!entry.draft.includes(quote))entry.actions.insertQuote?.(body,documentQuoteSummary(body,path,selected || path));
        }
        await entry.actions.submit();
      } catch(e){if(controller.pendingChatEdit===intent)controller.pendingChatEdit=undefined;setError(e.message);} finally {submitting.current=false;}
    };
    const file=controller.source.path.split(/[\\/]/).at(-1);
    const recent=controller.state.changes?.findLast(change=>change.appliedAt>=period.since);
    const jump=()=>{setOpen(false);controller.frame?.focus();};
    return h('section',{ref:panel,popover:frameFullscreen?'manual':undefined,className:'office-focus-chat','data-expanded':expanded,style:{bottom:bottom ?? 46},'aria-label':en?'Conversation':'对话'},
      h('button',{type:'button',className:'office-chat-head','aria-expanded':expanded,disabled:needsAnswer,onClick:()=>setOpen(!open)},
        h('img',{src:officeIcon,alt:'',width:14,height:14}),
        h('strong',null,entry?.title || (en?'New chat':'新会话')),
        !needsAnswer && h('svg',{width:14,height:14,viewBox:'0 0 16 16',fill:'none','aria-hidden':true},
          h('path',{d:expanded?'M3 8h10':'M4 10l4-4 4 4',stroke:'currentColor',strokeWidth:1.4,strokeLinecap:'round',strokeLinejoin:'round'}))),
      expanded && h('div',{className:'office-chat-log',ref:log,role:'log'},...rows.map(r=>r.role==='deck' ? h(DeckCard,{
        key:r.key,model:r.model,activities:r.activities,en,waiting:needsAnswer,openFile:(_saved,path)=>{setOpen(false);openDeck(path);},
      }) : h('p',{key:r.key,className:'office-chat-'+r.role},
        r.role === 'user' && renderFactorySlot ? renderFactorySlot('conversation.user.text', { text:r.text, facts:{}, interactive:false }) : r.text)),
        (error || controller.state.parallelError) && h('p',{className:'office-chat-error',role:'alert'},error || controller.state.parallelError),
        recent && h('button',{type:'button',className:'office-chat-file',onClick:async()=>{jump();try{await controller.agent('changeDetail',{id:recent.id,action:'select'});}catch(e){setError(e.message);}}},file,' ↗')),
      needsAnswer && h('div',{className:'office-chat-decision-list',role:'region','aria-label':en?'Your response is needed':'等待你的操作'},
        waiting && h('div',{className:'office-chat-decisions'},
          renderFactorySlot('conversation.content',{variant:'embedded',phase:'active',hero:false},{slots:embeddedSlots})),
        proposal && h('div',{className:'office-chat-proposal'},proposal),
        otherDecisions.length>0 && h('div',{className:'office-chat-other-decisions'},...otherDecisions)),
      !needsAnswer && h('div',{className:'office-chat-compose office-prompt'},h('textarea',{ref:input,rows:1,wrap:'off',value:entry?.draft ?? '',disabled:!entry?.actions,placeholder:en?'Enter editing instructions…':'输入修改要求…',
        'aria-label':en?'Message':'输入消息',onChange:e=>entry.actions.setDraft(e.target.value),
        onKeyDown:e=>{if(e.key==='Enter' && !e.shiftKey && !e.nativeEvent.isComposing && e.nativeEvent.keyCode!==229){e.preventDefault();void send();}if(e.key==='Escape')setOpen(false);}}),
        h('button',{type:'button',className:'office-chat-send office-prompt-send',disabled:!ready,onClick:send,'aria-label':en?'Send':'发送'},
          h('svg',{width:18,height:18,viewBox:'0 0 20 20',fill:'none','aria-hidden':true},
            h('path',{d:'M10 16V4m-5 5 5-5 5 5',stroke:'currentColor',strokeWidth:1.65,strokeLinecap:'round',strokeLinejoin:'round'})))),
      null);
  }
  return {ChangeRail,FocusChat,FocusDecision};
}
