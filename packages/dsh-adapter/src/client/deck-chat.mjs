import { FILE_ICONS } from './inspiration-icons.mjs';

export const DECK_CHAT = 'officesdk-presentation-generation';
const parse = value => {
  if (typeof value !== 'string') return value;
  try { return JSON.parse(value); } catch { return undefined; }
};
const counted = args => typeof args?.path === 'string' && Number.isInteger(args.slideCount) && args.slideCount > 0;
const tableSaves = new Set(['document_finish_project_table', 'document_finish_meeting_table', 'document_finish_proposal_table']);
const createMatch = event => {
  const { name, arguments: raw, turn } = event.data;
  const args=parse(raw);
  if (Number.isInteger(turn) && ((name==='skill' && /^(office|shimo)-(docx|xlsx|pptx)$/.test(args?.name))
      || (/^(document|spreadsheet)_(read|create|edit|save)$/.test(name) || tableSaves.has(name) || name === 'presentation_save') && typeof args?.path==='string')) {
    return {id:String(turn),role:'start'};
  }
  return name === 'presentation_create' && Number.isInteger(turn) && counted(parse(raw))
    ? { id: String(turn), role: 'start' } : null;
};

const PPT_INTENT='officesdk-presentation-request';
const pptIntent = /(?:制作|生成|创建|新建|做|写|设计|准备|出|create|build|make|generate)[^。！？\n]{0,100}(?:ppt\b|pptx\b|幻灯片|演示文稿|slides?\b|presentation\b|slide deck\b)/i;
const presentationRequest = {
  kind:PPT_INTENT,
  match:{'user/message':event=>{
    const {source,content,id}=event.data;
    return source?.kind==='user' && id && content?.some(b=>b.type==='text' && pptIntent.test(b.text))
      ? {id:String(id),role:'start'} : null;
  }},
  start:(_context,match,reader)=>{
    const turn=match.location?.turn?.turn;
    return {seq:match.event.seq,turn,first:reader.previous(PPT_INTENT)?.state.turn!==turn};
  },
  update:context=>context.state,
  buildLocationData(context,scope,previous){
    const location=context.start?.location;
    if(scope!=='turn' || location?.kind!=='step' || location.step.step!==1 || !context.state?.first) return null;
    // One opening message owns the hint. Later steering belongs to the running turn.
    return previous ?? {kind:'turn',turn:location.turn.turn,key:PPT_INTENT,value:context.state.seq};
  },
};
function suiteFor(event){
  const name=event.data.name;
  if(name==='skill')return parse(event.data.arguments)?.name?.split('-').at(-1);
  return {presentation:'pptx',spreadsheet:'xlsx',document:'docx'}[name?.split('_')[0]];
}

// The focus window can identify Office work without inserting a row into the live DSH flow.
export function focusOfficeTurn(turn) {
  const intent=turn?.data?.get(PPT_INTENT);
  const suite=turn?.data?.get(DECK_CHAT) ?? (intent != null ? 'pptx' : undefined);
  if (!suite || !turn.start) return undefined;
  return {suite,seq:intent ?? turn.start.seq,time:turn.start.time};
}
export const deckChatDefinition = {
  kind: DECK_CHAT, target: 'chat',
  match: {'turn/start':event=>({id:String(event.data.turn),role:'start'}),'tool/call':createMatch,'tool/ptc-dispatch-start':createMatch},
  start: (_context,match)=>suiteFor(match.event) ?? null,
  update: (context,match)=>context.state || suiteFor(match.event) || null,
  buildLocationData(context,scope,previous) {
    const turn=context.start?.location?.turn;
    if(scope!=='turn' || !turn || !context.state)return null;
    if(previous?.turn===turn.turn && previous.value===context.state)return previous;
    return {kind:'turn',turn:turn.turn,key:DECK_CHAT,value:context.state};
  },
  buildViewNode(context) {
    const first = context.start;
    const turn=first?.location?.turn;
    if(turn?.status!=='closed')return null;
    const intent=turn.data?.get(PPT_INTENT);
    const suite=context.state || (intent && 'pptx');
    if(!suite)return null;
    const control=turn.data?.get('turn-process');
    return {key:context.key, id:context.id, kind:DECK_CHAT, target:'chat', groupTree:'turn',
      anchorSeq:(control?.controlAnchorSeq ?? ((intent ?? first.event.seq)+1)) - 0.12, location:first.location,
      visibility:'visible', data:{turn:turn.turn,suite,seq:intent ?? first.event.seq,time:first.event.time}};
  },
};

const decoded = new WeakMap();
function decode(block) {
  if (decoded.has(block)) return decoded.get(block);
  const args = parse(block.call?.argsRaw ?? block.argsRaw);
  const text = block.content?.find(item => item.type === 'text')?.text;
  const value = { name:block.name, args, result:block.isError ? undefined : parse(text),
    error:block.isError ? text?.split('\n')[0] || block.error?.message : undefined,
    time:block.time, order:block.seq, settled:block.kind === 'tool-result' };
  decoded.set(block, value);
  return value;
}

function pathKey(path, cwd) {
  if (typeof path !== 'string') return '';
  const windows = /^[A-Za-z]:[\\/]|^\\\\/.test(path) || /^[A-Za-z]:[\\/]|^\\\\/.test(cwd ?? '');
  const slash = value => windows ? value.replaceAll('\\','/') : value;
  let value = slash(path);
  if (!/^(?:\/|[A-Za-z]:\/)/.test(value) && cwd) value = slash(cwd)+'/'+value;
  const parts = [];
  for (const part of value.split('/')) {
    if (part === '..' && parts.length && parts.at(-1) !== '..') parts.pop();
    else if (part && part !== '.') parts.push(part);
  }
  return (value.startsWith('/') ? '/' : '') + (windows ? parts.join('/').toLowerCase() : parts.join('/'));
}

function officeFileCalls(snapshot,turn) {
  const calls = [];
  function visit(block) {
    if (!block) return;
    if (tableSaves.has(block.name) || /^(?:presentation|document|spreadsheet)_(?:create|edit|save|undo)$/.test(block.name)) calls.push(decode(block));
    for (const sub of block.subCalls ?? []) visit(sub);
  }
  for (const key of snapshot.locations.getTurn(turn)) {
    const node = snapshot.nodes.get(key);
    if (node?.kind === 'tool-call') visit(node.data.root);
  }
  calls.sort((a,b) => (a.time ?? 0) - (b.time ?? 0) || (a.order ?? 0) - (b.order ?? 0));
  return calls;
}

// Count acknowledged slide IDs. A retry, a read, and an intermediate batch enclosing zero finished
// pages all leave the count alone. Two deck creations in one turn keep separate file identities.
export function deckModels(snapshot, turn, cwd) {
  const location = snapshot?.timeline?.turns?.get(turn);
  if (!location) return [];
  const decks = new Map(), failed = [];
  for (const {name,args,result,error,time,settled} of officeFileCalls(snapshot,turn)) {
    if (!name.startsWith('presentation_')) continue;
    const key = pathKey(args?.path,cwd);
    if (name === 'presentation_create' && counted(args)) {
      const model = {turn,path:args.path,count:args.slideCount,start:location.start?.time ?? time,
        end:location.end?.time,closed:location.status==='closed', pages:new Map(), edits:new Map(),
        created:result?.created===true, failure:error, retries:[], lastActivity:time};
      if (error || result?.declined) { failed.push(model); continue; }
      decks.set(key,model);
      continue;
    }
    const model = decks.get(key);
    if (!model) continue;
    model.lastActivity = Math.max(model.lastActivity ?? 0,time ?? 0);
    if (error || result?.declined) { model.failure=error || 'declined'; model.retries.push(model.failure); continue; }
    if (!settled) continue;
    if (name==='presentation_edit' && result?.applied>0 && !result.pendingReview) {
      const touched = new Set(result.slideIds ?? []);
      const ids = (args.completedSlideIds ?? [...touched]).filter(id => touched.has(id));
      for (const id of ids) model.pages.set(id,(model.pages.get(id) ?? 0)+1);
      if (result.editId) model.edits.set(result.editId,ids);
      model.savedPath=undefined;
      model.writing=true;
      model.failure=undefined;
    }
    if (name==='presentation_undo' && result?.undone) {
      for (const id of model.edits.get(result.undone === true ? args.editId : result.undone) ?? []) {
        const n=model.pages.get(id)-1;
        if(n>0) model.pages.set(id,n); else model.pages.delete(id);
      }
      model.edits.delete(result.undone === true ? args.editId : result.undone);
      model.savedPath=undefined;
    }
    if (name==='presentation_save' && (result?.saved || result?.unchanged)) {
      model.savedPath = result.path || model.path;
      model.savedAt=time;
      model.failure=undefined;
    }
  }
  const models = [...decks.values()];
  if (!models.length && failed.length) models.push(failed.pop());
  if (models.length) models[0].retries.unshift(...failed.map(item=>item.failure).filter(Boolean));
  return models.map(({pages,edits,...model}) => ({...model,completed:Math.min(pages.size,model.count)}));
}

export function deckStatus(m, waiting, en) {
  if (waiting && !m.closed) return en?'Waiting for your response':'等待你确认';
  if (m.savedPath && m.completed === m.count) {
    const seconds=Math.max(0,Math.floor(((m.end ?? m.savedAt)-m.start)/1000));
    const duration=seconds>=60 ? (en?Math.floor(seconds/60)+'m '+seconds%60+'s':Math.floor(seconds/60)+' 分 '+seconds%60+' 秒') : seconds+(en?'s':' 秒');
    return en?'Completed · '+m.count+' slides · '+duration:'已完成 · '+m.count+' 页 · 用时 '+duration;
  }
  if (m.closed) return (m.savedPath?(en?'Saved':'已保存'):(en?'Generation stopped':'生成未完成'))+' · '+m.completed+'/'+m.count+(en?' slides':' 页');
  if (m.completed === m.count) return en?'Slides ready · Saving presentation':'幻灯片已写入 · 正在等待保存';
  if (m.writing) return en?'Creating slides · '+m.completed+'/'+m.count+' complete':'正在制作幻灯片 · 已完成 '+m.completed+'/'+m.count+' 页';
  return en?m.count+' slides planned · Preparing materials':'已确定 '+m.count+' 页，正在准备素材';
}

const toolLabels={
  skill:'加载技能',bash:'运行命令',glob:'查找文件',grep:'搜索内容',read:'读取文件',
  read_image:'查看图片',edit:'修改文件',write:'写入文件',present:'展示文件',
  working_directory:'读取工作目录',office_recalculate:'重算工作簿',office_render:'渲染工作表',
  load_workspace_dependencies:'加载工作区依赖',
};
const officeNouns={presentation:'演示文稿',spreadsheet:'工作表',document:'文稿'};
const officeVerbs={create:'创建',read:'读取',edit:'编辑',save:'保存',undo:'撤销修改'};

export function turnActivity(snapshot,turn,en=false){
  const items=[];
  function tool(block){
    if(!block)return;
    const {name}=block;
    const args=parse(block.call?.argsRaw ?? block.argsRaw);
    const [suite,action]=name.split('_');
    const label=en?name.replaceAll('_',' '):(toolLabels[name] ||
      (officeNouns[suite] && officeVerbs[action] && officeVerbs[action]+officeNouns[suite]) || name);
    const file=args?.file_path ?? args?.path;
    const target=args?.summary || args?.description || (name==='skill' && args?.name) ||
      args?.pattern || (typeof file==='string' && file.split(/[\\/]/).at(-1));
    const parts=[label,typeof target==='string' && target,block.isError && (en?'Failed':'失败')];
    items.push({key:block.callId ?? String(items.length),time:block.callTime ?? block.time,text:parts.filter(Boolean).join(' · ')});
    for(const sub of block.subCalls ?? [])tool(sub);
  }
  for(const key of snapshot.locations.getTurn(turn)){
    const node=snapshot.nodes.get(key);
    if(!node || node.visibility==='hidden')continue;
    const d=node.data;
    if(node.kind==='tool-call')tool(d.root);
    if(node.kind==='assistant-step'){
      const text=d.blocks.filter(b=>b.kind==='text').map(b=>b.text).join(' ').trim();
      if(text)items.push({key,time:d.time,text});
    }
    if(node.kind==='model-retry')for(const a of d.attempts){
      items.push({key:key+'-'+a.seq,time:a.time,text:(en?'Retrying request':'重试请求')+' · '+a.retry});
    }
    if(node.kind==='context' && d.producer){
      items.push({key,time:d.time,text:[en?'Load context':'加载上下文',d.producer.label].filter(Boolean).join(' · ')});
    }
    if(node.kind==='command')items.push({key,time:d.time,text:d.name || (en?'Run command':'执行命令')});
  }
  return items.sort((a,b)=>a.time-b.time);
}

function duration(ms,en){
  const s=Math.floor(Math.max(0,ms)/1000), minutes=Math.floor(s/60);
  return (minutes?minutes+(en?'m ':' 分 '):'')+(s%60)+(en?'s':' 秒');
}

function createLiveProgress(React){
  const h=React.createElement;
  function FocusRunStatus({start,en}){
    const [now,setNow]=React.useState(Date.now);
    React.useEffect(()=>{
      const id=setInterval(()=>setNow(Date.now()),1000);
      return ()=>clearInterval(id);
    },[]);
    return h('div',{className:'office-focus-run-status'},en?'Working · ':'处理中，用时 ',
      h('span',{role:'timer','aria-live':'off'},duration(now-start,en)));
  }
  return function LiveProgress({suite,start,en,waiting,activities,focus=false}){
    const title=suite==='pptx'?(en?'Creating slides':'正在制作幻灯片'):
      suite==='xlsx'?(en?'Working on the spreadsheet':'正在处理表格'):(en?'Working on the document':'正在处理文稿');
    const current=waiting?(en?'Waiting for your response':'等待你确认'):
      activities.at(-1)?.text || (en?'Preparing…':'正在准备…');
    return h(React.Fragment,null,h('div',{className:'office-task-heading'},title),
      h('details',{className:'office-current-step'},
        h('summary',null,h('span',{className:'office-current-label',title:current},current),
          h('svg',{className:'office-step-chevron',width:14,height:14,viewBox:'0 0 16 16',fill:'none','aria-hidden':true},
            h('path',{d:'M4.5 6.25 8 9.75l3.5-3.5',stroke:'currentColor',strokeWidth:1.3,strokeLinecap:'round',strokeLinejoin:'round'}))),
        h('ul',{'aria-label':en?'Activity':'操作记录'},...activities.map(a=>h('li',{key:a.key},a.text)))),
      focus && h(FocusRunStatus,{start,en}));
  };
}

function retryLabel(line, en, recovered) {
  if (!recovered) return line;
  if (/\bexists:/.test(line)) return en?'Retried with a new file name':'文件重名，已换名创建';
  if (/invalid-edit:/.test(line)) return en?'Corrected the slide edit and retried':'已修正写入参数并重试';
  return en?'Recovered from a failed call':'一次调用失败后已重试成功';
}

export function createDeckCard(React) {
  const h=React.createElement;
  const LiveProgress=createLiveProgress(React);
  return function DeckCard({model:m,waiting,en,openFile,native=false,activities=[]}) {
    if(!m.closed)return h('section',{className:'office-deck-chat'},h(LiveProgress,{suite:m.suite ?? 'pptx',start:m.start,waiting,en,activities,focus:true}));
    const name=(m.savedPath ?? m.path).split(/[\\/]/).at(-1);
    return h('section',{className:'office-deck-chat'},
      h('details',null,h('summary',null,deckStatus(m,waiting,en),h('span',{'aria-hidden':true},'⌄')),
        h('ul',null,
          m.created && h('li',null,en?'Created '+m.count+' slides':'已建立 '+m.count+' 页幻灯片'),
          m.writing && h('li',null,en?m.completed+' slides written':'已写入 '+m.completed+' 页内容'),
          ...m.retries.map((line,i)=>h('li',{key:i},retryLabel(line,en,Boolean(m.savedPath && m.completed===m.count)))),
          m.failure && !m.retries.includes(m.failure) && h('li',null,m.failure),
          m.savedPath && h('li',null,en?'Saved '+name:'已保存《'+name+'》')),
        native && h('details',{className:'office-deck-history'},h('summary',null,en?'Full transcript':'完整记录'))),
      m.savedPath && h('button',{type:'button',className:'office-deck-file',onClick:()=>openFile(m.savedPath,m.path)},
        h('img',{alt:'',width:28,height:28,src:FILE_ICONS.ppt}),h('span',null,name),h('span',{'aria-hidden':true},'↗')));
  };
}

const fileTypes={
  presentation:{type:'PowerPoint',icon:'ppt'},
  spreadsheet:{type:'Excel',icon:'excel'},
  document:{type:'Word',icon:'word'},
};

// DSH already renders explicitly presented files. Recover other Office outputs only from
// acknowledged saves; a later unsaved edit invalidates the earlier delivery for that path.
function missingOfficeResults(snapshot,turn,cwd) {
  const presented=new Set((turn.data.get('deliverables')?.presented ?? []).map(f=>pathKey(f.path,cwd)));
  const files=new Map(), changed=new Set();
  for(const {name,args,result,settled} of officeFileCalls(snapshot,turn.turn)){
    if(!settled || !result || result.declined || result.pendingReview || typeof args?.path!=='string')continue;
    const [suite,action]=name.split('_');
    const key=pathKey(args.path,cwd);
    if((action==='create' && result.created) || (action==='edit' && result.applied>0) || (action==='undo' && result.undone)){
      changed.add(key);
      files.delete(key);
    }
    if ((action === 'save' || tableSaves.has(name)) && (result.saved === true || (action === 'save' && result.unchanged === true && changed.has(key)))) {
      const path=typeof result.path==='string' && result.path.length ? result.path : args.path;
      if(!presented.has(key) && !presented.has(pathKey(path,cwd)))files.set(key,{path,...fileTypes[suite]});
    }
  }
  return [...files.values()];
}

export function registerDeckChat(ctx,React) {
  const h=React.createElement;
  function CompletedOfficeTurn({node,useChat}) {
    const turn=node.data.turn, snapshot=useChat(s=>s);
    const location=snapshot?.timeline?.turns?.get(turn);
    if (location?.status!=='closed') return null;
    const row='[data-chat-turn="'+turn+'"]';
    const flow='[data-chat-flow]:has('+row+' .office-turn-disclosure)';
    // Fold intermediate replies with the native receipts, leaving the final answer visible.
    const answerStep=location.data?.get('turn-process')?.answerStep;
    const assistantRows=snapshot.locations.getTurn(turn).map(key=>snapshot.nodes.get(key))
      .filter(n=>n?.kind==='assistant-step'&&n.data?.step!==answerStep&&n.location?.step?.step!==answerStep);
    const folded=flow+':has([data-turn-process="'+turn+'"][aria-expanded="false"]) ';
    const hide=assistantRows.filter(n=>Number.isInteger(answerStep)).map(n=>folded+'[data-chat-node-key="'+CSS.escape(n.key)+'"]');
    return h('div',{className:'office-turn-disclosure'},
      hide.length>0&&h('style',null,hide.join(',')+'{display:none!important}'));
  }

  function OfficeFileTail({turn,useChat,useSessions,sessionId,openFile,useContextMenu,t}) {
    const snapshot=useChat(s=>s);
    const cwd=useSessions(s=>s.byId[sessionId]?.cwd);
    const menu=useContextMenu?.();
    if(turn.status!=='closed' || !focusOfficeTurn(turn))return null;
    const files=missingOfficeResults(snapshot,turn,cwd);
    const english=t?.('lang')==='en';
    return h('div',{className:'office-result-tail','data-lang':english?'en':'zh'},
      ...files.map(file=>{
        const label=file.path.split(/[\\/]/).at(-1);
        const open=()=>openFile(file.path);
        const trigger=menu?.getTriggerProps({kind:'file',value:{sessionId,path:file.path}});
        return h('div',{className:'office-result-file',key:file.path},
          h('button',{type:'button',className:'office-result-identity',onClick:open,title:file.path,...trigger},
            h('span',{className:'office-result-icon'},h('img',{src:FILE_ICONS[file.icon],alt:'',width:24,height:24})),
            h('span',{className:'office-result-description'},
              h('strong',null,label),h('small',null,file.type))),
          h('div',{className:'office-result-actions'},
            h('button',{type:'button',onClick:open,'aria-label':(english?'Open ':'打开 ')+label},english?'Open':'打开'),
            trigger?.onContextMenu && h('button',{type:'button',className:'office-result-more',
              'aria-haspopup':'menu','aria-label':english?'More file actions':'更多文件操作',
              onClick:trigger.onContextMenu,onKeyDown:trigger.onKeyDown},
              h('svg',{width:12,height:12,viewBox:'0 0 16 16',fill:'none','aria-hidden':true},
                h('path',{d:'m4 6 4 4 4-4',stroke:'currentColor',strokeWidth:1.5,strokeLinecap:'round',strokeLinejoin:'round'})))));
      }));
  }
  ctx.inject?.(['uiConversation'],scope=>{
    scope.uiConversation.events.register(presentationRequest);
    scope.uiConversation.events.register(deckChatDefinition);
    scope.slots.inject('conversation.chat.node',()=>scope.slots.register({
      name:'conversation.chat.node',key:DECK_CHAT,locale:'shimoPreview',
    },CompletedOfficeTurn));
    scope.slots.inject('conversation.chat.turnTail',()=>scope.slots.register({
      name:'conversation.chat.turnTail',id:'officesdk-saved-files',locale:'shimoPreview',
    },OfficeFileTail));
  });
}
