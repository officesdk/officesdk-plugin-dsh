import { FILE_ICONS } from './inspiration-icons.mjs';

export const DECK_CHAT = 'officesdk-presentation-generation';
const parse = value => {
  if (typeof value !== 'string') return value;
  try { return JSON.parse(value); } catch { return undefined; }
};
const counted = args => typeof args?.path === 'string' && Number.isInteger(args.slideCount) && args.slideCount > 0;
const createMatch = event => {
  const { name, arguments: raw, turn } = event.data;
  return name === 'presentation_create' && Number.isInteger(turn) && counted(parse(raw))
    ? { id: String(turn), role: 'start' } : null;
};

// The transcript owns the lifetime of this seat, including history paging, session changes and reloads.
export const deckChatDefinition = {
  kind: DECK_CHAT, target: 'chat',
  match: { 'tool/call':createMatch, 'tool/ptc-dispatch-start':createMatch },
  start: (_context, match) => match.event.data.turn,
  update: context => context.state,
  buildViewNode(context) {
    const first = context.start;
    if (!first?.location?.turn) return null;
    const control = first.location.turn.data?.get('turn-process');
    return {key:context.key, id:context.id, kind:DECK_CHAT, target:'chat', groupTree:'turn',
      anchorSeq:(control?.controlAnchorSeq ?? first.event.seq) - 0.12, location:first.location,
      visibility:'visible', data:{turn:context.state, seq:first.event.seq, time:first.event.time}};
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

// Count acknowledged slide IDs. A retry, a read, and an intermediate batch enclosing zero finished
// pages all leave the count alone. Two deck creations in one turn keep separate file identities.
export function deckModels(snapshot, turn, cwd) {
  const location = snapshot?.timeline?.turns?.get(turn);
  if (!location) return [];
  const calls = [];
  function visit(block) {
    if (!block) return;
    if (/^presentation_(create|edit|save|undo)$/.test(block.name)) calls.push(decode(block));
    for (const sub of block.subCalls ?? []) visit(sub);
  }
  for (const key of snapshot.locations.getTurn(turn)) {
    const node = snapshot.nodes.get(key);
    if (node?.kind === 'tool-call') visit(node.data.root);
  }
  calls.sort((a,b) => (a.time ?? 0) - (b.time ?? 0) || (a.order ?? 0) - (b.order ?? 0));
  const decks = new Map(), failed = [];
  for (const {name,args,result,error,time,settled} of calls) {
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

function retryLabel(line, en, recovered) {
  if (!recovered) return line;
  if (/\bexists:/.test(line)) return en?'Retried with a new file name':'文件重名，已换名创建';
  if (/invalid-edit:/.test(line)) return en?'Corrected the slide edit and retried':'已修正写入参数并重试';
  return en?'Recovered from a failed call':'一次调用失败后已重试成功';
}

export function createDeckCard(React) {
  const h=React.createElement;
  return function DeckCard({model:m,waiting,en,openFile,native=false}) {
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

export function registerDeckChat(ctx,React) {
  const Card=createDeckCard(React);
  const h=React.createElement;
  function NativeDeckRow({node,useChat,useSessionStatus,deckSession,t,cwd,openFile}) {
    const turn=node.data.turn, snapshot=useChat(s=>s);
    const waiting=useSessionStatus?.(s=>Boolean(s?.get?.(deckSession)?.pendingInteraction));
    const models=React.useMemo(()=>deckModels(snapshot,turn,cwd),[snapshot,turn,cwd]);
    if (!models.length || !Number.isInteger(turn)) return null;
    // Scope the native process disclosure to its own mounted chat column and turn. Opening our
    // disclosure exposes the original receipts and file actions again. The composer stays mounted.
    const selector='[data-chat-flow]:has([data-chat-turn="'+turn+'"] .office-deck-native):not(:has([data-chat-turn="'+turn+'"] .office-deck-native .office-deck-chat > details[open] > .office-deck-history[open]))';
    const row=selector+' [data-chat-turn="'+turn+'"]';
    const kinds=['turn-process','assistant-step','tool-call','model-retry'];
    if (models.every(m=>m.savedPath)) kinds.push('turn-tail');
    if (models.some(m=>m.closed && (!m.savedPath || m.completed<m.count))) kinds.splice(kinds.indexOf('assistant-step'),1);
    const hidden=kinds.map(kind=>row+'[data-chat-flow-kind="'+kind+'"]').concat(row+'[data-chat-group-key]');
    if (!models[0].closed && snapshot.timeline.turnOrder.at(-1)===turn) hidden.push(selector+' [data-chat-running]');
    return h('div',{className:'office-deck-native'},h('style',null,hidden.join(',')+'{display:none!important}'),
      ...models.map(m=>h(Card,{key:m.path,model:m,en:t?.('lang')==='en',waiting,native:true,openFile:path=>openFile(path)})));
  }
  ctx.inject?.(['uiConversation'],scope=>{
    scope.uiConversation.events.register(deckChatDefinition);
    scope.slots.inject('conversation.chat.node',()=>scope.slots.register({
      name:'conversation.chat.node',key:DECK_CHAT,locale:'shimoPreview',inject:id=>({deckSession:id}),
    },NativeDeckRow));
  });
}
