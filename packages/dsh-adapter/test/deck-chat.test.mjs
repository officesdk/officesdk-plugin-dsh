import test from 'node:test';
import assert from 'node:assert/strict';
import {DECK_CHAT,deckChatDefinition,deckModels,deckStatus,registerDeckChat,turnActivity} from '../src/client/deck-chat.mjs';
import {focusChatRows} from '../src/client/edit-surfaces.mjs';

const path='演示.pptx', turn=1;
function call(name,args,result,seq,isError=false) {
  return {kind:'tool-call',data:{root:{kind:'tool-result',name:'presentation_'+name,
    call:{argsRaw:JSON.stringify({path,...args})},time:1000+seq,seq,isError,
    content:[{type:'text',text:isError?result:JSON.stringify(result)}]}}};
}
function snapshot(rows,end) {
  const location={turn,status:end?'closed':'open',start:{time:1000,seq:1},data:new Map([[DECK_CHAT,'pptx'],['officesdk-presentation-request',4]]),end:end && {time:end}};
  const nodes=new Map(rows.map((node,i)=>[String(i),node]));
  return {timeline:{turnOrder:[turn],turns:new Map([[turn,location]])},nodes,locations:{getTurn:()=>[...nodes.keys()]}};
}
const created=call('create',{slideCount:2},{created:true},5);
const saved=call('save',{}, {saved:true,path:'/work/'+path},12);
test('only a planned presentation creates one turn seat, including on replay',()=>{
  const match=deckChatDefinition.match['tool/call'];
  assert.equal(match({data:{turn,name:'presentation_create',arguments:JSON.stringify({path})}}),null);
  assert.equal(match({data:{turn,name:'presentation_edit',arguments:'{}'}}),null);
  const input={turn,name:'presentation_create',arguments:JSON.stringify({path,slideCount:2})};
  assert.deepEqual(match({data:input}),{role:'start',id:'1'});
  assert.deepEqual(match({data:{...input,callId:'retry'}}),{role:'start',id:'1'});
});
test('progress counts written IDs once, waits on intermediate batches, routes saves by path',()=>{
  const events=[call('create',{path:'已有.pptx',slideCount:2},'exists',2,true),created,
    call('edit',{completedSlideIds:[]},{applied:1,slideIds:['a']},6),
    call('edit',{completedSlideIds:['a']},'invalid edit',7,true),
    call('edit',{}, {applied:3,slideIds:['a'],editId:'e1'},8),
    call('edit',{}, {applied:2,slideIds:['a'],editId:'e2'},9),
    call('edit',{path:'other.pptx'}, {applied:3,slideIds:['b']},10)];
  let [m]=deckModels(snapshot(events),turn,'/work');
  assert.equal(m.completed,1);assert.equal(m.retries.length,2);assert.equal(m.savedPath,undefined);
  assert.equal(deckStatus(m,true,false),'等待你确认');
  events.push(call('edit',{completedSlideIds:['b']},{applied:1,slideIds:['b']},11),saved);
  [m]=deckModels(snapshot(events,75000),turn,'/work');
  assert.equal(m.savedPath,'/work/'+path);assert.equal(deckStatus(m,false,false),'已完成 · 2 页 · 用时 1 分 14 秒');
});
test('failed saves remain incomplete and a later edit invalidates an earlier save',()=>{
  const both=call('edit',{}, {applied:2,slideIds:['a','b']},8);
  for(const last of [call('save',{},'permission denied',13,true),call('edit',{},{applied:1,slideIds:['b']},14)]){
    const events=last.data.root.name==='presentation_save'?[created,both,last]:[created,both,saved,last];
    const [m]=deckModels(snapshot(events,1800),turn);
    assert.equal(m.savedPath,undefined);assert.match(deckStatus(m,false,false),/^生成未完成/);
  }
});
test('focused chat keeps the user input and a single deck row, bounded by the focus period',()=>{
  const snap=snapshot([created,call('edit',{}, {applied:2,slideIds:['a','b']},8),saved],1800);
  const chat={nodes:[{kind:'user',seq:2,time:1002,content:[{type:'text',text:'制作两页'}]},
    {kind:'assistant',turn,seq:7,time:1007,blocks:[{kind:'text',text:'过程'}]}]};
  assert.deepEqual(focusChatRows(chat,{since:999,lastSeq:0},false,false,snap).map(r=>r.role),['user','deck']);
  assert.deepEqual(focusChatRows(chat,{since:2000,lastSeq:14},false,false,snap),[]);
});
test('entering focus during generation shows the next acknowledgement, without replaying older work',()=>{
  const period={since:1007,lastSeq:7};
  const chat={nodes:[]};
  assert.equal(focusChatRows(chat,period,false,true,snapshot([created])).length,0);
  const updated=snapshot([created,call('edit',{}, {applied:1,slideIds:['a']},8)]);
  assert.equal(focusChatRows(chat,period,false,true,updated)[0]?.model.completed,1);
});

test('status lists actual calls including nested calls in start order and keeps public text',()=>{
  const root=call('create',{slideCount:2},{created:true},5).data.root;
  root.callId='batch';root.callTime=1001;root.time=1040;
  root.subCalls=[{name:'read_image',callId:'photo',time:1010,argsRaw:'{"file_path":"a.jpg"}'}];
  const snap=snapshot([{kind:'tool-call',data:{root}},{kind:'assistant-step',data:{time:1020,blocks:[
    {kind:'reasoning',text:'private'},{kind:'text',text:'检查版式'}]}}]);
  assert.deepEqual(turnActivity(snap,turn).map(a=>a.text),['创建演示文稿 · 演示.pptx','查看图片 · a.jpg','检查版式']);
});
test('live Office turns leave every row and process group to DSH, including PPT before its first tool',()=>{
  const turnInfo={turn,status:'open',data:new Map()};
  const start={location:{turn:turnInfo},event:{seq:2,time:900,data:{turn}}};
  for(const suite of [null,'docx','xlsx','pptx']){
    const context={start,state:suite};
    const data=deckChatDefinition.buildLocationData(context,'turn',null);
    if(suite)assert.equal(data.value,suite);
    assert.equal(deckChatDefinition.buildViewNode(context),null);
  }
  turnInfo.data.set('officesdk-presentation-request',3);
  assert.equal(deckChatDefinition.buildViewNode({state:null,start}),null);
  turnInfo.status='closed';
  const row=deckChatDefinition.buildViewNode({state:null,start});
  assert.equal(row.data.suite,'pptx');
  assert.equal(row.data.turn,turn);
  assert.ok(row.anchorSeq>3);
  turnInfo.data.clear();
  for(const suite of ['docx','xlsx','pptx'])assert.equal(deckChatDefinition.buildViewNode({state:suite,start}).data.suite,suite);
  assert.equal(deckChatDefinition.buildViewNode({state:null,start}),null);
});

test('completion folds only intermediate replies and leaves the native deliverable in place',t=>{
  if(!globalThis.CSS){
    Object.defineProperty(globalThis,'CSS',{configurable:true,value:{escape:value=>value}});
    t.after(()=>delete globalThis.CSS);
  }
  const renderers=new Map();
  const slots={inject:(_name,install)=>install(),register:(definition,component)=>renderers.set(definition.name,component)};
  registerDeckChat({inject:(_services,install)=>install({slots,uiConversation:{events:{register(){}}}})},
    {createElement:(type,props,...children)=>({type,props,children})});
  const snap=snapshot([created,saved],1800);
  for(const [key,step] of [['intermediate',1],['answer',2]]){
    snap.nodes.set(key,{kind:'assistant-step',key,data:{step}});
  }
  const info=snap.timeline.turns.get(turn);
  info.data.set('turn-process',{answerStep:2});
  info.data.set('deliverables',{presented:[{path:'/work/'+path}]});
  const props={node:{data:{turn}},useChat:selector=>selector(snap),cwd:'/work'};
  const complete=renderers.get('conversation.chat.node')(props);
  const css=complete.children.find(child=>child?.type==='style').children.join('');
  assert.match(css,/aria-expanded="false"/);
  assert.match(css,/data-chat-node-key="intermediate"/);
  assert.doesNotMatch(css,/data-chat-node-key="answer"/);
  assert.ok(!complete.children.some(child=>child?.type==='button'));
  const opened=[];
  const tail=()=>renderers.get('conversation.chat.turnTail')({...props,turn:info,sessionId:'test',
    useSessions:s=>s({byId:{test:{cwd:'/work'}}}),openFile:p=>opened.push(p)});
  assert.equal(tail().children.length,0);
  info.data.delete('deliverables');
  assert.equal(tail().children.length,1);
  const [identity,actions]=tail().children[0].children;
  assert.equal(identity.children[1].children[0].children[0],path);
  assert.equal(identity.children[1].children[1].children[0],'PowerPoint');
  actions.children[0].props.onClick();
  assert.deepEqual(opened,['/work/'+path]);
  for(const [index,suite] of ['document','spreadsheet'].entries()){
    const name=index?'财报.xlsx':'结论.docx';
    const [create,save]=[call('create',{path:name},{created:true},20+index),call('save',{path:name},{saved:true,path:'/work/'+name},22+index)];
    create.data.root.name=suite+'_create';save.data.root.name=suite+'_save';
    snap.nodes.set(name+'-create',create);snap.nodes.set(name+'-save',save);
  }
  assert.equal(tail().children.length,3);
  snap.nodes.set('edit-later',call('edit',{}, {applied:1},26));
  assert.equal(tail().children.length,2);
  info.status='open';
  assert.equal(tail(),null);
});
