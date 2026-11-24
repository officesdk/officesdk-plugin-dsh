import test from 'node:test';
import assert from 'node:assert/strict';
import {DECK_CHAT,deckChatDefinition,deckModels,deckStatus} from '../src/client/deck-chat.mjs';
import {focusChatRows} from '../src/client/edit-surfaces.mjs';

const path='演示.pptx', turn=1;
function call(name,args,result,seq,isError=false) {
  return {kind:'tool-call',data:{root:{kind:'tool-result',name:'presentation_'+name,
    call:{argsRaw:JSON.stringify({path,...args})},time:1000+seq,seq,isError,
    content:[{type:'text',text:isError?result:JSON.stringify(result)}]}}};
}
function snapshot(rows,end) {
  const location={turn,status:end?'closed':'open',start:{time:1000},end:end && {time:end}};
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
  const seat={kind:DECK_CHAT,data:{turn,seq:4,time:1004}};
  const snap=snapshot([seat,created,call('edit',{}, {applied:2,slideIds:['a','b']},8),saved],1800);
  const chat={nodes:[{kind:'user',seq:2,time:1002,content:[{type:'text',text:'制作两页'}]},
    {kind:'assistant',turn,seq:7,time:1007,blocks:[{kind:'text',text:'过程'}]}]};
  assert.deepEqual(focusChatRows(chat,{since:999,lastSeq:0},false,false,snap).map(r=>r.role),['user','deck']);
  assert.deepEqual(focusChatRows(chat,{since:2000,lastSeq:14},false,false,snap),[]);
});
test('entering focus during generation shows the next acknowledgement, without replaying older work',()=>{
  const seat={kind:DECK_CHAT,data:{turn,seq:4,time:1004}};
  const period={since:1007,lastSeq:7};
  const chat={nodes:[]};
  assert.equal(focusChatRows(chat,period,false,true,snapshot([seat,created])).length,0);
  const updated=snapshot([seat,created,call('edit',{}, {applied:1,slideIds:['a']},8)]);
  assert.equal(focusChatRows(chat,period,false,true,updated)[0]?.model.completed,1);
});
