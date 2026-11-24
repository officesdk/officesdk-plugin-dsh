import test from 'node:test';
import assert from 'node:assert/strict';
import { createSelectionFeature, isSend, isShortcut, SELECTION_TEXTS } from '../src/browser/selection-ui.mjs';
import { docxReadToken } from '../src/agent/docx.mjs';
import { modelEditRange, replaceInParagraph } from '../src/browser/selection-edit.mjs';
import { textChunks } from '../src/browser/selection.mjs';

function setup() {
  const ui = { handlers:{},texts:SELECTION_TEXTS.zh,render(state){this.state=structuredClone(state);},showChip(point){this.point=point;},hideChip(){},openPanel(){},closePanel(){},busy(){},hint(){},dispose(){} };
  const win={setTimeout,clearTimeout,addEventListener(){},removeEventListener(){},innerHeight:600};
  const doc={addEventListener(){},removeEventListener(){}};
  const selected={kind:'docx',anchor:{paragraphs:[0,0],start:0,end:3},original:'测试字',box:{left:20,top:50,width:30,height:10}};
  const events=[],writes=[];
  const agent={readSelection:async()=>selected,rewriteSelection:async args=>{writes.push(args);return {reply:{applied:true}};}};
  const feature=createSelectionFeature({ui,agent,window:win,document:doc,notify:(name,data)=>events.push({name,data})});
  return {feature,ui,selected,events,writes,agent};
}

test('submitted Word work belongs to this request from send to its matching answer, including before the acknowledgement', async () => {
  const {feature,ui,selected,events,agent} = setup();
  let disposed=0, passed;
  const handle={dispose(){disposed++;}};
  agent.beginSelectionWork=()=>handle;
  agent.rewriteSelection=async (_args,_signal,work)=>{passed=work;return {reply:{applied:true}};};
  await feature.refresh();ui.handlers.open();ui.handlers.submit('扩写');
  const request=events.at(-1).data;
  await feature.rewrite({...selected,id:'unrelated',original:'其他原文'});
  assert.equal(passed,undefined);
  assert.equal(disposed,0);
  await feature.rewrite({...selected,id:'sel-0000000a'});
  assert.equal(passed,handle);
  assert.equal(disposed,1);
  feature.status({requestId:request.requestId,ok:true,ids:[{number:request.references[0].number,id:'sel-0000000a'}]});
  assert.equal(ui.state.pending.length,0);
  feature.dispose();assert.equal(disposed,1);
});

test('an idle snapshot during dispatch keeps the placeholder, and a completed turn without a rewrite removes it', async () => {
  const {feature,ui,agent,events}=setup();
  let disposed=0;
  agent.beginSelectionWork=()=>({dispose(){disposed++;}});
  await feature.refresh();ui.handlers.open();ui.handlers.submit('润色');
  const request=events.at(-1).data;
  feature.composer({canSend:true});
  feature.status({requestId:request.requestId,ok:true,ids:[{number:request.references[0].number,id:'sel-0000000b'}]});
  feature.composer({canSend:false,reason:'running'});
  feature.composer({canSend:false,reason:'no-composer'});
  assert.equal(disposed,0);
  feature.composer({canSend:true});
  assert.equal(disposed,1);assert.equal(ui.state.pending.length,0);
  assert.ok(events.some(({name,data})=>name==='selection-update' && data.selection==='sel-0000000b' && data.state==='cancelled'));
});

test('local edit sends only the latest confirmed selection and consumes it after sending', async () => {
  const {feature,ui,selected,events} = setup();
  await feature.refresh();
  assert.deepEqual(ui.point,{x:35,y:50});
  ui.handlers.open();
  ui.handlers.draft('缩短');
  selected.anchor = {paragraphs:[1,1],start:0,end:3};
  selected.original = '另一处';
  await feature.refresh();
  ui.handlers.open();
  assert.equal(ui.state.draft,'');
  ui.handlers.submit('');
  assert.ok(!events.some(e => e.name === 'selection-request'));
  ui.handlers.draft('扩写');
  ui.handlers.submit('扩写');
  const request = events.at(-1).data;
  assert.equal(request.instruction,'扩写');
  assert.deepEqual(request.references.map(r => r.original),['另一处']);
  assert.deepEqual(request.references[0].anchor,selected.anchor);
  feature.status({requestId:request.requestId,ok:true,ids:[{number:request.references[0].number,id:'sel-00000002'}]});
  assert.equal(ui.state.pending[0].id,'sel-00000002');
  assert.equal(ui.state.draft,'');
  ui.handlers.submit('下一条输入');
  assert.equal(events.filter(e => e.name === 'selection-request').length,1);
});

test('cancel and remount detach the range, including old multi-region drafts', async () => {
  const {feature,ui,selected,events} = setup();
  await feature.refresh();
  ui.handlers.open();
  ui.handlers.draft('简化');
  const snapshot = events.at(-1).data.draftState;
  assert.deepEqual(snapshot.references,[]);
  ui.handlers.close();
  ui.handlers.submit('简化');
  assert.ok(!events.some(e => e.name === 'selection-request'));
  const next = setup();
  assert.equal(snapshot.draft,'');
  next.feature.hydrate({state:{...snapshot,draft:'旧指令',references:[{...selected,number:1},{...selected,number:2}]}});
  assert.equal(next.ui.state.draft,'');
  assert.deepEqual(next.ui.state.references,[]);
  next.ui.handlers.submit('简化');
  assert.ok(!next.events.some(e => e.name === 'selection-request'));
  assert.ok(isSend({key:'Enter'}) && !isSend({key:'Enter',shiftKey:true}) && !isSend({key:'Enter',keyCode:229}) && !isSend({key:'Enter',isComposing:true}));
  assert.ok(isShortcut({key:'J',metaKey:true}));
});

test('a host refresh and a failed dispatch preserve the current instruction for retry', async () => {
  const {feature,ui,events,agent} = setup();
  agent.beginSelectionWork = () => ({dispose(){}});
  await feature.refresh();
  ui.handlers.open();
  const snapshot = events.at(-1).data.draftState;
  ui.handlers.draft('扩写当前句');
  feature.hydrate({state:{...snapshot,draft:'过期指令'}});
  assert.equal(ui.state.draft,'扩写当前句');
  ui.handlers.submit(ui.state.draft);
  feature.status({requestId:events.at(-1).data.requestId,ok:false,reason:'failed'});
  feature.hydrate({state:snapshot});
  assert.equal(ui.state.draft,'扩写当前句');
  ui.handlers.submit(ui.state.draft);
  assert.equal(events.at(-1).data.instruction,'扩写当前句');
});

test('an acknowledgement for the first range cannot clear the next prompt', async () => {
  const {feature,ui,selected,events} = setup();
  await feature.refresh();
  ui.handlers.open();
  ui.handlers.draft('改标题');
  ui.handlers.submit('改标题');
  const first = events.at(-1).data;
  ui.handlers.close();
  selected.anchor = {paragraphs:[1,1],start:0,end:2};
  selected.original = '正文';
  await feature.refresh();
  ui.handlers.open();
  ui.handlers.draft('改正文');
  feature.status({requestId:first.requestId,ok:true,ids:[{number:first.references[0].number,id:'sel-00000001'}]});
  assert.equal(ui.state.draft,'改正文');
  ui.handlers.submit('改正文');
  assert.deepEqual(events.at(-1).data.references.map(r => r.original),['正文']);
});
test('a card follow-up targets its current revision and waits for dispatch acknowledgement', async () => {
  const {feature,ui,agent,events,selected}=setup();
  const record={id:'sel-00000001',before:'旧标题',after:'测试字'};
  feature.hydrate({state:{...ui.state,changes:[record]}});
  let selectedTimes=0;
  agent.selectChange=async card=>{assert.equal(card.id,record.id);selectedTimes++;return selected;};
  await feature.changeDetail({id:record.id,action:'select'});
  const send=()=>feature.changeDetail({id:record.id,action:'followup',instruction:'润色'});
  const first=send();
  await Promise.resolve();
  const sent=events.at(-1).data;
  assert.equal(selectedTimes,2);
  assert.deepEqual(sent.references.map(({original})=>original),[record.after]);
  await assert.rejects(send());
  feature.status({requestId:sent.requestId,ok:false,reason:'running'});
  await assert.rejects(first,/AI 正在回复/);
  const retry=send();
  await Promise.resolve();
  const next=events.at(-1).data;
  feature.status({requestId:next.requestId,ok:true,ids:[{number:next.references[0].number,id:'sel-00000002'}]});
  assert.deepEqual(await retry,{sent:true});
  agent.selectChange=async()=>{throw Error('原文位置已改变');};
  await assert.rejects(send(),/原文位置已改变/);
  assert.equal(events.filter(({name})=>name==='selection-request').length,2);
});

test('chat sends hold the same work handle across a model wait and discard it if no edit arrives',async()=>{
  const {feature,agent,selected}=setup(), handles=[];
  agent.chatSelection=async()=>selected;
  agent.beginSelectionWork=()=>{const h={dispose(){h.ended=true;}};handles.push(h);return h;};
  agent.rewriteSelection=async(args,signal,handle)=>{assert.equal(handle,handles[0]);return {reply:{applied:true}};};
  const running={reason:'running',chatEdit:{instruction:'润色这段'}};
  feature.composer(running);await Promise.resolve();
  feature.composer({reason:'running'});
  assert.equal(handles.length,1);assert.ok(!handles[0].ended);
  await feature.rewrite({...selected,id:'change',text:'新内容'});
  assert.equal(handles[0].ended,true);
  feature.composer({canSend:true});feature.composer(running);await Promise.resolve();
  feature.composer({canSend:true});assert.equal(handles[1].ended,true);
  feature.dispose();
});
test('model locations supersede late guesses, survive repeat reads and are consumed by their own result', async () => {
  const {feature,agent,selected,writes}=setup();
  const doc={paragraphs:[{text:'😀甲乙丙'},{text:'另一个位置'}],tables:[]};
  const readToken=docxReadToken(doc);
  const jobs=[];
  Object.assign(agent,{inspect:async()=>doc,modelEditRange,beginSelectionWork(ref){
    const handle={ref,dispose(){handle.closed=true;}};
    jobs.push(handle);return handle;
  }});
  let resolveGuess;
  agent.chatSelection=()=>new Promise(resolve=>{resolveGuess=resolve;});
  feature.composer({reason:'running',chatEdit:{instruction:'修改文章中提到的那件事'}});
  const mark=(paragraph,find)=>feature.inspectEditTarget({paragraph,find,readToken});
  assert.equal(await mark(0,'乙'),doc);
  await mark(1);await mark(0,'乙');
  resolveGuess(selected);await Promise.resolve();
  assert.equal(jobs.length,2);
  assert.deepEqual(jobs[0].ref.anchor,{paragraphs:[0,0],start:2,end:3});
  assert.deepEqual(writes,[]);
  agent.rewriteSelection=async(args,signal,handle)=>{assert.equal(handle,jobs[1]);return {reply:{applied:true}};};
  await feature.rewrite({...jobs[1].ref,id:'second',text:'新句'});
  assert.ok(jobs[1].closed);assert.ok(!jobs[0].closed);
  const claimed=feature.takeChatPreview().claim(jobs[0].ref);
  assert.ok(claimed);claimed.dispose();assert.ok(jobs[0].closed);
  await mark(1);feature.composer({canSend:true});assert.ok(jobs[2].closed);
  feature.dispose();
});
test('a stale, ambiguous or non-body location cannot start a chat work frame',async()=>{
  const {feature,agent}=setup();
  const doc={paragraphs:[{text:'甲甲'},{text:'表格',table:true}],tables:[]};
  const readToken=docxReadToken(doc);
  agent.inspect=async()=>doc;agent.modelEditRange=modelEditRange;
  agent.beginSelectionWork=()=>assert.fail('Unexpected frame');
  for (const target of [{paragraph:2},{paragraph:1},{paragraph:0,find:'甲'},{paragraph:0,find:'无'},{paragraph:0,readToken:'old'}]) {
    await assert.rejects(feature.inspectEditTarget({readToken,...target}));
  }
  feature.dispose();
});
test('accepting a change republishes its review state without rewriting the original',async()=>{
  const {feature,ui,events,writes}=setup();
  feature.hydrate({state:{...ui.state,changes:[{id:'a',after:'改文'},{id:'b',after:'另一处'}]}});
  assert.deepEqual(await feature.changeDetail({id:'a',action:'accept'}),{accepted:true});
  assert.deepEqual(events.at(-1).data.draftState.changes,[{id:'b',after:'另一处'}]);
  assert.deepEqual(writes,[]);
});
test('ask mode stages each answer without a write until explicit approval, including after a tab remount',async()=>{
  const {feature,ui,writes}=setup();
  feature.hydrate({permission:'ask'});
  for(const id of ['sel-00000001','sel-00000002'])assert.equal((await feature.rewrite({id,kind:'docx',text:'候选'})).pendingReview,true);
  assert.equal(writes.length,0);const next=setup();next.feature.hydrate({state:ui.state});
  await next.ui.handlers.decide('sel-00000001','restore');assert.equal(next.writes.length,0);
  await next.ui.handlers.decide('sel-00000002','always');assert.equal(next.writes.length,1);assert.equal(next.ui.state.candidates.length,0);
  await next.feature.rewrite({id:'sel-00000003',text:'另一个'});assert.equal(next.writes.length,2);
});
/** One paragraph's text and its Writer subranges (code points), every write recorded. */
function fakeWordParagraph(text) {
  const state = { text, writes: [] };
  const chars = () => Array.from(state.text);
  const subrange = (start, length) => ({
    insertText(value, where) {
      const all = chars();
      if (where === 'Replace') all.splice(start, length, ...Array.from(value));
      else if (where === 'After') all.splice(start + length, 0, ...Array.from(value));
      else all.splice(start, 0, ...Array.from(value));
      state.text = all.join('');
      state.writes.push([where, value]);
    },
    delete() { const all = chars(); all.splice(start, length); state.text = all.join(''); state.writes.push(['Delete']); },
  });
  const paragraph = { getRange: () => ({ getSubrange: subrange, getRange: () => subrange(0, 0) }) };
  const Word = { run: async fn => fn({ document: { body: { paragraphs: { load() {}, items: [paragraph] } } }, sync: async () => {} }) };
  return { Word, state };
}

test('an in-paragraph rewrite replaces only the selection, in growing parts, counting emoji as one', async () => {
  const { Word, state } = fakeWordParagraph('A😀B中文加粗尾巴');
  const seen = [];
  const text = '这是一段比二十四个字还要更长一些的新内容😀，用来分段写入结尾';
  await replaceInParagraph(Word, { index: 0, start: 3, length: 2, text, chunk: textChunks, onWrite: part => seen.push(part) });
  assert.equal(state.text, `A😀B${text}加粗尾巴`);
  assert.ok(seen.length > 1 && seen.at(-1) === text);
  // Taking it back writes the original in place again.
  await replaceInParagraph(Word, { index: 0, start: 3, length: Array.from(text).length, text: '中文' });
  assert.equal(state.text, 'A😀B中文加粗尾巴');
  // Rewriting to nothing deletes; taking that back inserts at the position.
  await replaceInParagraph(Word, { index: 0, start: 3, length: 2, text: '' });
  assert.equal(state.text, 'A😀B加粗尾巴');
  await replaceInParagraph(Word, { index: 0, start: 3, length: 0, text: '中文' });
  assert.equal(state.text, 'A😀B中文加粗尾巴');
  await assert.rejects(replaceInParagraph(Word, { index: 0, start: 0, length: 1, text: 'x', signal: AbortSignal.abort() }), error => error.code === 'cancelled');
});
