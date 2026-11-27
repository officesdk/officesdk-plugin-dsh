import assert from 'node:assert/strict';
import test from 'node:test';
import {wordTemplate,pptTemplate} from '../src/host/template-match.mjs';
import {newPresentationRequest,watchPresentationTurns} from '../src/client/presentation-preview.mjs';
import { AgentBridge } from '../src/client/agent.mjs';

test('a script output is reported open only after its own native editor is ready', async () => {
  let ready;
  const rendered = new Promise(resolve => { ready = resolve; });
  const controller = { whenReady: () => rendered };
  const bridge = new AgentBridge({
    controllers: {byAddress:new Map(), wait:async () => controller},
    openPreview:async () => ({opened:true}),
  });
  let finished = false;
  const opening = bridge.run('s',{kind:'read',path:'target.pptx',previewPhase:'ready'},new AbortController().signal)
    .then(result => { finished = true; return result; });
  await new Promise(resolve => setImmediate(resolve));
  assert.equal(finished,false);
  ready();
  assert.deepEqual(await opening,{opened:true});
});

test('the requested deliverable selects an installed, compatible template',()=>{
  const ids=['blank','gongwen','weekly-report','meeting-minutes','product-requirements','project-plan','project-proposal','formal-letter'];
  for(const [intent,expected] of [['整理本周的产品需求','product-requirements'],['例会记录','meeting-minutes'],['立项建议书','project-proposal'],['随手记','blank'],['董事会致函','formal-letter'],['周报','weekly-report'],['实施计划','project-plan'],['下发通知','gongwen']]) {
    assert.equal(wordTemplate(intent,'draft.docx',ids),expected);
  }
  const packs=[{id:'slide-002',count:8,terms:[]},{id:'slide-008',count:7,terms:[]},{id:'slide-000',count:1,terms:[]},{id:'design',terms:['活动策划']}];
  assert.equal(pptTemplate('员工培训','季度汇报.pptx',packs,7),'slide-008');
  assert.equal(pptTemplate('团队汇报','draft.pptx',packs,8),'slide-002');
  assert.equal(pptTemplate('空白幻灯片','draft.pptx',packs,1),'slide-000');
  assert.equal(pptTemplate('员工培训','draft.pptx',packs,4),undefined);
  assert.equal(pptTemplate('社区活动策划','draft.pptx',packs,4),'design');
});

test('a new deck request opens preparation once, never during history replay or an in-place edit',()=>{
  assert.equal(newPresentationRequest('把这个 PPT 的标题改一下',true),false);
  assert.equal(newPresentationRequest('如何制作 PPT？'),false);
  assert.equal(newPresentationRequest('给刚才的汇报换个配色，另存一份',true),true);
  assert.equal(newPresentationRequest('给新同事做一份培训 PPT'),true);
  let notify;const seen=[];let session='s';
  const node={kind:'user',time:8,turn:1,content:[{type:'text',text:'另存一份新的 PPT'}]};
  const entry={running:true,chat:{nodes:[node]}};
  const watch=watchPresentationTurns({startedAt:10,composers:{get:()=>entry,subscribeEntries:f=>{notify=f;return ()=>{};}},sidebarOf:()=>({mounted:{getSnapshot:()=>session}}),activePresentation:()=>true,begin:(id,t)=>seen.push([id,t.turn]),abandon(){}});
  notify('s');node.time=20;session='other';notify('s');session='s';notify('s');notify('s');
  assert.deepEqual(seen,[['s',1]]);
  watch.dispose();
});
