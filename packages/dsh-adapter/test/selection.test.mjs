import test from 'node:test';
import assert from 'node:assert/strict';
import { mkdtemp, mkdir, realpath, rm, writeFile } from 'node:fs/promises';
import { join } from 'node:path';
import { tmpdir } from 'node:os';
import { Context } from '@deepseek-ai/cordis';
import ShimoPreview from '../src/host/index.mjs';
import { SelectionStore } from '../src/host/selection.mjs';
import { AgentBridge, ControllerRegistry } from '../src/client/agent.mjs';
import { composerState, createComposerRegistry, handleSelectionRequest, pushComposerState } from '../src/client/selection-bridge.mjs';
import { SELECTION_REWRITE_DESCRIPTION } from '../src/tools/selection.mjs';
import { apply as applyTools } from '../src/tools/index.mjs';
import { fakeConverter } from '../../core/test/helpers.mjs';

const code = name => error => error?.code === name;

/** The Host over a real workspace folder with a .docx and a .xlsx in it; the test plays the Client. */
async function load(t, agent) {
  const root = await realpath(await mkdtemp(join(tmpdir(), 'suite-selection-')));
  t.after(() => rm(root, { recursive: true, force: true }));
  const workspace = join(root, 'ws');
  await mkdir(workspace); await mkdir(join(root, 'runtime')); await writeFile(join(root, 'runtime', 'preview.html'), '');
  await writeFile(join(workspace, '报告.docx'), 'x'); await writeFile(join(workspace, '销售.xlsx'), 'x');
  const app = new Context();
  const absolute = path => path.startsWith('/') ? path : join(workspace, path);
  const services = {
    workspaceFiles: { stat: async (scope, path) => ({ absolutePath: absolute(path), version: 'v1', size: 1 }), readBytes: async () => { throw new Error('not read'); } },
    webServer: { register: () => () => {} },
    connection: { admit: () => ({ rejection: 403 }) },
    typert: { lookups: { get: name => ({ resolve: async id => name === 'agent' ? agent : {sessionId:id,workspaceRoot:workspace} }) } },
  };
  for (const [name, value] of Object.entries(services)) { app.provide(name); app.set(name, value); }
  const environment = { converter: fakeConverter(), runtimeRoot: join(root, 'runtime'), tempRoot: join(root, 'tmp'), dataRoot: join(root, 'data'), writer: { canWrite: () => true } };
  class Plugin extends ShimoPreview { constructor(ctx, config) { super(ctx, config, environment); } }
  Plugin.inject = ShimoPreview.inject;
  await app.plugin(Plugin, {});
  await new Promise(resolve => setTimeout(resolve, 20));
  const service = app.get('shimoPreview');
  service.remotePeer = () => ({ id: 'tab', principal: 'owner', signal: new AbortController().signal, remote: { principal: { kind: 'owner' } } });
  const scope = sessionId => ({ sessionId, workspaceRoot: workspace });
  const client = async (answer, sessionId = 's1') => { const command = await service.agents.next(sessionId); service.agents.reply(sessionId, command.requestId, await answer(command)); return command; };
  return { service, workspace, scope, client };
}

const DOC_REQUEST = { kind: 'docx', anchor: { paragraphs: [2, 2], start: 4, end: 9 }, original: '第一句话。', instruction: '改得更正式' };
const SHEET_REQUEST = { kind: 'xlsx', anchor: { sheet: 'Sheet1', row: 1, column: 1, rowCount: 2, columnCount: 3 }, cells: Array.from({ length: 2 }, (_, r) => Array.from({ length: 3 }, (_, c) => ({ text: `v${r}${c}`, raw: `v${r}${c}` }))), instruction: '翻成英文' };

test('store: one selection, one rewrite; other Sessions, expiry and the user\'s reports', () => {
  let now = 0;
  const store = new SelectionStore({ now: () => now, random: (() => { let n = 0; return () => String(++n).padStart(8, '0'); })() });
  const record = store.create('s1', { path: 'a.docx', realPath: '/w/a.docx', name: 'a.docx', request: DOC_REQUEST });
  assert.equal(record.id, 'sel-00000001');
  assert.throws(() => store.get('s2', record.id), code('selection-unknown'));
  assert.equal(store.take('s1', record.id).state, 'applying');
  assert.throws(() => store.take('s1', record.id), code('busy'));
  store.set('s1', record.id, 'used');
  assert.throws(() => store.take('s1', record.id), code('selection-used'));
  // Kept and reverted are the user's word after a rewrite; cancelling only applies while it waits.
  assert.deepEqual(store.report('s1', record.id, 'cancelled'), { id: record.id, state: 'used' });
  assert.deepEqual(store.report('s1', record.id, 'reverted'), { id: record.id, state: 'reverted' });
  assert.throws(() => store.report('s1', record.id, 'used'), code('invalid-edit'));
  const waiting = store.create('s1', { path: 'a.docx', realPath: '/w/a.docx', name: 'a.docx', request: DOC_REQUEST });
  store.report('s1', waiting.id, 'cancelled');
  assert.throws(() => store.take('s1', waiting.id), code('cancelled'));
  const other = store.create('s1', { path: 'a.docx', realPath: '/w/a.docx', name: 'a.docx', request: DOC_REQUEST });
  assert.equal(store.staleFor('s1', '/w/b.docx'), 0);
  assert.equal(store.staleFor('s1', '/w/a.docx'), 1);
  assert.throws(() => store.take('s1', other.id), code('selection-stale'));
  const late = store.create('s1', { path: 'a.docx', realPath: '/w/a.docx', name: 'a.docx', request: DOC_REQUEST });
  now += 31 * 60_000;
  assert.throws(() => store.take('s1', late.id), code('selection-unknown'));
  // At most 20 per Session: the oldest go.
  const many = Array.from({ length: 25 }, () => store.create('s3', { path: 'a.docx', realPath: '/w/a.docx', name: 'a.docx', request: DOC_REQUEST }));
  assert.throws(() => store.get('s3', many[0].id), code('selection-unknown'));
  assert.equal(store.get('s3', many[24].id).state, 'pending');
});

test('Host: the tab records a selection and gets the whole message; selection_rewrite applies it once, without a card', async t => {
  const { service, scope, client } = await load(t);
  assert.deepEqual(/\(([^)]*)\)/.exec(ShimoPreview.prototype.selectionCreate.toString())[1].split(',').map(s => s.trim()), ['workspaceFileScope', 'path', 'selectionRequest', 'selectionLanguage', 'signal']);
  assert.deepEqual(/\(([^)]*)\)/.exec(ShimoPreview.prototype.selectionUpdate.toString())[1].split(',').map(s => s.trim()), ['workspaceFileScope', 'id', 'selectionState', 'signal']);
  const long = { ...DOC_REQUEST, anchor: { paragraphs: [2, 2], start: 0, end: 900 }, original: '很长的一句话'.repeat(150) };
  const { id, message } = await service.selectionCreate(scope('s1'), '报告.docx', long, 'zh');
  assert.match(id, /^sel-[0-9a-f]{8}$/);
  assert.ok(message.includes(id) && message.includes(long.original) && message.includes('要求：改得更正式') && message.includes('「报告.docx」'));
  await assert.rejects(service.selectionCreate(scope('s1'), '销售.xlsx', DOC_REQUEST, 'zh'), error => /officesdk\/invalid-edit/.test(error.code ?? error.name ?? '') || /does not match/.test(error.message));

  // The wrong kind of answer is refused and the selection keeps waiting.
  await assert.rejects(service.aiRewriteSelection('s1', { selection: id, values: [['x']] }), code('invalid-edit'));
  // Another Session cannot use it.
  await assert.rejects(service.aiRewriteSelection('s2', { selection: id, text: 'x' }), code('selection-unknown'));
  const rewriting = service.aiRewriteSelection('s1', { selection: id, text: '正式的句子。' });
  const command = await client(() => ({ ok: true, result: { applied: true, mode: 'inline', before: long.original, after: '正式的句子。' } }));
  assert.deepEqual([command.kind, command.path, command.text, command.auto, command.selection.id, command.selection.anchor], ['rewrite', '报告.docx', '正式的句子。', undefined, id, long.anchor]);
  const result = await rewriting;
  assert.deepEqual([result.applied, result.selection, result.unsaved], [true, id, true]);
  await assert.rejects(service.aiRewriteSelection('s1', { selection: id, text: '再改一次' }), code('selection-used'));
});

test('Host: spreadsheets take values of the selection\'s size; a changed selection is stale; a cancelled one is not applied', async t => {
  const { service, scope, client } = await load(t);
  const sheet = await service.selectionCreate(scope('s1'), '销售.xlsx', SHEET_REQUEST, 'en');
  assert.ok(sheet.message.includes('Sheet1!B2:D3') && sheet.message.includes('| v00 | v01 | v02 |'));
  await assert.rejects(service.aiRewriteSelection('s1', { selection: sheet.id, values: [['a', 'b']] }), code('invalid-edit'));
  await assert.rejects(service.aiRewriteSelection('s1', { selection: sheet.id, text: 'a' }), code('invalid-edit'));
  const writing = service.aiRewriteSelection('s1', { selection: sheet.id, values: [['a', 'b', 'c'], ['d', 'e', 1]] });
  const command = await client(() => ({ ok: false, code: 'source-changed', message: 'Cell B2 changed since it was selected.' }));
  assert.deepEqual(command.values, [['a', 'b', 'c'], ['d', 'e', 1]]);
  assert.deepEqual(command.selection.cells, SHEET_REQUEST.cells);
  await assert.rejects(writing, code('selection-stale'));
  await assert.rejects(service.aiRewriteSelection('s1', { selection: sheet.id, values: [['a', 'b', 'c'], ['d', 'e', 'f']] }), code('selection-stale'));
  // The user pressed × on the waiting selection.
  const doc = await service.selectionCreate(scope('s1'), '报告.docx', DOC_REQUEST, 'zh');
  assert.deepEqual(await service.selectionUpdate(scope('s1'), doc.id, 'cancelled'), { id: doc.id, state: 'cancelled' });
  await assert.rejects(service.aiRewriteSelection('s1', { selection: doc.id, text: 'x' }), code('cancelled'));
});

test('Host: a document_edit landing on the file makes its waiting selections stale', async t => {
  const { service, scope, client } = await load(t);
  const waiting = await service.selectionCreate(scope('s1'), '报告.docx', DOC_REQUEST, 'zh');
  const editing = service.aiEdit('s1', { path: '报告.docx', ops: [{ op: 'insertParagraph', text: 'x' }], summary: 's' });
  await client(() => ({ ok: true, result: { applied: 1 } }));
  await editing;
  await assert.rejects(service.aiRewriteSelection('s1', { selection: waiting.id, text: 'x' }), code('selection-stale'));
});

test('independent selections stream concurrently on the session model and commit in response order', async t => {
  const first = Promise.withResolvers(), second = Promise.withResolvers(), both = Promise.withResolvers();
  const requests = [];
  const timeline = [];
  const agent = { session:{id:'s1',append:(type,data)=>timeline.push({type,...data})}, model:{current:()=>({provider:'current-route',model:'current-model'})}, ctx:{llm:{
    async *stream(options) {
      requests.push(options);
      if (requests.length===2) both.resolve();
      const text=await (JSON.parse(options.messages[0].content[0].text).selection==='第一句话。'?first:second).promise;
      yield {type:'text-delta',text:JSON.stringify({text})};
      yield {type:'finish',reason:{kind:'stop'}};
    },
  }}};
  const {service,scope,client}=await load(t,agent);
  const a=await service.selectionCreate(scope('s1'),'报告.docx',{...DOC_REQUEST,trackRange:true},'zh');
  const b=await service.selectionCreate(scope('s1'),'报告.docx',{...DOC_REQUEST,instruction:'改写第二段的写法，保留原意',anchor:{paragraphs:[1,1],start:0,end:5},original:'第二句话。',trackRange:true},'zh');
  const runA=service.selectionParallel(scope('s1'),a.id), runB=service.selectionParallel(scope('s1'),b.id);
  await both.promise;
  assert.deepEqual(timeline.map(e=>e.type), ['command/run','command/run']);
  assert.equal(timeline[0].args, DOC_REQUEST.instruction+'\n\n> 编辑目标：报告.docx\n> 第 3 段\n> 第一句话。');
  assert.equal(timeline[1].args, '改写第二段的写法，保留原意\n\n> 编辑目标：报告.docx\n> 第 2 段\n> 第二句话。');
  const { parallelInputDefinition, parallelMessage } = await import('../src/parallel-message.mjs');
  const event = { type: timeline[1].type, seq: 8, data: timeline[1] };
  const state = parallelInputDefinition.start({}, {event});
  const replay = parallelInputDefinition.buildViewNode({id:b.id, key:b.id, state});
  assert.equal(parallelMessage(replay.data).text, timeline[1].args);
  const payload = JSON.parse(requests[1].messages[0].content[0].text);
  assert.equal(payload.instruction, '改写当前提供的选区的写法，保留原意');
  assert.deepEqual(payload.resolvedTarget, {document:'Word',firstParagraphIndex:1, lastParagraphIndex:1,startOffset:0,endOffset:5,indexBase:0});
  assert.ok(requests.every(r=>r.model==='current-model' && r.provider==='current-route' && r.sessionId==='s1'));
  assert.equal(service.selections.staleFor('s1',service.selections.get('s1',a.id).realPath),0);
  second.resolve('第二处先完成');
  const writtenB=await client(()=>({ok:true,result:{applied:true}}));
  assert.equal(writtenB.selection.id,b.id);assert.equal(writtenB.text,'第二处先完成');await runB;
  assert.equal(service.selections.get('s1',a.id).state,'pending');
  first.resolve('第一处后完成');
  const writtenA=await client(()=>({ok:true,result:{applied:true}}));
  assert.equal(writtenA.selection.id,a.id);assert.equal(writtenA.text,'第一处后完成');await runA;
  assert.deepEqual(timeline.slice(2).map(e=>[e.commandId,e.kind]), [[timeline[1].commandId,'success'],[timeline[0].commandId,'success']]);

});

test('tools: selection_rewrite passes the selection and the replacement with the agent Session', async () => {
  const tools = new Map();
  const calls = [];
  const ctx = { shimoPreview: { aiRewriteSelection: async (...args) => { calls.push(args); return { applied: true }; } }, tools: { register: tool => tools.set(tool.name, tool) } };
  applyTools(ctx);
  const tool = tools.get('selection_rewrite');
  assert.ok(tool && SELECTION_REWRITE_DESCRIPTION.includes('sel-') && /formatting/.test(tool.description));
  await tool.execute({ selection: 'sel-00000001', text: '新' }, { agent: { session: { id: 's9' } }, signal: undefined });
  await tool.execute({ selection: 'sel-00000002', values: [[1]] }, { agent: { session: { id: 's9' } } });
  assert.deepEqual(calls.map(([sessionId, args]) => [sessionId, args]), [['s9', { selection: 'sel-00000001', text: '新' }], ['s9', { selection: 'sel-00000002', values: [[1]] }]]);
});

/** A document tab as the bridge sees it: Host calls and page commands recorded. */
function fakeTab({ sessionId = 's1', fail = {} } = {}) {
  const calls = [], agents = [];
  return {
    calls, agents,
    source: { sessionId, path: 'docs/报告.docx' }, state: { ready: true, mode: 'edit' },
    call: async (method, args) => { calls.push([method, args]); if (fail[method]) throw new Error(fail[method]); return method === 'selectionCreate' ? { id: 'sel-0000000a', message: '修改「报告.docx」…' } : {}; },
    agent: async (op, args) => { agents.push([op, args]); return {}; },
  };
}
const actions = ({ submitFails = false, submitRejects = false, turn } = {}) => { const log = []; return { log, setDraft: text => log.push(['draft', text]), submit: () => { log.push(['submit']); if (submitFails) throw new Error('no'); if (submitRejects) return Promise.reject(new Error('later')); return turn ?? Promise.resolve(); } }; };
const now = run => run();

test('client: the request becomes the user\'s message when the input box is empty and the Session idle', async () => {
  const registry = createComposerRegistry();
  const tab = fakeTab();
  // DSH's submit settles when the turn is over: the page must hear before that.
  const box = actions({ turn: new Promise(() => {}) });
  registry.set('s1', 'pane', { actions: box, draft: '', attachments: 0, running: false });
  await handleSelectionRequest({ controller: tab, registry, defer: now, message: { requestId: 'q1', ...DOC_REQUEST, view: { texts: ['x'] } } });
  assert.deepEqual(tab.calls[0], ['selectionCreate', { path: 'docs/报告.docx', selectionRequest: { kind: 'docx', anchor: DOC_REQUEST.anchor, original: DOC_REQUEST.original, instruction: DOC_REQUEST.instruction }, selectionLanguage: 'zh' }]);
  assert.deepEqual(box.log, [['draft', '修改「报告.docx」…'], ['submit']]);
  assert.deepEqual(tab.agents, [['selectionStatus', { requestId: 'q1', ok: true, id: 'sel-0000000a' }]]);
});

test('client: a local edit with zero or multiple regions is never submitted', async () => {
  const registry = createComposerRegistry();
  const tab = fakeTab();
  const box = actions();
  registry.set('s1', 'pane', { actions: box, draft: '', attachments: 0, running: false });
  for (const references of [[], [DOC_REQUEST, DOC_REQUEST]]) {
    await handleSelectionRequest({ controller:tab, registry, defer:now, message:{
      kind:'region-batch', requestId:'q1', instruction:'缩短', references
    }});
    assert.deepEqual(tab.agents.at(-1),['selectionStatus',{requestId:'q1',ok:false,reason:'failed'}]);
  }
  assert.deepEqual(tab.calls,[]);
  assert.deepEqual(box.log,[]);
});

test('client: a draft, attachments, a running turn or no input box keep the request out; a failed send leaves no selection', async () => {
  const cases = [
    [{ draft: '写了一半', attachments: 0, running: false }, 'draft'],
    [{ draft: '', attachments: 1, running: false }, 'draft'],
    [{ draft: '', attachments: 0, running: true }, 'running'],
    [undefined, 'no-composer'],
  ];
  for (const [entry, reason] of cases) {
    const registry = createComposerRegistry();
    if (entry) registry.set('s1', 'pane', { actions: actions(), ...entry });
    const tab = fakeTab();
    await handleSelectionRequest({ controller: tab, registry, defer: now, message: { requestId: 'q2', ...DOC_REQUEST } });
    assert.deepEqual(tab.calls, [], reason);
    assert.deepEqual(tab.agents, [['selectionStatus', { requestId: 'q2', ok: false, reason }]]);
  }
  // Submit fails: the selection is cancelled on the Host and the draft goes back.
  const registry = createComposerRegistry();
  const box = actions({ submitFails: true });
  registry.set('s1', 'pane', { actions: box, draft: '', attachments: 0, running: false });
  const tab = fakeTab();
  await handleSelectionRequest({ controller: tab, registry, defer: now, message: { requestId: 'q3', ...DOC_REQUEST } });
  assert.deepEqual(tab.calls.map(([method, args]) => [method, args.id, args.selectionState]), [['selectionCreate', undefined, undefined], ['selectionUpdate', 'sel-0000000a', 'cancelled']]);
  assert.deepEqual(box.log.at(-1), ['draft', '']);
  assert.deepEqual(tab.agents.at(-1), ['selectionStatus', { requestId: 'q3', ok: false, reason: 'failed' }]);
  // The send fails later (the turn could not start): the selection is cancelled, the page already heard.
  const late = createComposerRegistry();
  late.set('s1', 'pane', { actions: actions({ submitRejects: true }), draft: '', attachments: 0, running: false });
  const later = fakeTab();
  await handleSelectionRequest({ controller: later, registry: late, defer: now, message: { requestId: 'q6', ...DOC_REQUEST } });
  await new Promise(resolve => setImmediate(resolve));
  assert.deepEqual(later.agents, [['selectionStatus', { requestId: 'q6', ok: true, id: 'sel-0000000a' }], ['selectionStatus', { dropped: 'sel-0000000a' }]]);
  assert.deepEqual(later.calls.at(-1), ['selectionUpdate', { id: 'sel-0000000a', selectionState: 'cancelled' }]);
  // The Host refuses: nothing is written into the input box.
  const refused = createComposerRegistry();
  const untouched = actions();
  refused.set('s1', 'pane', { actions: untouched, draft: '', attachments: 0, running: false });
  const failing = fakeTab({ fail: { selectionCreate: 'invalid' } });
  await handleSelectionRequest({ controller: failing, registry: refused, defer: now, message: { requestId: 'q4', ...DOC_REQUEST } });
  assert.deepEqual(untouched.log, []);
  assert.deepEqual(failing.agents.at(-1), ['selectionStatus', { requestId: 'q4', ok: false, reason: 'failed' }]);
});

test('client: "put it in the chat box" writes the message after the draft, unsent', async () => {
  const registry = createComposerRegistry();
  const box = actions();
  registry.set('s1', 'pane', { actions: box, draft: '我写了一半  ', attachments: 0, running: true });
  const tab = fakeTab();
  await handleSelectionRequest({ controller: tab, registry, defer: now, message: { requestId: 'q5', draftOnly: true, ...DOC_REQUEST } });
  assert.deepEqual(box.log, [['draft', '我写了一半\n\n修改「报告.docx」…']]);
  assert.deepEqual(tab.agents.at(-1), ['selectionStatus', { requestId: 'q5', ok: true, id: 'sel-0000000a', drafted: true }]);
});

test('client: the input box state reaches the Session\'s ready edit tabs when it changes', () => {
  const registry = createComposerRegistry();
  const seen = [];
  registry.subscribe(sessionId => seen.push([sessionId, registry.state(sessionId)]));
  registry.set('s1', 'a', { actions: actions(), draft: '', attachments: 0, running: false });
  registry.set('s1', 'a', { actions: actions(), draft: '', attachments: 0, running: false });
  registry.set('s1', 'a', { actions: actions(), draft: 'x', attachments: 0, running: false });
  registry.remove('s1', 'a');
  assert.deepEqual(seen, [['s1', { canSend: true }], ['s1', { canSend: false, reason: 'draft' }], ['s1', { canSend: false, reason: 'no-composer' }]]);
  assert.deepEqual(composerState({ actions: {} }), { canSend: false, reason: 'no-composer' });
  const ready = fakeTab(), other = fakeTab({ sessionId: 's2' }), viewing = fakeTab();
  viewing.state = { ready: true, mode: 'view' };
  registry.set('s1', 'b', { actions: actions(), draft: '', attachments: 0, running: true });
  pushComposerState([ready, other, viewing], registry, 's1');
  assert.deepEqual(ready.agents, [['composer', { canSend: false, reason: 'running' }]]);
  assert.deepEqual([other.agents, viewing.agents], [[], []]);
});

test('client: the bridge hands selection_rewrite to the tab\'s page without a card', async () => {
  const registry = new ControllerRegistry();
  const page = [];
  const controller = { ensureMode: async () => {}, agent: async (op, args) => { page.push([op, args]); return { applied: true }; }, propose: async () => { throw new Error('no card'); } };
  const bridge = new AgentBridge({ rpc: {}, openResource: address => registry.add(address, controller), controllers: registry, mounted: { getSnapshot: () => 's1', subscribe: () => () => {} } });
  const result = await bridge.run('s1', { kind: 'rewrite', path: 'docs/报告.docx', selection: { id: 'sel-0000000a', kind: 'docx', anchor: DOC_REQUEST.anchor, original: DOC_REQUEST.original }, text: '新句子' });
  assert.deepEqual(result, { applied: true });
  assert.deepEqual(page, [['rewrite', { id: 'sel-0000000a', kind: 'docx', anchor: DOC_REQUEST.anchor, original: DOC_REQUEST.original, text: '新句子' }]]);
});


test('a region edit uses its own message when chat holds a draft or an attachment', async () => {
  for (const [draft, attachments] of [['修改第一段', 0], ['', 1], ['扩写这一段', 0]]) {
    const registry = createComposerRegistry(), box = actions(), tab = fakeTab();
    registry.set('s1', 'pane', {actions:box, draft, attachments, running:false});
    await handleSelectionRequest({controller:tab, registry, message:{
      kind:'region-batch', requestId:'region', references:[{...DOC_REQUEST,number:1}], instruction:'扩写这一段',
    }});
    assert.deepEqual(tab.calls.map(([method])=>method), ['selectionCreate','selectionParallel']);
    assert.equal(tab.agents[0][1].independent, true);
    assert.deepEqual(box.log, []);
    assert.equal(registry.get('s1').draft, draft);
    assert.equal(registry.get('s1').attachments, attachments);
  }
});

test('typing into the main composer while selectionCreate is pending preserves the new draft', async () => {
  const registry = createComposerRegistry(), box = actions(), tab = fakeTab();
  registry.set('s1','pane',{actions:box, draft:'', attachments:0, running:false});
  const creation = Promise.withResolvers(), call = tab.call;
  tab.call = async (method, args) => {
    const result = await call(method,args);
    if (method === 'selectionCreate') await creation.promise;
    return result;
  };
  const sending = handleSelectionRequest({controller:tab,registry,message:{
    kind:'region-batch',requestId:'region',references:[{...DOC_REQUEST,number:1}],instruction:'缩短',
  }});
  registry.set('s1','pane',{actions:box,draft:'刚输入的草稿',attachments:0,running:false});
  creation.resolve();
  await sending;
  assert.deepEqual(tab.calls.map(([method])=>method), ['selectionCreate','selectionParallel']);
  assert.deepEqual(box.log, []);
});

test('a missing parallel endpoint ends only that request and tells the user how to recover', async () => {
  const registry = createComposerRegistry(), box = actions();
  registry.set('s1','pane',{actions:box,running:true,draft:'',attachments:0});
  const tab = fakeTab({fail:{selectionParallel:'transport failure for /api/shimoPreview/selectionParallel: HTTP 404'}});
  tab.set = patch => Object.assign(tab.state,patch);
  await handleSelectionRequest({controller:tab, registry, message:{kind:'region-batch',requestId:'two',
    references:[{...DOC_REQUEST,number:2}],instruction:'改写第二段'}});
  assert.match(tab.state.parallelError,/重启 DSH/);
  assert.ok(tab.agents.some(([op,args])=>op==='selectionStatus' && args.dropped==='sel-0000000a'));
  assert.deepEqual(box.log,[]);
});


test('focus history excludes backfilled messages and a pre-entry turn, but tracks new local edits and streaming text',async()=>{
  const {focusChatRows}=await import('../src/client/edit-surfaces.mjs');
  const period={since:100,lastSeq:7};
  const text=value=>[{type:'text',text:value}];
  const edit={kind:'command',name:'officesdk-edit',seq:16,time:104,args:'改写第二段'};
  const chat={turnTimings:new Map([[1,{startTime:82}],[2,{startTime:103}]]),nodes:[
    {kind:'user',seq:7,time:100,content:text('进入时已显示')},
    {kind:'user',seq:9,time:89,content:text('延迟加载的旧消息')},
    {kind:'assistant',seq:11,time:101,turn:1,blocks:text('早前开始的回答')},
    {kind:'user',seq:13,time:102,content:text('润色第一段')},edit,
  ],partial:{turn:2,blocks:text('本次回答')}};
  assert.deepEqual(focusChatRows(chat,period).map(r=>r.text),['润色第一段','改写第二段','正在修改…','本次回答']);
  edit.outcome={kind:'success'};chat.partial={turn:1,blocks:text('之前的流式回答')};
  assert.deepEqual(focusChatRows(chat,period).map(r=>r.text),['润色第一段','改写第二段','这处修改已完成。']);
  assert.deepEqual(focusChatRows(chat,{since:105,lastSeq:16}),[]);
});

test('focus chat surfaces the running turn, tool activity and failure without leaking the previous turn',async()=>{
  const {focusChatRows}=await import('../src/client/edit-surfaces.mjs');
  const period={since:100,lastSeq:5};
  const chat={turnTimings:new Map([[1,{startTime:50}]]),nodes:[]};
  const rows=running=>focusChatRows(chat,period,false,running).map(r=>r.text);
  assert.deepEqual(rows(true),[]);
  chat.turnTimings.set(2,{startTime:110});
  assert.deepEqual(rows(true),['正在处理…']);
  chat.runningCalls=[{turn:2}];
  assert.deepEqual(rows(true),['正在调用工具…']);
  chat.partial={turn:2,blocks:[{type:'text',text:'改写中'}]};
  assert.deepEqual(rows(true),['改写中']);
  chat.nodes=[{kind:'turn-error',seq:8,time:120,turn:1,message:'old'},{kind:'turn-error',seq:14,time:160,turn:2,message:'连接中断'}];
  chat.partial=null;chat.turnTimings.set(2,{startTime:110,endTime:160});
  assert.deepEqual(rows(false),['连接中断']);
});

test('the visible parallel instruction stays at its send position after a later completion', async () => {
  const { parallelInputDefinition: definition, parallelMessage } = await import('../src/parallel-message.mjs');
  const run = { type: 'command/run', seq: 10, data: { name: 'officesdk-edit', args: '改写第二段', commandId: 'cmd-officesdk-b' } };
  const done = { type: 'command/done', seq: 22, data: { commandId: 'cmd-officesdk-b', kind: 'success' } };
  const start = definition.start({}, { event: run });
  const updated = definition.update({ state: start }, { event: done });
  for (const state of [start, updated]) {
    const node = definition.buildViewNode({ state, key: 'input', id: 'b' });
    assert.equal(node.anchorSeq, 10);
    assert.equal(node.groupTree, 'input');
    assert.equal(parallelMessage(node.data).text, '改写第二段');
  }
});


test('a quoted parallel message preserves the selected lines and sheet address in the send language', async () => {
  const { parallelEditMessage } = await import('../../core/src/index.mjs');
  const multi = { kind:'docx', instruction:'润色', anchor:{paragraphs:[1,2]}, original:'甲\n\n> 乙' };
  const text = parallelEditMessage({name:'计划.docx', path:'草稿/计划.docx', request:multi, lang:'zh'});
  assert.equal(text,'润色\n\n> 编辑目标：计划.docx\n> 第 2–3 段\n> 甲\n> \n> > 乙\n> 草稿/计划.docx');
  const sheet = {kind:'xlsx', instruction:'Update these cells', anchor:{sheet:'Q4',row:1,column:1,rowCount:2,columnCount:2}, original:'1\t2\n3\t4'};
  assert.equal(parallelEditMessage({name:'Plan.xlsx', request:sheet, lang:'en'}),
    'Update these cells\n\n> Editing target: Plan.xlsx\n> Q4!B2:C3\n> 1\t2\n> 3\t4');
});
