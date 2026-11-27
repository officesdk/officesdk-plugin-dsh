import test from 'node:test';
import assert from 'node:assert/strict';
import { mkdtemp, mkdir, readFile, realpath, rm, writeFile } from 'node:fs/promises';
import { join } from 'node:path';
import { tmpdir } from 'node:os';
import { fileURLToPath } from 'node:url';
import { Context } from '@deepseek-ai/cordis';
import ShimoPreview from '../src/host/index.mjs';
import { loadTemplate, nounOf, templateNames } from '../src/host/document.mjs';
import { AgentBridge, ControllerRegistry } from '../src/client/agent.mjs';
import { documentOutline } from '../src/client/document-agent.mjs';
import { TOOL_VIEW_MESSAGES, callKey, createToolViewStore, documentProgress, documentRows, toolViewModel } from '../src/client/tool-views.mjs';
import { apply as applyTools } from '../src/tools/index.mjs';
import { fakeConverter } from '../../core/test/helpers.mjs';
import { SuiteError, docxReadToken, listZip, readZipEntry } from '../../core/src/index.mjs';

const TEMPLATES = fileURLToPath(new URL('../skills/shimo-docx/templates/', import.meta.url));

/** The Host over a real workspace folder, as in generate.test.mjs: new files are written on disk. */
async function load(t, { docTemplateRoot = TEMPLATES } = {}) {
  const root = await realpath(await mkdtemp(join(tmpdir(), 'suite-document-')));
  t.after(() => rm(root, { recursive: true, force: true }));
  const workspace = join(root, 'ws');
  await mkdir(workspace); await mkdir(join(root, 'runtime')); await writeFile(join(root, 'runtime', 'preview.html'), '');
  const app = new Context();
  const absolute = path => path.startsWith('/') ? path : join(workspace, path);
  const services = {
    workspaceFiles: { stat: async (scope, path) => ({ absolutePath: absolute(path), version: 'v1', size: 1 }), readBytes: async () => { throw new Error('not read'); } },
    webServer: { register: () => () => {} },
    connection: { admit: () => ({ rejection: 403 }) },
    typert: { lookups: { get: () => ({ resolve: async id => ({ sessionId: id, workspaceRoot: workspace }) }) } },
  };
  for (const [name, value] of Object.entries(services)) { app.provide(name); app.set(name, value); }
  const created = [];
  const writer = {
    canWrite: () => true,
    create: async (scope, path, data, signal, { approved = false } = {}) => {
      if (path.startsWith('/far/') && !approved) throw new SuiteError('sandbox-denied', 'outside', { outside: true });
      const target = path.startsWith('/far/') ? join(root, 'far', path.slice(5)) : absolute(path);
      await mkdir(join(target, '..'), { recursive: true });
      await writeFile(target, data, { flag: 'wx' }).catch(error => { if (error.code === 'EEXIST') throw new SuiteError('exists', 'A file with that name already exists.'); throw error; });
      created.push({ path, approved });
      return { absolutePath: target, version: 'v1' };
    },
  };
  const environment = { converter: fakeConverter(), runtimeRoot: join(root, 'runtime'), tempRoot: join(root, 'tmp'), dataRoot: join(root, 'data'), writer, docTemplateRoot };
  class Plugin extends ShimoPreview { constructor(ctx, config) { super(ctx, config, environment); } }
  Plugin.inject = ShimoPreview.inject;
  await app.plugin(Plugin, {});
  await new Promise(resolve => setTimeout(resolve, 20));
  const service = app.get('shimoPreview');
  const client = async answer => { const command = await service.agents.next('s1'); service.agents.reply('s1', command.requestId, await answer(command)); return command; };
  return { service, workspace, created, client };
}

const readResult = { ok: true, result: { format: 'docx', path: 'new.docx', unsaved: false, document: { paragraphCount: 1, paragraphs: [{ index: 0, text: '' }], tables: [], readToken: 'd1' } } };

test('the bundled template is gongwen, with the official page margins', async () => {
  assert.deepEqual(templateNames(TEMPLATES), ['blank','formal-letter','gongwen','meeting-minutes','product-requirements','project-plan','project-proposal','weekly-report']);
  const data = loadTemplate(TEMPLATES, 'gongwen');
  assert.ok(data.length > 1000);
  assert.throws(() => loadTemplate(TEMPLATES, 'letter'), error => error.code === 'invalid-edit' && /available: .*gongwen/.test(error.message));
  assert.equal(nounOf('a.docx'), 'document');
  assert.equal(nounOf('a.PPTX'), 'presentation');
  assert.equal(nounOf('a.xlsx'), 'spreadsheet');
});

test('document_create writes the template, opens it, and its edits apply without a card until Stop', async t => {
  const { service, workspace, created, client } = await load(t);
  const creating = service.aiCreateDocument('s1', { path: 'docs/new.docx' });
  const opened = await client(() => readResult);
  assert.deepEqual([opened.kind, opened.path, opened.creationArea], ['read', 'docs/new.docx', 'page']);
  const result = await creating;
  assert.deepEqual([result.created, result.template, result.generating, result.format], [true, 'blank', true, 'docx']);
  assert.deepEqual(await readFile(join(workspace, 'docs/new.docx')), await readFile(join(TEMPLATES, 'blank.docx')));
  assert.deepEqual(created, [{ path: 'docs/new.docx', approved: false }]);
  // Generating: the edit goes out with auto and the readToken the model passed.
  const editing = service.aiEdit('s1', { path: 'docs/new.docx', ops: [{ op: 'insertParagraph', text: '标题', style: 'title' }], summary: 's', callId: 'c1', readToken: 'd1' });
  const edit = await client(() => ({ ok: true, result: { applied: 1, readToken: 'd2' } }));
  assert.deepEqual([edit.kind, edit.auto, edit.readToken, edit.callId], ['edit', true, 'd1', 'c1']);
  assert.deepEqual(await editing, { applied: 1, readToken: 'd2', editId: edit.editId, unsaved: true, generating: true });
  // A stop from the tab ends the window: the next batch asks.
  await service.agents.alive?.('s1', 'x');
  const stopping = service.aiEdit('s1', { path: 'docs/new.docx', ops: [{ op: 'insertParagraph', text: '一' }], summary: 's' });
  await client(() => ({ ok: false, code: 'cancelled', message: 'Stopped by the user. Steps 1–0 of 1' }));
  await assert.rejects(stopping, { code: 'cancelled' });
  const asking = service.aiEdit('s1', { path: 'docs/new.docx', ops: [{ op: 'insertParagraph', text: '二' }], summary: 's' });
  const next = await client(() => ({ ok: true, result: { declined: true } }));
  assert.equal(next.auto, undefined);
  assert.deepEqual(await asking, { declined: true });
});

test('document_create refuses other extensions, existing files and unknown templates; outside the workspace it asks', async t => {
  const { service, client } = await load(t);
  await assert.rejects(service.aiCreateDocument('s1', { path: 'a.doc' }), { code: 'invalid-path' });
  await assert.rejects(service.aiCreateDocument('s1', { path: 'a.docx', template: 'nope' }), { code: 'invalid-edit' });
  const first = service.aiCreateDocument('s1', { path: 'a.docx', template: 'gongwen' });
  await client(() => readResult);
  assert.equal((await first).template, 'gongwen');
  await assert.rejects(service.aiCreateDocument('s1', { path: 'a.docx' }), { code: 'exists' });
  await assert.rejects(service.aiCreateDocument('s1', { path: '/far/b.docx' }, { approve: async () => 'rejected' }), error => error.code === 'sandbox-denied' && /did not allow/.test(error.message));
  const allowed = service.aiCreateDocument('s1', { path: '/far/c.docx' }, { approve: async reason => { assert.match(reason, /new document/); return 'allowed-once'; } });
  await client(() => readResult);
  assert.equal((await allowed).created, true);
});

test('a template name the model may still use (redhead, basic) starts from gongwen and says so; a misspelt one still fails (261113)', async t => {
  const { service, workspace, client } = await load(t);
  const creating = service.aiCreateDocument('s1', { path: 'notice.docx', template: 'redhead' });
  await client(() => readResult);
  const result = await creating;
  assert.equal(result.template, 'gongwen');
  assert.deepEqual([result.templateFallback.requested, result.templateFallback.used], ['redhead', 'gongwen']);
  assert.match(result.templateFallback.reason, /does not include the "redhead" template.*Tell the user/);
  assert.deepEqual(await readFile(join(workspace, 'notice.docx')), await readFile(join(TEMPLATES, 'gongwen.docx')));
  const basic = service.aiCreateDocument('s1', { path: 'plain.docx', template: 'basic' });
  await client(() => readResult);
  assert.equal((await basic).templateFallback.requested, 'basic');
  await assert.rejects(service.aiCreateDocument('s1', { path: 'x.docx', template: 'redhed' }), /Unknown template "redhed"; available: .*gongwen/);
});

test('tools: document_* call the service with the agent Session and pass readToken through', async () => {
  const calls = [];
  const service = {
    aiCreateDocument: async (...args) => { calls.push(['create', ...args.slice(0, 2)]); return { created: true }; },
    aiRead: async (...args) => { calls.push(['read', args[0], args[1], args[3]]); return { document: {} }; },
    aiEdit: async (...args) => { calls.push(['edit', ...args.slice(0, 2)]); return { applied: 1 }; },
    aiSave: async (sessionId, path, { approve }) => { calls.push(['save', sessionId, path]); return { outcome: await approve('write it') }; },
  };
  const tools = new Map();
  const asked = [];
  const approval = { request: async request => { asked.push(request); return 'allowed-once'; } };
  applyTools({ tools: { register: tool => tools.set(tool.name, tool) }, shimoPreview: service, get: name => name === 'approval' ? approval : undefined });
  const exec = { agent: { session: { id: 's9' } }, callId: 'c1', name: 'document_save' };
  await tools.get('document_create').execute({ path: 'a.docx' }, exec);
  await tools.get('document_read').execute({ path: 'a.docx' }, exec);
  await tools.get('document_edit').execute({ path: 'a.docx', summary: 'x', operations: [], readToken: 'd5' }, exec);
  await tools.get('document_edit').execute({ path: 'a.docx', summary: 'y', operations: [] }, exec);
  assert.deepEqual(await tools.get('document_save').execute({ path: '/far/a.docx' }, exec), { outcome: 'allowed-once' });
  assert.deepEqual(calls, [
    ['create', 's9', { path: 'a.docx' }], ['read', 's9', 'a.docx', { callId: 'c1' }],
    ['edit', 's9', { path: 'a.docx', ops: [], summary: 'x', callId: 'c1', readToken: 'd5' }], ['edit', 's9', { path: 'a.docx', ops: [], summary: 'y', callId: 'c1' }],
    ['save', 's9', '/far/a.docx'],
  ]);
  assert.match(asked[0].displayReason, /文档/);
  const edit = tools.get('document_edit');
  assert.match(edit.description, /There is no rollback/);
  assert.match(edit.description, /document_read before trying again/);
  assert.match(edit.description, /never change the file with any other tool/);
});

test('a targeted Word read conveys its passage from the model tool through the host without an edit', async t => {
  const {service,client}=await load(t),tools=new Map();
  applyTools({tools:{register:tool=>tools.set(tool.name,tool)},shimoPreview:service,get:()=>undefined});
  const editTarget={paragraph:3,find:'原文',readToken:'revision'};
  const result=tools.get('document_read').execute({path:'draft.docx',editTarget},{agent:{session:{id:'s1'}},callId:'target-read'});
  const command=await client(()=>({ok:true,result:{document:{readToken:'revision'}}}));
  assert.equal(command.kind,'read');assert.deepEqual(command.editTarget,editTarget);
  assert.equal(command.callId,'target-read');assert.equal((await result).document.readToken,'revision');
});
/** A tab for a Word document: inspect gives `doc`, apply answers what the page would. */
function wordController({ doc = { paragraphs: [{ text: '旧' }], tables: [] }, approve = true, apply } = {}) {
  const calls = [];
  return {
    calls, savedVersion: 'v1', editRevision: 0, state: {},
    ensureMode: async mode => calls.push(['mode', mode]),
    agent: async (op, args) => {
      calls.push([op, args]);
      if (op === 'inspect') return doc;
      if (op === 'apply') return apply ? apply(args) : { applied: args.ops.length, readToken: 'next', editRevision: 3 };
      throw new Error(`unexpected ${op}`);
    },
    flushEditor: async () => calls.push(['flush']),
    propose: async proposal => { calls.push(['propose', proposal]); return approve; },
    set(patch) { this.state = { ...this.state, ...patch }; },
  };
}

function bridgeFor(controller, views) {
  const replies = [];
  const registry = new ControllerRegistry();
  const rpc = { call: async (_, method, { args }) => { if (method.endsWith('/agentReply')) replies.push(args); return { status: 'success', value: null }; } };
  const bridge = new AgentBridge({ rpc, openResource: address => registry.add(address, controller), controllers: registry, mounted: { getSnapshot: () => 's1', subscribe: () => () => {} }, views, scheduler: { setTimeout: () => 0, clearTimeout() {} } });
  return { bridge, replies };
}

test('client: a Word read is summarised with numbers and a readToken, and recorded for its chat row', async () => {
  const doc = { paragraphs: [{ text: '通知' }, { text: '正文', list: true }], tables: [{ rows: 1, columns: 1, values: [['x']] }] };
  const controller = wordController({ doc });
  const views = createToolViewStore();
  const { bridge, replies } = bridgeFor(controller, views);
  await bridge.serve('s1', { requestId: 'r1', kind: 'read', path: 'a.docx', callId: 'c1' });
  const { result } = replies[0];
  assert.deepEqual(result.document.paragraphs, [{ index: 0, text: '通知' }, { index: 1, text: '正文', list: true }]);
  assert.equal(result.document.readToken, docxReadToken(doc));
  assert.equal(result.format, 'docx');
  assert.deepEqual(views.get(callKey('c1')), { outline: { paragraphs: 2, tables: 1, first: '通知' }, unsaved: false });
  assert.deepEqual(documentOutline({ paragraphs: [{ text: ' ' }, { text: 'x'.repeat(90) }] }).first.length, 80);
});

test('declaring a located passage returns just that passage at its document index', async () => {
  const doc = { paragraphs: [{ text: 'Other passage' }, { text: 'Located passage' }], tables: [] };
  const { bridge, replies } = bridgeFor(wordController({ doc }));
  const editTarget = { paragraph: 1, readToken: docxReadToken(doc) };
  await bridge.serve('s1', { requestId: 'r1', kind: 'read', path: 'a.docx', editTarget });
  const { document } = replies[0].result;
  assert.deepEqual(document.paragraphs, [{ index: 1, text: 'Located passage' }]);
  assert.equal(document.scope, 'editTarget');
  assert.equal(document.readToken, editTarget.readToken);
});

test('client: a Word edit shows the document card, applies with the readToken, and a decline never applies', async () => {
  const doc = { paragraphs: [{ text: '旧段落' }, { text: '留下的一段' }], tables: [] };
  const controller = wordController({ doc });
  const { bridge, replies } = bridgeFor(controller);
  const token = docxReadToken(doc);
  await bridge.serve('s1', { requestId: 'r1', kind: 'edit', path: 'a.docx', editId: 'e1', summary: '改一段', readToken: token, ops: [{ op: 'replaceText', paragraph: 0, find: '旧', replace: '新' }, { op: 'deleteParagraph', paragraph: 0 }] });
  const proposal = controller.calls.find(([name]) => name === 'propose')[1];
  assert.equal(proposal.kind, 'document');
  assert.equal(proposal.danger, true);
  assert.deepEqual(proposal.rows.map(row => [row.op, row.where]), [['replaceText', '¶1'], ['deleteParagraph', '¶1']]);
  assert.deepEqual(controller.calls.find(([name]) => name === 'apply')[1], { ops: [{ op: 'replaceText', paragraph: 0, find: '旧', replace: '新' }, { op: 'deleteParagraph', paragraph: 0 }], editId: 'e1', readToken: token });
  assert.deepEqual(replies[0].result, { applied: 2, readToken: 'next', editRevision: 3 });
  const declined = wordController({ doc, approve: false });
  const second = bridgeFor(declined);
  await second.bridge.serve('s1', { requestId: 'r2', kind: 'edit', path: 'a.docx', summary: 'x', readToken: token, ops: [{ op: 'setText', paragraph: 0, text: 'x' }] });
  assert.deepEqual(second.replies[0].result, { declined: true });
  assert.ok(!declined.calls.some(([name]) => name === 'apply'));
});

test('client: a stale or missing readToken and invalid batches stop before the card', async () => {
  const controller = wordController();
  const { bridge, replies } = bridgeFor(controller);
  await bridge.serve('s1', { requestId: 'r1', kind: 'edit', path: 'a.docx', summary: 'x', readToken: 'dold', ops: [{ op: 'setText', paragraph: 0, text: 'x' }] });
  await bridge.serve('s1', { requestId: 'r2', kind: 'edit', path: 'a.docx', summary: 'x', ops: [{ op: 'setText', paragraph: 0, text: 'x' }] });
  await bridge.serve('s1', { requestId: 'r3', kind: 'edit', path: 'a.docx', summary: 'x', ops: [{ op: 'insertParagraph', text: 'a\nb' }] });
  assert.deepEqual(replies.map(reply => [reply.ok, reply.code]), [[false, 'source-changed'], [false, 'invalid-edit'], [false, 'invalid-edit']]);
  assert.ok(!controller.calls.some(([name]) => name === 'propose' || name === 'apply'));
  await bridge.serve('s1', { requestId: 'r4', kind: 'undo', path: 'a.docx', editId: 'e' });
  assert.deepEqual([replies[3].ok, replies[3].code], [false, 'unsupported-operation']);
});

test('client: a generated Word file applies without a card and shows Stop; a failure passes the progress on', async () => {
  const controller = wordController({ apply: () => { throw Object.assign(new Error('Stopped by the user. Steps 1–1 of 2 were applied and stay in the editor: nothing was undone and nothing was saved.'), { code: 'officesdk/cancelled' }); } });
  const { bridge, replies } = bridgeFor(controller);
  await bridge.serve('s1', { requestId: 'r1', kind: 'edit', path: 'a.docx', summary: 'x', auto: true, ops: [{ op: 'insertParagraph', text: '一' }, { op: 'insertParagraph', text: '二' }] });
  assert.ok(!controller.calls.some(([name]) => name === 'propose'));
  assert.deepEqual([replies[0].ok, replies[0].code], [false, 'cancelled']);
  assert.match(replies[0].message, /Steps 1–1 of 2 were applied/);
  assert.ok(controller.state.generating, 'the tab offers Stop while the model generates');
});

test('chat rows: Word tools have their own titles, paragraph rows and the steps a stopped edit kept', () => {
  const t = key => TOOL_VIEW_MESSAGES.zh[key];
  const rows = documentRows(t, [{ op: 'insertParagraph', text: '一、总体要求', style: 'heading1', after: 2 }, { op: 'formatParagraph', paragraph: 0, bold: true }, { op: 'insertTable', values: [['a', 'b']] }]);
  assert.deepEqual(rows.map(row => [row.label, row.where, row.text]), [['新增段落', '第 3 段之后', 'heading1 · 一、总体要求'], ['改格式', '第 1 段', 'bold=true'], ['新增表格', '末尾', '1×2: a | b']]);
  assert.equal(documentProgress(t, 'cancelled: Stopped by the user. Steps 1–3 of 7 were applied and stay'), '已完成第 1–3 步（共 7 步），这些改动留在编辑器里，没有撤销，也没有保存');
  assert.equal(documentProgress(t, 'source-changed: x None of the 2 steps was applied.'), '一步也没有执行');
  const running = toolViewModel(t, { toolName: 'document_create', phase: 'start', block: { argsRaw: JSON.stringify({ path: 'a/通知.docx' }) } });
  assert.deepEqual([running.title, running.summary, running.file], ['正在新建文档', '', '通知.docx']);
  const read = toolViewModel(t, { toolName: 'document_read', phase: 'result', block: { content: [{ type: 'text', text: JSON.stringify({ unsaved: true, document: { paragraphs: [{}, {}, {}], tables: [{}] } }) }], call: { argsRaw: '{"path":"a.docx"}' } } });
  assert.deepEqual([read.title, read.summary], ['读取了文档', '共 3 段 · 1 个表格 · 有未保存的修改']);
  const edited = toolViewModel(t, { toolName: 'document_edit', phase: 'result', block: { content: [{ type: 'text', text: JSON.stringify({ applied: 4, unsaved: true, editId: 'e' }) }], call: { argsRaw: JSON.stringify({ path: 'a.docx', summary: '写第一节', operations: [] }) } } });
  assert.deepEqual([edited.title, edited.lines], ['修改了文档', ['已应用 4 步 · 还没保存到文件']], 'no "can be undone" for Word');
  const failed = toolViewModel(t, { toolName: 'document_edit', phase: 'result', block: { isError: true, content: [{ type: 'text', text: 'Error: editor-failed: The editor refused. Steps 1–2 of 3 were applied and stay in the editor: nothing was undone and nothing was saved.' }], call: { argsRaw: '{"path":"a.docx"}' } } });
  assert.deepEqual([failed.state, failed.title, failed.lines], ['error', '修改文档失败', ['已完成第 1–2 步（共 3 步），这些改动留在编辑器里，没有撤销，也没有保存']]);
  const ppt = toolViewModel(t, { toolName: 'presentation_edit', phase: 'start', block: { argsRaw: '{}' } });
  assert.equal(ppt.title, '正在修改演示文稿', 'presentation rows keep their titles');
});

test('261109: gongwen carries page numbers and starts with one empty paragraph', () => {
  const part = (name, path) => {
    const archive = listZip(loadTemplate(TEMPLATES, name));
    const entry = archive.entries.find(item => item.name === path);
    assert.ok(entry, `${name}: ${path}`);
    return readZipEntry(archive, entry, 5_000_000).toString('utf8');
  };
  assert.doesNotMatch(part('gongwen', 'word/styles.xml'), /outlineLvl/, 'Kit cannot save outline levels (K15)');
  assert.match(part('gongwen', 'word/document.xml'), /<w:body><w:p\/>/, 'gongwen starts with one empty paragraph');
});

test('a neutral filename and meeting intent open the live minutes template', async t => {
  const {service, client} = await load(t);
  const creation = service.aiCreateDocument('s1', {path:'draft.docx', intent:'整理例会记录'});
  const opening = await client(() => readResult);
  assert.equal((await creation).template, 'meeting-minutes');
  assert.equal(opening.creationArea, 'page');
});

test('project plan creation opens a real templated file and starts its table frame', async t => {
  const {service, workspace, client} = await load(t);
  const creating=service.aiCreateDocument('s1',{path:'project.docx',template:'project-plan',projectMetadata:{name:'演示项目',owner:'林悦',period:'本月'}});
  const opened=await client(async command => { assert.equal(command.kind,'read'); return readResult; });
  assert.equal(opened.creationArea,'page');
  assert.equal(opened.initialBlank,'project');
  const result=await creating; assert.equal(result.created,true);assert.equal(result.template,'project-plan');assert.equal(result.generating,true);
  const contents=await readFile(join(workspace,'project.docx'));assert.ok(contents.length>1000);
});
