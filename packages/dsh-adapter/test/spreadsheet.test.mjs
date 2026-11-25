import test from 'node:test';
import assert from 'node:assert/strict';
import { copyFile, mkdtemp, mkdir, readFile, realpath, rm, writeFile } from 'node:fs/promises';
import { join } from 'node:path';
import { tmpdir } from 'node:os';
import { fileURLToPath } from 'node:url';
import { Context } from '@deepseek-ai/cordis';
import ShimoPreview from '../src/host/index.mjs';
import { templateNames } from '../src/host/document.mjs';
import { AgentBridge, ControllerRegistry } from '../src/client/agent.mjs';
import { workbookOutline } from '../src/client/spreadsheet-agent.mjs';
import { TOOL_VIEW_MESSAGES, callKey, createToolViewStore, spreadsheetRows, toolViewModel } from '../src/client/tool-views.mjs';
import { apply as applyTools } from '../src/tools/index.mjs';
import { fakeConverter } from '../../core/test/helpers.mjs';
import { SuiteError } from '../../core/src/index.mjs';

// No spreadsheet skill is bundled: the create path is exercised with a template directory of the test's own (a neutral workbook).
const SAMPLE_XLSX = fileURLToPath(new URL('../../../test/fixtures/sample.xlsx', import.meta.url));

async function load(t) {
  const root = await realpath(await mkdtemp(join(tmpdir(), 'suite-sheet-')));
  t.after(() => rm(root, { recursive: true, force: true }));
  const TEMPLATES = join(root, 'templates');
  await mkdir(TEMPLATES); await copyFile(SAMPLE_XLSX, join(TEMPLATES, 'blank.xlsx'));
  const workspace = join(root, 'ws');
  await mkdir(workspace); await mkdir(join(root, 'runtime')); await writeFile(join(root, 'runtime', 'preview.html'), '');
  const app = new Context();
  const services = {
    workspaceFiles: { stat: async () => ({ version: 'v1', size: 1 }), readBytes: async () => { throw new Error('not read'); } },
    webServer: { register: () => () => {} },
    connection: { admit: () => ({ rejection: 403 }) },
    typert: { lookups: { get: () => ({ resolve: async id => ({ sessionId: id, workspaceRoot: workspace }) }) } },
  };
  for (const [name, value] of Object.entries(services)) { app.provide(name); app.set(name, value); }
  const writer = {
    canWrite: () => true,
    create: async (scope, path, data) => {
      const target = join(workspace, path);
      await writeFile(target, data, { flag: 'wx' }).catch(error => { if (error.code === 'EEXIST') throw new SuiteError('exists', 'A file with that name already exists.'); throw error; });
      return { absolutePath: target, version: 'v1' };
    },
  };
  const environment = { converter: fakeConverter(), runtimeRoot: join(root, 'runtime'), tempRoot: join(root, 'tmp'), dataRoot: join(root, 'data'), writer, sheetTemplateRoot: TEMPLATES };
  class Plugin extends ShimoPreview { constructor(ctx, config) { super(ctx, config, environment); } }
  Plugin.inject = ShimoPreview.inject;
  await app.plugin(Plugin, {});
  await new Promise(resolve => setTimeout(resolve, 20));
  const service = app.get('shimoPreview');
  const client = async answer => { const command = await service.agents.next('s1'); service.agents.reply('s1', command.requestId, await answer(command)); return command; };
  return { service, workspace, client, TEMPLATES };
}

test('live Excel creation opens the bundled blank template when the spreadsheet skill is absent', async t => {
  const { service, workspace, client } = await load(t);
  service.skills = { names: ['shimo-docx', 'shimo-pptx'] };
  await assert.rejects(service.aiCreateSpreadsheet('s1',{path:'invalid.xlsx',outputArea:'A9999999'}),{code:'invalid-edit'});
  const task = service.aiCreateSpreadsheet('s1',{path:'blank.xlsx',outputArea:'B2:D8'});
  const initial = await client(() => ({ok:true,result:{format:'xlsx'}}));
  assert.equal(initial.creationArea,'B2:D8');
  assert.equal((await task).created,true);
  assert.deepEqual(await readFile(join(workspace,'blank.xlsx')),await readFile(new URL('../templates/blank.xlsx',import.meta.url)));
});

test('spreadsheet_create writes the blank workbook, opens it, and generates without a card', async t => {
  const { service, workspace, client, TEMPLATES } = await load(t);
  assert.deepEqual(templateNames(TEMPLATES, '.xlsx'), ['blank']);
  await assert.rejects(service.aiCreateSpreadsheet('s1', { path: 'a.csv' }), error => error.code === 'invalid-path' && /\.xlsx/.test(error.message));
  const creating = service.aiCreateSpreadsheet('s1', { path: '预算.xlsx' });
  await client(() => ({ ok: true, result: { format: 'xlsx', workbook: { sheets: [{ name: 'Sheet1', active: true }] } } }));
  const result = await creating;
  assert.deepEqual([result.created, result.template, result.generating], [true, 'blank', true]);
  assert.deepEqual(await readFile(join(workspace, '预算.xlsx')), await readFile(join(TEMPLATES, 'blank.xlsx')));
  const editing = service.aiEdit('s1', { path: '预算.xlsx', ops: [{ op: 'setValues', range: 'A1', values: [['x']] }], summary: 's' });
  const edit = await client(() => ({ ok: true, result: { applied: 1 } }));
  assert.equal(edit.auto, true);
  assert.equal((await editing).generating, true);
});

test('tools: spreadsheet_* call the service with the agent Session', async () => {
  const calls = [];
  const service = {
    aiCreateSpreadsheet: async (...args) => { calls.push(['create', ...args.slice(0, 2)]); return {}; },
    aiRead: async (...args) => { calls.push(['read', args[0], args[1]]); return {}; },
    aiEdit: async (...args) => { calls.push(['edit', ...args.slice(0, 2)]); return {}; },
    aiSave: async (sessionId, path) => { calls.push(['save', sessionId, path]); return {}; },
  };
  const tools = new Map();
  applyTools({ tools: { register: tool => tools.set(tool.name, tool) }, shimoPreview: service });
  const exec = { agent: { session: { id: 's9' } }, callId: 'c1' };
  await tools.get('spreadsheet_create').execute({ path: 'a.xlsx' }, exec);
  await tools.get('spreadsheet_read').execute({ path: 'a.xlsx' }, exec);
  await tools.get('spreadsheet_edit').execute({ path: 'a.xlsx', summary: 'x', operations: [] }, exec);
  await tools.get('spreadsheet_save').execute({ path: 'a.xlsx' }, exec);
  assert.deepEqual(calls, [['create', 's9', { path: 'a.xlsx', template: 'blank' }], ['read', 's9', 'a.xlsx'], ['edit', 's9', { path: 'a.xlsx', ops: [], summary: 'x', callId: 'c1' }], ['save', 's9', 'a.xlsx']]);
  assert.match(tools.get('spreadsheet_edit').description, /There is no rollback/);
  await assert.rejects(tools.get('spreadsheet_read').execute({ path: 'a.xlsx' }, {}), /Spreadsheet tools need an agent Session/);
});

const BOOK = { sheets: [{ id: 's1', name: 'Sheet1', active: true, rowCount: 100, columnCount: 20, cells: { range: { row: 0, column: 0, rowCount: 1, columnCount: 2 }, rows: [[{ text: '项目' }, { text: '8', formula: 'SUM(B2:B3)' }]] } }] };
function sheetController({ approve = true } = {}) {
  const calls = [];
  return {
    calls, savedVersion: 'v1', editRevision: 0, state: {},
    ensureMode: async () => {},
    agent: async (op, args) => { calls.push([op, args]); return op === 'inspect' ? BOOK : { applied: args.ops.length }; },
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

test('client: a workbook read is summarised with formulas, an edit shows the spreadsheet card, bad batches stop first', async () => {
  const controller = sheetController();
  const views = createToolViewStore();
  const { bridge, replies } = bridgeFor(controller, views);
  await bridge.serve('s1', { requestId: 'r1', kind: 'read', path: 'a.xlsx', callId: 'c1' });
  assert.deepEqual(replies[0].result.workbook.sheets[0], { name: 'Sheet1', active: true, rows: 100, columns: 20, used: 'A1:B1', values: [['项目', '8']], formulas: { B1: '=SUM(B2:B3)' } });
  assert.deepEqual(views.get(callKey('c1')).outline, { sheets: 1, active: 'Sheet1', used: 'A1:B1' });
  await bridge.serve('s1', { requestId: 'r2', kind: 'edit', path: 'a.xlsx', summary: '加合计', editId: 'e', ops: [{ op: 'setValues', range: 'A5', values: [['合计', '=SUM(B2:B4)']] }, { op: 'clearRange', range: 'C1:C3' }] });
  const proposal = controller.calls.find(([name]) => name === 'propose')[1];
  assert.deepEqual([proposal.kind, proposal.danger, proposal.rows.map(row => row.where)], ['spreadsheet', true, ['Sheet1!A5:B5', 'Sheet1!C1:C3']]);
  assert.deepEqual(replies[1].result, { applied: 2 });
  // The sheet the card showed goes along, so a sheet switch while the card is open cannot move the write.
  assert.deepEqual(controller.calls.find(([name]) => name === 'apply')[1].ops.map(op => op.sheet), ['Sheet1', 'Sheet1']);
  await bridge.serve('s1', { requestId: 'r3', kind: 'edit', path: 'a.xlsx', summary: 'x', ops: [{ op: 'setValues', range: 'ZZ1', values: [[1]] }] });
  assert.deepEqual([replies[2].ok, replies[2].code], [false, 'invalid-edit']);
  await bridge.serve('s1', { requestId: 'r4', kind: 'undo', path: 'a.xlsx', editId: 'e' });
  assert.deepEqual([replies[3].ok, replies[3].code], [false, 'unsupported-operation']);
  const declined = sheetController({ approve: false });
  const second = bridgeFor(declined);
  await second.bridge.serve('s1', { requestId: 'r5', kind: 'edit', path: 'a.xlsx', summary: 'x', ops: [{ op: 'setValues', range: 'A1', values: [[1]] }] });
  assert.deepEqual(second.replies[0].result, { declined: true });
  assert.deepEqual(workbookOutline({}), { sheets: 0, active: '', used: '' });
});

test('chat rows: spreadsheet tools have their own titles, range rows and kept steps', () => {
  const t = key => TOOL_VIEW_MESSAGES.zh[key];
  assert.deepEqual(spreadsheetRows(t, [{ op: 'setValues', range: 'A1', values: [['项目', '数量'], ['苹果', 3]] }, { op: 'formatRange', sheet: '明细', range: 'A1:B1', bold: true }, { op: 'addSheet', name: '汇总' }]).map(row => [row.label, row.where, row.text]),
    [['写入数值', 'A1', '项目 | 数量 / 苹果 | 3'], ['改格式', '明细!A1:B1', 'bold=true'], ['新建工作表', '汇总', '']]);
  const created = toolViewModel(t, { toolName: 'spreadsheet_create', phase: 'start', block: { argsRaw: '{"path":"预算.xlsx"}' } });
  assert.deepEqual([created.title, created.summary], ['正在新建表格', '模板 blank']);
  const read = toolViewModel(t, { toolName: 'spreadsheet_read', phase: 'result', block: { content: [{ type: 'text', text: JSON.stringify({ unsaved: false, workbook: { sheets: [{ name: 'Sheet1', active: true, used: 'A1:D9' }, { name: '汇总' }] } }) }], call: { argsRaw: '{"path":"a.xlsx"}' } } });
  assert.deepEqual([read.title, read.summary], ['读取了表格', '2 个工作表 · 当前 Sheet1 · 已用 A1:D9']);
  const stopped = toolViewModel(t, { toolName: 'spreadsheet_edit', phase: 'result', block: { isError: true, content: [{ type: 'text', text: 'Error: cancelled: Stopped by the user. Steps 1–2 of 4 were applied and stay in the editor: nothing was undone and nothing was saved.' }], call: { argsRaw: '{"path":"a.xlsx"}' } } });
  assert.deepEqual([stopped.state, stopped.lines], ['stopped', ['已完成第 1–2 步（共 4 步），这些改动留在编辑器里，没有撤销，也没有保存']]);
  const declined = toolViewModel(t, { toolName: 'spreadsheet_edit', phase: 'result', block: { content: [{ type: 'text', text: '{"declined":true}' }], call: { argsRaw: '{"path":"a.xlsx"}' } } });
  assert.equal(declined.title, '没有修改表格');
});

test('tools: without shimo-xlsx (the public build) the built-in empty workbook supports live creation', () => {
  const tools = new Map();
  applyTools({ tools: { register: tool => tools.set(tool.name, tool) }, shimoPreview: { skillIncluded: name => name !== 'shimo-xlsx' } });
  assert.ok(tools.has('spreadsheet_create'));
  for (const name of ['spreadsheet_read', 'spreadsheet_edit', 'spreadsheet_save', 'document_create', 'presentation_create']) assert.ok(tools.has(name), name);
  const all = new Map();
  applyTools({ tools: { register: tool => all.set(tool.name, tool) }, shimoPreview: { skillIncluded: () => true } });
  assert.ok(all.has('spreadsheet_create'));
});

test('tools read the package\'s own skills-catalog.json: the public catalog offers blank Excel creation and gongwen alone for Word (impl review R1-3/4)', () => {
  const publicSkills = { 'shimo-pptx': [{ version: '0.0.9', stylePacks: ['default'] }], 'shimo-docx': [{ version: '0.0.5', templates: ['gongwen'] }] };
  const tools = new Map();
  // The Host stub claims every skill: the packaged catalog decides.
  applyTools({ tools: { register: tool => tools.set(tool.name, tool) }, shimoPreview: { skillIncluded: () => true } }, {}, { packagedSkills: publicSkills });
  assert.ok(tools.has('spreadsheet_create') && tools.has('spreadsheet_read'));
  const description = tools.get('document_create').description;
  assert.match(description, /only template in this build/);
  assert.ok(!description.includes('"redhead"') && !description.includes('"basic"'), description);
  // A package that did carry a spreadsheet skill would offer spreadsheet_create.
  const withSheet = new Map();
  applyTools({ tools: { register: tool => withSheet.set(tool.name, tool) }, shimoPreview: {} }, {}, { packagedSkills: { ...publicSkills, 'shimo-xlsx': [{ version: '0.0.3', templates: ['blank'] }] } });
  assert.ok(withSheet.has('spreadsheet_create'));
});
