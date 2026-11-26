import test from 'node:test';
import assert from 'node:assert/strict';
import { cp, mkdtemp, mkdir, readFile, realpath, rename, rm, writeFile } from 'node:fs/promises';
import { join } from 'node:path';
import { tmpdir } from 'node:os';
import { fileURLToPath } from 'node:url';
import { Context } from '@deepseek-ai/cordis';
import ShimoPreview, { GENERATION_LIMITS } from '../src/host/index.mjs';
import { AgentBridge, ControllerRegistry } from '../src/client/agent.mjs';
import { STYLE_CONTEXT, apply as applyTools } from '../src/tools/index.mjs';
import { fakeConverter } from '../../core/test/helpers.mjs';
import { PPTX_LIMITS, SuiteError } from '../../core/src/index.mjs';

const STYLE_ROOT = fileURLToPath(new URL('../skills/shimo-pptx/styles/', import.meta.url));
const PNG = Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0, 0, 0, 13]);

/** The Host over a real workspace folder: files are read and created on disk, as the Session's sandbox would allow. */
async function load(t, extra = {}) {
  const root = await realpath(await mkdtemp(join(tmpdir(), 'suite-generate-')));
  t.after(() => rm(root, { recursive: true, force: true }));
  const workspace = join(root, 'ws');
  await mkdir(workspace); await mkdir(join(root, 'runtime')); await writeFile(join(root, 'runtime', 'preview.html'), '');
  const app = new Context();
  const reads = [];
  const absolute = path => path.startsWith('/') ? path : join(workspace, path);
  const services = {
    workspaceFiles: {
      stat: async (scope, path) => { const data = await readFile(absolute(path)); return { absolutePath: absolute(path), version: 'v1', size: data.length }; },
      // Like DSH's WorkspaceFiles: a ranged read is at most `maxBytes` (2 MB); a whole file at most `maxFileBytes` (32 MB).
      readBytes: async (scope, path, options) => {
        reads.push(path);
        const range = options?.range;
        if (range && (range.length ?? 0) > 2 * 1024 * 1024) throw Object.assign(new Error('window too large'), { code: 'workspace-file/too-large' });
        const data = await readFile(absolute(path));
        if (!range && data.length > 32 * 1024 * 1024) throw Object.assign(new Error('file too large'), { code: 'workspace-file/too-large' });
        return { absolutePath: absolute(path), version: 'v1', data: range ? data.subarray(range.offset, range.offset + range.length) : data, eof: true };
      },
    },
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
      created.push({ path, approved, bytes: data.length });
      return { absolutePath: target, version: 'v1' };
    },
  };
  const environment = { converter: fakeConverter(), runtimeRoot: join(root, 'runtime'), tempRoot: join(root, 'tmp'), dataRoot: join(root, 'data'), writer, styleRoot: STYLE_ROOT, ...extra };
  class Plugin extends ShimoPreview { constructor(ctx, config) { super(ctx, config, environment); } }
  Plugin.inject = ShimoPreview.inject;
  await app.plugin(Plugin, {});
  await new Promise(resolve => setTimeout(resolve, 20));
  const service = app.get('shimoPreview');
  /** Plays the Client: collects the next command and answers it. */
  const client = async answer => { const command = await service.agents.next('s1'); service.agents.reply('s1', command.requestId, await answer(command)); return command; };
  return { service, workspace, created, client, reads };
}

const read = { ok: true, result: { format: 'pptx', path: 'new.pptx', unsaved: false, deck: { slides: [], layouts: [] } } };
const applied = { ok: true, result: { applied: 1, slideIds: ['s1'], refs: { p: 's1' }, undoable: 'x' } };

test('fixed training and blank styles return lossless JSON and open with generation masking', async t => {
  const { service, client } = await load(t);
  for (const [style, count] of [['slide-008', 7], ['slide-000', 1]]) {
    const creating = service.aiCreate('s1', { path: `${style}.pptx`, style, slideCount: count });
    const command = await client(() => read);
    const result = await creating;
    assert.equal(command.initialBlank, 'ppt-template');
    assert.equal(result.style.fixedDeck.slideCount, count);
    assert.deepEqual(result, JSON.parse(JSON.stringify(result)));
  }
});

test('presentation_create ends the Session\'s style recommendation: the strip has nothing left to ask', async t => {
  const { service, client } = await load(t);
  await service.aiSuggestStyles('s1', { candidates: [{ style: 'default', reason: '稳妥' }] });
  assert.ok(service.styleSuggestions.bySession.has('s1'));
  const creating = service.aiCreate('s1', { path: 'new.pptx' });
  await client(() => read);
  await creating;
  assert.equal(service.styleSuggestions.bySession.has('s1'), false);
});

test('a slide plan reaches the opening tab before the first content batch', async t => {
  const { service, client, created } = await load(t);
  await assert.rejects(service.aiCreate('s1',{path:'bad.pptx',slideCount:2,slideLayouts:['cover']}), {code:'invalid-edit'});
  assert.equal(created.length,0);
  const opening = service.aiCreate('s1',{path:'planned.pptx',slideCount:2,slideLayouts:['cover','closing']});
  const open = await client(() => read);
  assert.deepEqual([open.kind,open.slideCount,open.slideLayouts],['read',2,['cover','closing']]);
  await opening;
  const writing = service.aiEdit('s1',{path:'planned.pptx',ops:[{op:'setText',slideId:'s1',shapeId:'title',text:'Title'}],completedSlideIds:[]});
  const edit = await client(() => applied);
  assert.deepEqual(edit.completedSlideIds,[]);
  await writing;
});

test('presentation_create writes a style pack template, opens it, and the file is generated without a card per batch', async t => {
  const { service, workspace, created, client } = await load(t);
  const creating = service.aiCreate('s1', { path: 'new.pptx' });
  const opened = await client(() => read);
  assert.deepEqual([opened.kind, opened.path], ['read', 'new.pptx']);
  const result = await creating;
  assert.equal(result.created, true);
  assert.equal(result.style.name, 'default');
  assert.equal(result.style.palette.primary, '#2826C9');
  assert.equal(result.generating, true);
  assert.deepEqual(await readFile(join(workspace, 'new.pptx')), await readFile(join(STYLE_ROOT, 'default', 'template.pptx')), 'the default style is the default');
  // Edits to the created file go out as `auto`; another file still asks.
  const editing = service.aiEdit('s1', { path: 'new.pptx', ops: [{ op: 'addSlide', ref: 'p' }], summary: 'cover' });
  const auto = await client(() => applied);
  assert.equal(auto.auto, true);
  assert.equal((await editing).generating, true);
  await writeFile(join(workspace, 'other.pptx'), 'x');
  const other = service.aiEdit('s1', { path: 'other.pptx', ops: [{ op: 'addSlide' }], summary: 'x' });
  assert.equal((await client(() => applied)).auto, undefined);
  assert.equal((await other).generating, undefined);
  // The same file through another spelling of its path is the same window.
  const spelled = service.aiEdit('s1', { path: './new.pptx', ops: [{ op: 'addSlide' }], summary: 'x' });
  assert.equal((await client(() => applied)).auto, true);
  await spelled;
  // Another Session never inherits the window.
  assert.equal(await service.generationOf('s2', { workspaceRoot: workspace }, 'new.pptx'), undefined);
  await assert.rejects(service.aiCreate('s1', { path: 'new.pptx' }), { code: 'exists' });
  await assert.rejects(service.aiCreate('s1', { path: 'notes.docx' }), { code: 'invalid-path' });
  await assert.rejects(service.aiCreate('s1', { path: 'x.pptx', style: 'fancy' }), { code: 'invalid-edit', message: /Unknown style "fancy"; available: (?:[a-z0-9-]+（[^）]+）, )*default（蓝色品牌）/ });
  assert.deepEqual(created.map(item => item.path), ['new.pptx']);
  // The inspiration library writes display names: one names its pack.
  const named = service.aiCreate('s1', { path: 'named.pptx', style: '蓝色品牌' });
  await client(() => read);
  const result2 = await named;
  assert.deepEqual([result2.style.name, result2.styleSource], ['default', 'given']);
  assert.deepEqual(await readFile(join(workspace, 'named.pptx')), await readFile(join(STYLE_ROOT, 'default', 'template.pptx')));
});

test('the window ends when the user stops it, when a batch is cancelled, or when its budget runs out', async t => {
  const { service, workspace, client } = await load(t);
  const make = async name => { const creating = service.aiCreate('s1', { path: name, style: 'default' }); await client(() => read); await creating; };
  const autoOf = async (path, ops = [{ op: 'addSlide' }]) => { const editing = service.aiEdit('s1', { path, ops, summary: 'x' }); const command = await client(() => applied); await editing; return command.auto === true; };
  // Stop from the tab (agentStop).
  await make('a.pptx');
  assert.equal(await readFile(join(workspace, 'a.pptx')).then(data => data.length), (await readFile(join(STYLE_ROOT, 'default', 'template.pptx'))).length);
  service.remotePeer = () => ({ id: 'tab', principal: 'owner', signal: new AbortController().signal });
  service.suite.authorize = async () => {};
  assert.deepEqual(await service.agentStop({ sessionId: 's1', workspaceRoot: workspace }, 'a.pptx'), { stopped: true });
  assert.equal(await autoOf('a.pptx'), false);
  // A batch cancelled while it ran (Stop pressed during it).
  await make('b.pptx');
  const editing = service.aiEdit('s1', { path: 'b.pptx', ops: [{ op: 'addSlide' }], summary: 'x' });
  await client(() => ({ ok: false, code: 'cancelled', message: 'The edit was cancelled; nothing was changed.' }));
  await assert.rejects(editing, { code: 'cancelled' });
  assert.equal(await autoOf('b.pptx'), false);
  // The budget: a batch that would pass it asks.
  await make('c.pptx');
  // Filler steps that change one slide (a batch adds at most four new slides).
  const filler = () => ({ op: 'setBackground', slideId: 's1', color: '#FFFFFF' });
  const half = Array.from({ length: GENERATION_LIMITS.maxSteps / 2 }, filler);
  assert.equal(await autoOf('c.pptx', half), true);
  assert.equal(await autoOf('c.pptx', half), true);
  assert.equal(await autoOf('c.pptx'), false);
  // Batches still running count too: two parallel batches cannot both take the last of the budget.
  await make('d.pptx');
  await autoOf('d.pptx', half);
  const near = Array.from({ length: GENERATION_LIMITS.maxSteps / 2 - 1 }, filler);
  const first = service.aiEdit('s1', { path: 'd.pptx', ops: near, summary: 'x' });
  const second = service.aiEdit('s1', { path: 'd.pptx', ops: [{ op: 'addSlide' }, { op: 'addSlide' }], summary: 'x' });
  const commands = [await service.agents.next('s1'), await service.agents.next('s1')];
  for (const command of commands) service.agents.reply('s1', command.requestId, applied);
  await Promise.all([first, second]);
  assert.deepEqual(commands.map(command => command.auto === true), [true, false]);
  // A path that does not end at the written file itself (a symlinked folder) opens no window, and says so.
  const { symlink } = await import('node:fs/promises');
  await mkdir(join(workspace, 'real-dir'));
  await symlink(join(workspace, 'real-dir'), join(workspace, 'linked'));
  const viaLink = service.aiCreate('s1', { path: 'linked/f.pptx' });
  await client(() => read);
  assert.equal((await viaLink).generating, false);
  // A file swapped in at the created path does not inherit the window.
  await make('e.pptx');
  await writeFile(join(workspace, 'e2.pptx'), 'other');
  await rename(join(workspace, 'e2.pptx'), join(workspace, 'e.pptx'));
  assert.equal(await autoOf('e.pptx'), false);
});

test('addImage paths become image bytes on the Host; anything else is refused before the editor sees it', async t => {
  const { service, workspace, client } = await load(t);
  await writeFile(join(workspace, 'logo.png'), PNG);
  await writeFile(join(workspace, 'notes.txt'), 'hello');
  const editing = service.aiEdit('s1', { path: 'deck.pptx', summary: 'logo', ops: [
    { op: 'addImage', slideId: 's1', path: 'logo.png', imageBase64: 'Zm9yZ2Vk', imageType: 'image/gif', left: 0, top: 0, width: 10, height: 10 },
    { op: 'addSlide' },
  ] });
  const command = await client(() => applied);
  await editing;
  assert.deepEqual(command.ops[0], { op: 'addImage', slideId: 's1', path: 'logo.png', left: 0, top: 0, width: 10, height: 10, imageType: 'image/png', imageBase64: PNG.toString('base64') }, 'the model\'s own bytes and type are replaced by the file\'s');
  assert.deepEqual(command.ops[1], { op: 'addSlide' });
  await assert.rejects(service.aiEdit('s1', { path: 'deck.pptx', summary: 'x', ops: [{ op: 'addImage', slideId: 's1', path: 'notes.txt', left: 0, top: 0, width: 1, height: 1 }] }), { code: 'invalid-edit', message: /not a PNG or JPEG/ });
  await assert.rejects(service.aiEdit('s1', { path: 'deck.pptx', summary: 'x', ops: [{ op: 'addImage', slideId: 's1', path: 'missing.png', left: 0, top: 0, width: 1, height: 1 }] }), { code: 'invalid-edit', message: /cannot read missing\.png/ });
  await assert.rejects(service.aiEdit('s1', { path: 'deck.pptx', summary: 'x', ops: [{ op: 'addImage', slideId: 's1', imageBase64: 'AAAA', imageType: 'image/png', left: 0, top: 0, width: 1, height: 1 }] }), { code: 'invalid-edit', message: /needs the path/ });
  assert.equal(service.agents.queues.get('s1')?.length ?? 0, 0, 'refused edits never reach the Client');
});

test('a batch of images is sized before any file is read, and each read is capped', async t => {
  const { service, workspace, reads, client } = await load(t);
  const nine = Buffer.concat([PNG, Buffer.alloc(9 * 1024 * 1024)]);
  await writeFile(join(workspace, 'nine.png'), nine);
  await writeFile(join(workspace, 'big.png'), Buffer.concat([PNG, Buffer.alloc(PPTX_LIMITS.maxImageBytes)]));
  const image = path => ({ op: 'addImage', slideId: 's1', path, left: 0, top: 0, width: 1, height: 1 });
  reads.length = 0;
  await assert.rejects(service.aiEdit('s1', { path: 'deck.pptx', summary: 'x', ops: [image('nine.png'), image('nine.png'), image('nine.png')] }), { code: 'invalid-edit', message: /Operation 3: images in one edit add up to more than 20 MB/ });
  await assert.rejects(service.aiEdit('s1', { path: 'deck.pptx', summary: 'x', ops: [image('big.png')] }), { code: 'invalid-edit', message: /larger than 10 MB/ });
  await assert.rejects(service.aiEdit('s1', { path: 'deck.pptx', summary: 'x', ops: Array.from({ length: PPTX_LIMITS.maxOps + 1 }, () => image('nine.png')) }), { message: /At most 200 operations/ });
  // Too many new slides fails before any image is read (the image here is missing).
  await assert.rejects(service.aiEdit('s1', { path: 'deck.pptx', summary: 'x', ops: [...Array.from({ length: 5 }, () => ({ op: 'addSlide' })), image('missing.png')] }), { code: 'invalid-edit', message: /At most 4 new slides per edit/ });
  assert.deepEqual(reads, [], 'nothing was read for a batch over the limits');
  // An image above the 2 MB ranged-read window still loads.
  const editing = service.aiEdit('s1', { path: 'deck.pptx', summary: 'x', ops: [image('nine.png')] });
  const command = await client(() => ({ ok: true, result: { applied: 1, slideIds: ['s1'] } }));
  await editing;
  assert.equal(Buffer.from(command.ops[0].imageBase64, 'base64').length, nine.length);
});

test('creating outside the workspace asks through DSH approval; a refusal creates nothing', async t => {
  const { service, created, client } = await load(t);
  await assert.rejects(service.aiCreate('s1', { path: '/far/x.pptx' }, { approve: async () => 'denied' }), { code: 'sandbox-denied', message: /did not allow/ });
  const reasons = [];
  const creating = service.aiCreate('s1', { path: '/far/y.pptx' }, { approve: async reason => { reasons.push(reason); return 'allowed-once'; } });
  await client(() => read);
  assert.equal((await creating).created, true);
  assert.deepEqual(created.map(item => [item.path, item.approved]), [['/far/y.pptx', true]]);
  assert.match(reasons[0], /outside the workspace/);
});

test('presentation_create without a style uses the workspace\'s remembered one while its pack exists; a named style always wins', async t => {
  const packs = await realpath(await mkdtemp(join(tmpdir(), 'suite-generate-styles-')));
  t.after(() => rm(packs, { recursive: true, force: true }));
  await cp(STYLE_ROOT, packs, { recursive: true });
  await cp(join(packs, 'default'), join(packs, 'second'), { recursive: true });
  const style = JSON.parse(await readFile(join(packs, 'second', 'style.json'), 'utf8'));
  await writeFile(join(packs, 'second', 'style.json'), JSON.stringify({ ...style, name: 'second', default_palette: 'color-2' }));
  const { service, workspace, client } = await load(t, { styleRoot: packs });
  const make = async (path, style) => { const creating = service.aiCreate('s1', { path, style }); await client(() => read); return creating; };
  let made = await make('a.pptx');
  assert.deepEqual([made.style.name, made.styleSource], ['default', 'default']);
  await service.styleMemory.set(workspace, 'second');
  made = await make('b.pptx');
  assert.deepEqual([made.style.name, made.styleSource, made.style.palette.primary], ['second', 'workspace', style.palettes['color-2'].primary]);
  made = await make('c.pptx', 'default');
  assert.deepEqual([made.style.name, made.styleSource], ['default', 'given']);
  await rm(join(packs, 'second'), { recursive: true });
  made = await make('d.pptx');
  assert.deepEqual([made.style.name, made.styleSource], ['default', 'default'], 'a deleted pack falls back at once, not after the cache expires');
});

test('tools: presentation_create calls the Host with the agent Session and asks DSH for a path outside the workspace', async () => {
  const tools = new Map(), calls = [], asked = [];
  const service = { aiCreate: async (sessionId, args, { approve }) => { calls.push([sessionId, args]); return { outcome: await approve('create it') }; } };
  applyTools({ tools: { register: tool => tools.set(tool.name, tool) }, shimoPreview: service, get: name => name === 'approval' ? { request: async request => { asked.push(request); return 'allowed-once'; } } : undefined });
  const exec = { agent: { session: { id: 's9' } }, name: 'presentation_create', callId: 'c1', signal: new AbortController().signal };
  assert.deepEqual(await tools.get('presentation_create').execute({ path: '/far/a.pptx', slideCount: 6 }, exec), { outcome: 'allowed-once' });
  assert.deepEqual(calls, [['s9', { path: '/far/a.pptx', style: undefined, slideCount: 6 }]], 'no style is left for the Host to choose (the workspace default)');
  const create = tools.get('presentation_create');
  assert.doesNotMatch(create.parameters.properties.style.description, /default "default"/);
  assert.match(create.description, /quarterly\/work report \(slide-002\)/);
  assert.match(create.description, /leave it out otherwise/);
  assert.match(create.description, /styleSource/);
  assert.match(asked[0].displayReason, /\/far\/a\.pptx/);
  const edit = tools.get('presentation_edit');
  assert.ok(!edit.parameters.properties.operations.items.properties.op.enum.includes('addTable'), 'tables stay held back (K6)');
  assert.ok(edit.parameters.properties.operations.items.properties.op.enum.includes('addImage'));
  assert.match(edit.description, /at most 200 operations and 4 new slides/i);
  assert.match(edit.description, /sorter/);
  for (const tool of [create, edit]) assert.match(tool.description, /without narrating progress/);
  assert.match(edit.description, /whole or not at all/);
  assert.doesNotMatch(edit.description, /once per slide/);
});

function fakeController() {
  const calls = [];
  let finish;
  return {
    calls, savedVersion: 'v1', editRevision: 0, state: {},
    ensureMode: async () => {},
    agent: async (op, args, signal) => {
      calls.push(op);
      if (op === 'inspect') return { slides: [{ id: 's1', shapes: [] }] };
      return new Promise((resolve, reject) => { finish = resolve; signal?.addEventListener('abort', () => reject(Object.assign(new Error('The edit was cancelled; nothing was changed.'), { code: 'officesdk/cancelled' }))); });
    },
    finish: value => finish(value),
    propose: async () => { calls.push('propose'); return true; },
    set(patch) { this.state = { ...this.state, ...patch }; },
  };
}

test('client: a generated batch skips the card; the tab offers Stop, which cancels the batch and ends the window', async () => {
  const controller = fakeController();
  const replies = [], stops = [], timers = [];
  const registry = new ControllerRegistry();
  const scheduler = { setTimeout: (fn, ms) => { timers.push({ fn, ms }); return timers.length; }, clearTimeout: id => { if (timers[id - 1]) timers[id - 1].cleared = true; } };
  const rpc = { call: async (_, method, { args }) => { if (method.endsWith('/agentReply')) replies.push(args); if (method.endsWith('/agentStop')) stops.push(args); return { status: 'success', value: null }; } };
  const bridge = new AgentBridge({ rpc, openResource: address => registry.add(address, controller), controllers: registry, mounted: { getSnapshot: () => 's1', subscribe: () => () => {} }, scheduler, aliveMs: 1e9 });
  const ops = [{ op: 'addSlide', ref: 'p' }];
  const first = bridge.serve('s1', { requestId: 'r1', kind: 'edit', path: 'new.pptx', summary: 'cover', editId: 'e1', ops, auto: true });
  await new Promise(resolve => setTimeout(resolve, 0));
  assert.deepEqual(controller.calls, ['inspect', 'apply'], 'no card');
  assert.equal(controller.state.generating.running, true);
  controller.finish({ applied: 1 });
  await first;
  assert.equal(replies[0].ok, true);
  assert.equal(controller.state.generating.running, false, 'between batches the tab still offers Stop for a while');
  const idle = timers.findLast(timer => timer.ms === 60_000 && !timer.cleared);
  assert.ok(idle, 'and clears it after a quiet spell, long enough for the model to write a batch of several slides');
  // Stop during the next batch: the batch is cancelled and the Host is told.
  const second = bridge.serve('s1', { requestId: 'r2', kind: 'edit', path: 'new.pptx', summary: 'page 2', editId: 'e2', ops, auto: true });
  await new Promise(resolve => setTimeout(resolve, 0));
  controller.state.generating.stop();
  await second;
  assert.deepEqual([replies[1].ok, replies[1].code], [false, 'cancelled']);
  assert.deepEqual(stops, [{ workspaceFileScopeId: 's1', path: 'new.pptx' }]);
  assert.equal(controller.state.generating.stopped, true);
  // An `auto` batch the Host sent before it heard of the Stop still gets the card in this tab.
  controller.calls.length = 0;
  const late = bridge.serve('s1', { requestId: 'r4', kind: 'edit', path: 'new.pptx', summary: 'page 3', editId: 'e4', ops, auto: true });
  await new Promise(resolve => setTimeout(resolve, 0));
  controller.finish({ applied: 1 });
  await late;
  assert.deepEqual(controller.calls, ['inspect', 'propose', 'apply']);
  // Without `auto` the card is back.
  controller.calls.length = 0;
  const third = bridge.serve('s1', { requestId: 'r3', kind: 'edit', path: 'new.pptx', summary: 'page 2', editId: 'e3', ops });
  await new Promise(resolve => setTimeout(resolve, 0));
  controller.finish({ applied: 1 });
  await third;
  assert.deepEqual(controller.calls, ['inspect', 'propose', 'apply']);
});

test('client: a save ends an idle "generating · Stop" at once; a running batch keeps it', async () => {
  const controller = fakeController();
  const saved = [];
  controller.markSaved = (version, editRevision) => saved.push([version, editRevision]);
  const registry = new ControllerRegistry();
  const timers = [];
  const scheduler = { setTimeout: (fn, ms) => { timers.push({ fn, ms }); return timers.length; }, clearTimeout: id => { if (timers[id - 1]) timers[id - 1].cleared = true; } };
  const rpc = { call: async () => ({ status: 'success', value: null }) };
  const bridge = new AgentBridge({ rpc, openResource: address => registry.add(address, controller), controllers: registry, mounted: { getSnapshot: () => 's1', subscribe: () => () => {} }, scheduler, aliveMs: 1e9 });
  const ops = [{ op: 'addSlide', ref: 'p' }];
  const first = bridge.serve('s1', { requestId: 'r1', kind: 'edit', path: 'new.pptx', summary: 'cover', editId: 'e1', ops, auto: true });
  await new Promise(resolve => setTimeout(resolve, 0));
  // A save while the batch runs leaves the pill alone.
  await bridge.serve('s1', { kind: 'saved', path: 'new.pptx', version: 'v1', editRevision: 1 });
  assert.equal(controller.state.generating.running, true);
  controller.finish({ applied: 1 });
  await first;
  const idle = timers.findLast(timer => timer.ms === 60_000);
  assert.equal(controller.state.generating.running, false);
  await bridge.serve('s1', { kind: 'saved', path: 'new.pptx', version: 'v2', editRevision: 2 });
  assert.equal(controller.state.generating, undefined, 'gone with the save');
  assert.equal(idle.cleared, true);
  assert.deepEqual(saved, [['v1', 1], ['v2', 2]]);
});

test('tools: the workspace\'s remembered PPT style reaches the model as runtime context, only where DSH has systemPrompt', async () => {
  const entries = [], injected = [];
  const service = { styleContext: cwd => (cwd === '/ws/sub' ? 'Workspace default PPT style (shimo-pptx): second.' : cwd === '/boom' ? (() => { throw new Error('x'); })() : '') };
  const ctx = { tools: { register: () => {} }, shimoPreview: service, get: () => undefined,
    inject: (names, apply) => { injected.push(names); apply({ systemPrompt: { context: entry => entries.push(entry) } }); } };
  applyTools(ctx);
  assert.deepEqual(injected, [['systemPrompt']]);
  assert.equal(entries.length, 1);
  const [entry] = entries;
  assert.deepEqual([entry.name, entry.order, entry.interpolate], [STYLE_CONTEXT.name, 300, false]);
  const text = cwd => entry.text({ agent: { session: { header: { cwd } } } });
  assert.equal(text('/ws/sub'), 'Workspace default PPT style (shimo-pptx): second.');
  assert.equal(text('/elsewhere'), '', 'empty is left out by DSH');
  assert.equal(text('/boom'), '', 'never throws into the agent loop');
  assert.equal(entry.text({}), '', 'no agent');
  assert.equal(typeof entry.text({ agent: { session: { header: {} } } }), 'string');
  // A composition without inject (or without systemPrompt) still gets the tools.
  const tools = new Map();
  applyTools({ tools: { register: tool => tools.set(tool.name, tool) }, shimoPreview: service, get: () => undefined });
  for (const name of ['presentation_create','presentation_edit','presentation_save','document_create','document_finish_meeting_table','document_finish_project_table','document_finish_proposal_table','spreadsheet_create','selection_rewrite']) assert.ok(tools.has(name), name);
});

test('tools + Host: the context names the remembered style through a real Host, through cordis inject', async t => {
  const packs = await realpath(await mkdtemp(join(tmpdir(), 'suite-context-styles-')));
  t.after(() => rm(packs, { recursive: true, force: true }));
  await cp(STYLE_ROOT, packs, { recursive: true });
  await cp(join(packs, 'default'), join(packs, 'second'), { recursive: true });
  const style = JSON.parse(await readFile(join(packs, 'second', 'style.json'), 'utf8'));
  await writeFile(join(packs, 'second', 'style.json'), JSON.stringify({ ...style, name: 'second', meta: { id: 'second', name: { zh: '黑底画报', en: 'Dark editorial' } } }));
  const { service, workspace } = await load(t, { styleRoot: packs });
  const app = service.ctx.root ?? service.ctx;
  const entries = [];
  app.provide('systemPrompt'); app.set('systemPrompt', { context: entry => entries.push(entry) });
  app.provide('tools'); app.set('tools', { register: () => {} });
  await app.plugin({ name: 'tools-under-test', inject: ['tools', 'shimoPreview'], apply: applyTools });
  await new Promise(resolve => setTimeout(resolve, 20));
  assert.equal(entries.length, 1);
  assert.equal(entries[0].text({ agent: { session: { header: { cwd: workspace } } } }), '');
  await service.styleMemory.set(workspace, 'default');
  assert.equal(entries[0].text({ agent: { session: { header: { cwd: workspace } } } }), '', 'default needs no reminder');
  await service.styleMemory.set(workspace, 'second');
  await mkdir(join(workspace, 'sub'));
  assert.equal(entries[0].text({ agent: { session: { header: { cwd: join(workspace, 'sub') } } } }),
    'Workspace default PPT style (shimo-pptx): second (黑底画报 / Dark editorial). When the user names no style, use it: leave style out of presentation_create, and pass --style second to brand_deck.py.');
});

test('reporting create validates chart inputs before creating, and returns eight live pages with target slots', async t => {
  const { service, client, created } = await load(t);
  const args = { path: 'quarter.pptx', style: 'slide-002', slideCount: 8 };
  await assert.rejects(service.aiCreate('s1', { ...args, reportingData: { chart: { values: [1] } } }), { code: 'invalid-edit' });
  assert.equal(created.length, 0);
  const creating = service.aiCreate('s1', { ...args, reportingData: { chart: { categories: ['甲', '乙', '丙'], values: [2, 3, 5] }, bars: [5, 4, 3, 2, 1] } });
  const command = await client(() => read);
  assert.equal(command.initialBlank, 'ppt-template');
  const result = await creating;
  assert.equal(result.style.textSlots.length, 8);
  assert.deepEqual(result, JSON.parse(JSON.stringify(result)));
  const editPage = async (id, completedSlideIds, expected, appliedCount = 1) => {
    const writing = service.aiEdit('s1', { path: args.path, ops: [{ op: 'setText', slideId: id, shapeId: 'title', text: '正文' }], completedSlideIds });
    const cmd = await client(() => ({ ok: true, result: { applied: appliedCount, slideIds: [id] } }));
    assert.deepEqual(cmd.slideProgress, expected);
    await writing;
  };
  await editPage('s1', [], { done: [], partial: [] });
  await editPage('s2', ['s2'], { done: [], partial: ['s1'] }, 0);
  await editPage('s2', ['s2'], { done: [], partial: ['s1'] });
  await editPage('s1', ['s1'], { done: ['s2'], partial: ['s1'] });
});
