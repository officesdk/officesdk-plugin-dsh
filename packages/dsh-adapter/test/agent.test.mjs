import test from 'node:test';
import assert from 'node:assert/strict';
import { mkdtemp, mkdir, rm, writeFile } from 'node:fs/promises';
import { join } from 'node:path';
import { tmpdir } from 'node:os';
import { Context } from '@deepseek-ai/cordis';
import ShimoPreview from '../src/host/index.mjs';
import { AgentHub } from '../src/host/agent.mjs';
import { AgentBridge, ControllerRegistry, deckOpener, sessionAddress } from '../src/client/agent.mjs';
import { callKey, createToolViewStore, editKey } from '../src/client/tool-views.mjs';
import { apply as applyTools, inject as toolInject } from '../src/tools/index.mjs';
import { fakeConverter } from '../../core/test/helpers.mjs';
import { SuiteError } from '../../core/src/index.mjs';

const tick = () => new Promise(resolve => setTimeout(resolve, 0));
const fakeTimers = () => {
  const timers = [];
  return { timers, setTimeout: (fn, ms) => { timers.push({ fn, ms }); return timers.length; }, clearTimeout: id => { if (timers[id - 1]) timers[id - 1].cleared = true; }, fire: ms => { for (const timer of timers) if (timer.ms === ms && !timer.cleared) { timer.cleared = true; timer.fn(); } } };
};

test('hub: a command reaches the polling Client and its reply settles the tool call', async () => {
  const hub = new AgentHub();
  const pending = hub.request('s1', { kind: 'read', path: 'a.pptx' });
  const command = await hub.next('s1');
  assert.deepEqual({ ...command, requestId: undefined }, { kind: 'read', path: 'a.pptx', requestId: undefined });
  assert.equal(hub.reply('s2', command.requestId, { ok: true, result: 1 }), false, 'another Session cannot answer');
  assert.equal(hub.reply('s1', command.requestId, { ok: true, result: { slides: 2 } }), true);
  assert.deepEqual(await pending, { slides: 2 });
  const failing = hub.request('s1', { kind: 'edit' });
  const next = await hub.next('s1');
  hub.reply('s1', next.requestId, { ok: false, code: 'rm -rf', message: 'nope' });
  await assert.rejects(failing, { code: 'editor-failed', message: 'nope' });
});

test('hub: every command carries the Host\'s platform when it has one, so the Client spells the path the Host\'s way', async () => {
  const hub = new AgentHub({ platform: 'win32' });
  void hub.request('s1', { kind: 'read', path: 'sub\\a.pptx' }).catch(() => {});
  const command = await hub.next('s1');
  assert.deepEqual([command.kind, command.path, command.platform], ['read', 'sub\\a.pptx', 'win32']);
  hub.reply('s1', command.requestId, { ok: true, result: {} });
});

test('hub: nobody showing the Session times out; cancellation withdraws or forwards', async () => {
  const timers = fakeTimers();
  const hub = new AgentHub({ scheduler: timers, pickupMs: 5, pollMs: 7 });
  const orphan = hub.request('s1', { kind: 'read' });
  timers.fire(5);
  await assert.rejects(orphan, { code: 'no-editor' });
  const idle = hub.next('s1');
  timers.fire(7);
  assert.equal(await idle, null, 'an idle poll ends empty');
  const early = new AbortController();
  const withdrawn = hub.request('s1', { kind: 'read' }, early.signal);
  early.abort();
  await assert.rejects(withdrawn, { code: 'cancelled' });
  assert.equal(hub.queues.get('s1').length, 0, 'an uncollected command is withdrawn');
  const late = new AbortController();
  const running = hub.request('s1', { kind: 'edit' }, late.signal);
  const picked = await hub.next('s1');
  timers.fire(5);
  late.abort();
  await assert.rejects(running, { code: 'cancelled' });
  assert.deepEqual(await hub.next('s1'), { kind: 'cancel', target: picked.requestId }, 'a collected command is cancelled in the Client');
});

async function load(t) {
  const root = await mkdtemp(join(tmpdir(), 'suite-agent-'));
  t.after(() => rm(root, { recursive: true, force: true }));
  await mkdir(join(root, 'runtime')); await writeFile(join(root, 'runtime', 'preview.html'), '');
  const file = { data: Buffer.from('before'), version: 'v1' };
  const app = new Context();
  const services = {
    workspaceFiles: {
      stat: async (scope, path) => ({ absolutePath: `/w/${path}`, version: file.version, bytes: file.data.length }),
      readBytes: async (scope, path) => ({ absolutePath: `/w/${path}`, version: file.version, data: file.data, eof: true }),
    },
    webServer: { register: () => () => {} },
    connection: { admit: () => ({ rejection: 403 }) },
    typert: { lookups: { get: () => ({ resolve: async id => ({ sessionId: id, workspaceRoot: '/w' }) }) } },
  };
  for (const [name, value] of Object.entries(services)) { app.provide(name); app.set(name, value); }
  const writes = [], outside = new Set(), denied = new Set();
  const writer = {
    canWrite: () => true,
    writeAccess: async (scope, path) => denied.has(path) ? 'denied' : outside.has(path) ? 'outside' : 'allowed',
    replace: async (scope, path, data, { expectedVersion, approved }) => {
      if (outside.has(path) && !approved) throw new SuiteError('sandbox-denied', 'outside');
      if (expectedVersion !== file.version) throw new SuiteError('source-changed', 'changed');
      writes.push({ path, approved: approved === true });
      file.data = Buffer.from(data); file.version = `v${Number(file.version.slice(1)) + 1}`;
      return { version: file.version, absolutePath: `/w/${path}` };
    },
  };
  const environment = { converter: fakeConverter(), runtimeRoot: join(root, 'runtime'), tempRoot: join(root, 'tmp'), dataRoot: join(root, 'data'), writer };
  class Plugin extends ShimoPreview { constructor(ctx, config) { super(ctx, config, environment); } }
  Plugin.inject = ShimoPreview.inject;
  await app.plugin(Plugin, {});
  await new Promise(resolve => setTimeout(resolve, 20));
  const service = app.get('shimoPreview');
  const peer = { id: 'tab', principal: 'owner', signal: new AbortController().signal, remote: { principal: { kind: 'owner' } } };
  const scope = { sessionId: 's1', workspaceRoot: '/w' };
  const signal = () => new AbortController().signal;
  /** The tab: an edit session whose editor staged `body`. */
  const openEdited = async (path, body) => {
    const lease = await service.suite.open(scope, path, 1, signal(), peer, { mode: 'edit' });
    if (body) await service.suite.stage(lease.id, 'snapshot', Buffer.from(body), { baseEditRevision: 0 }, peer, signal());
    return lease;
  };
  /** Plays the Client: collects the next command and answers it. */
  const client = async answer => { const command = await service.agents.next('s1'); service.agents.reply('s1', command.requestId, await answer(command)); return command; };
  return { service, file, writes, outside, denied, root, client, openEdited };
}

test('host: an AI edit stays in the editor; nothing is written until a save', async t => {
  const { service, writes, client } = await load(t);
  const editing = service.aiEdit('s1', { path: 'deck.pptx', ops: [{ op: 'addSlide' }], summary: 'add a slide' });
  const command = await client(() => ({ ok: true, result: { editRevision: 1, applied: 1, slideIds: ['n'], undoable: 'x' } }));
  assert.equal(command.kind, 'edit');
  assert.match(command.editId, /^[0-9a-f-]{36}$/);
  assert.deepEqual(await editing, { editRevision: 1, applied: 1, slideIds: ['n'], undoable: 'x', editId: command.editId, unsaved: true });
  assert.deepEqual(writes, []);
  const undoing = service.aiUndo('s1', { path: 'deck.pptx', editId: command.editId });
  const undo = await client(() => ({ ok: true, result: { editRevision: 2, undone: command.editId, undoable: null } }));
  assert.deepEqual([undo.kind, undo.editId, undo.path], ['undo', command.editId, 'deck.pptx']);
  assert.equal((await undoing).undone, command.editId);
  const declined = service.aiEdit('s1', { path: 'deck.pptx', ops: [], summary: '' });
  await client(() => ({ ok: true, result: { declined: true } }));
  assert.deepEqual(await declined, { declined: true });
  const failed = service.aiEdit('s1', { path: 'deck.pptx', ops: [], summary: '' });
  await client(() => ({ ok: false, code: 'invalid-edit', message: 'Operation 1: bad' }));
  await assert.rejects(failed, { code: 'invalid-edit' });
  assert.equal('callId' in command, false, 'no call id, no field');
});

test('host: the tool call id reaches the Client with read and edit commands, and nothing else changes', async t => {
  const { service, client } = await load(t);
  const reading = service.aiRead('s1', 'deck.pptx', undefined, { callId: 'call-7' });
  const read = await client(() => ({ ok: true, result: { deck: {} } }));
  assert.deepEqual({ ...read, requestId: undefined }, { kind: 'read', path: 'deck.pptx', callId: 'call-7', requestId: undefined, platform: process.platform }, 'the Host names its own platform');
  assert.deepEqual(await reading, { deck: {} });
  const editing = service.aiEdit('s1', { path: 'deck.pptx', ops: [{ op: 'addSlide' }], summary: 's', callId: 'call-8' });
  const edit = await client(() => ({ ok: true, result: { applied: 1, slideIds: ['n'] } }));
  assert.equal(edit.callId, 'call-8');
  assert.deepEqual(await editing, { applied: 1, slideIds: ['n'], editId: edit.editId, unsaved: true });
});

test('host: presentation_save writes the editor\'s changes and tells the tab; outside the workspace it asks first', async t => {
  const { service, writes, outside, openEdited } = await load(t);
  await assert.rejects(service.aiSave('s1', 'deck.pptx', {}), { code: 'not-found' });
  await openEdited('deck.pptx');
  assert.deepEqual(await service.aiSave('s1', 'deck.pptx', {}), { saved: false, unchanged: true, version: 'v1', path: '/w/deck.pptx' });
  const deck = await openEdited('edited.pptx', '{"slides":[2]}');
  const saved = await service.aiSave('s1', 'edited.pptx', {});
  assert.deepEqual([saved.saved, saved.version, saved.path], [true, 'v2', '/w/edited.pptx']);
  assert.deepEqual(writes, [{ path: 'edited.pptx', approved: false }]);
  const notice = await service.agents.next('s1');
  assert.deepEqual([notice.kind, notice.path, notice.version, notice.editRevision], ['saved', 'edited.pptx', 'v2', 1]);
  assert.equal(service.suite.dirty(service.suite.sessions.get(deck.id)), false);

  outside.add('far.pptx');
  await openEdited('far.pptx', '{"slides":[3]}');
  const asked = [];
  await assert.rejects(service.aiSave('s1', 'far.pptx', { approve: async reason => { asked.push(reason); return 'rejected'; } }),
    error => error.code === 'sandbox-denied' && /did not allow/.test(error.message));
  await assert.rejects(service.aiSave('s1', 'far.pptx', {}), /no approval channel/, 'without an approval service it fails closed');
  assert.equal(writes.length, 1);
  assert.match(asked[0], /\/w\/far\.pptx/);
  const approved = await service.aiSave('s1', 'far.pptx', { approve: async () => 'allowed-once' });
  assert.equal(approved.saved, true);
  assert.deepEqual(writes.at(-1), { path: 'far.pptx', approved: true });
});

test('host: a read-only Session is not asked about "outside the workspace"; it cannot save at all', async t => {
  const { service, writes, denied, openEdited } = await load(t);
  denied.add('ro.pptx');
  await openEdited('ro.pptx', '{"slides":[4]}');
  const asked = [];
  await assert.rejects(service.aiSave('s1', 'ro.pptx', { approve: async reason => { asked.push(reason); return 'allowed-once'; } }),
    error => error.code === 'sandbox-denied' && /read only/.test(error.message) && !error.details?.outside);
  assert.deepEqual([asked, writes], [[], []]);
});

test('host: only the owner may approve a panel save outside the workspace', async t => {
  const { service, writes, outside, openEdited } = await load(t);
  outside.add('far.pptx');
  const deck = await openEdited('far.pptx', '{"slides":[5]}');
  service.remotePeer = () => ({ id: 'tab', principal: 'guest', signal: new AbortController().signal, remote: { principal: { kind: 'remote' } } });
  await assert.rejects(service.save({ sessionId: 's1', workspaceRoot: '/w' }, deck.id, 1, true, new AbortController().signal), error => /forbidden/.test(error.code) && /Only the owner/.test(error.message));
  assert.deepEqual(writes, []);
});

function fakeController({ approve = true } = {}) {
  const calls = [];
  return {
    calls, savedVersion: 'v1', editRevision: 0, state: {},
    ensureMode: async mode => calls.push(['mode', mode]),
    agent: async (op, args) => {
      calls.push([op, args]);
      if (op === 'inspect') return { slides: [{ id: 's1', shapes: [{ id: 't', name: 'T', left: 0, top: 0, width: 1, height: 1, text: 'old' }] }] };
      if (op === 'undo') return { editRevision: 2, undone: args.editId, undoable: null };
      return { editRevision: 1, applied: 1, slideIds: ['s1'], undoable: args.editId };
    },
    propose: async proposal => { calls.push(['propose', proposal]); return approve; },
    markSaved(version, editRevision) { calls.push(['saved', version, editRevision]); },
    set(patch) { this.state = { ...this.state, ...patch }; },
  };
}

function bridgeFor(controller, views) {
  const replies = [], opened = [];
  const registry = new ControllerRegistry();
  const rpc = { call: async (_, method, { args }) => { if (method.endsWith('/agentReply')) replies.push(args); return { status: 'success', value: null }; } };
  const bridge = new AgentBridge({ rpc, openResource: address => { opened.push(address); registry.add(address, controller); }, controllers: registry, mounted: { getSnapshot: () => 's1', subscribe: () => () => {} }, views });
  return { bridge, replies, opened, registry };
}

test('client: the bridge records each call for its chat row: read outline, pages before approval, actual pages after', async () => {
  const deckOf = ids => ({ slides: ids.map(id => ({ id, shapes: [{ id: 't', name: `T-${id}`, left: 0, top: 0, width: 1, height: 1, text: `text ${id}` }] })) });
  let slides = ['a', 'b'], made = 0, release;
  const gate = new Promise(resolve => { release = resolve; });
  const calls = [];
  const controller = {
    calls, savedVersion: 'v1', editRevision: 0, state: { dirty: true },
    ensureMode: async () => {},
    agent: async (op, args) => {
      calls.push(op);
      if (op === 'inspect') return deckOf(slides);
      // Each batch appends one slide (no index), like generation does.
      const id = `n${++made}`;
      slides = [...slides, id];
      return { editRevision: made, applied: args.ops.length, slideIds: [id], refs: { p: id }, undoable: args.editId };
    },
    propose: async () => { await gate; return true; },
    set(patch) { this.state = { ...this.state, ...patch }; },
  };
  const views = createToolViewStore();
  const { bridge } = bridgeFor(controller, views);
  await bridge.serve('s1', { requestId: 'r0', kind: 'read', path: 'a.pptx', callId: 'c0' });
  assert.deepEqual(views.get(callKey('c0')), { outline: [{ shapes: 1, text: 'text a' }, { shapes: 1, text: 'text b' }], unsaved: true });
  const ops = [{ op: 'addSlide', ref: 'p' }, { op: 'addTextBox', slideRef: 'p', text: 'x', left: 0, top: 0, width: 10, height: 10 }, { op: 'setText', slideId: 'b', shapeId: 't', text: 'y' }];
  const first = bridge.serve('s1', { requestId: 'r1', kind: 'edit', path: 'a.pptx', summary: 'one', editId: 'e1', callId: 'c1', ops: [...ops, { op: 'addImage', slideId: 'b', imageBase64: 'QUJD', imageType: 'image/png', path: 'i.png', left: 0, top: 0, width: 1, height: 1 }] });
  const second = bridge.serve('s1', { requestId: 'r2', kind: 'edit', path: 'a.pptx', summary: 'two', editId: 'e2', callId: 'c2', ops });
  await tick(); await tick();
  const waiting = views.get(callKey('c1'));
  assert.equal(waiting.awaiting, true, 'waiting for the user');
  assert.deepEqual([waiting.slides, waiting.shapes, waiting.summary], [{ b: 2 }, { 'b/t': 'T-b' }, 'one']);
  assert.ok(waiting.ops.every(op => !('imageBase64' in op)), 'image bytes stay out of the record');
  assert.equal(views.get(editKey('e1')), waiting, 'an undo finds it by edit id');
  release();
  await Promise.all([first, second]);
  for (let i = 0; i < 5; i++) await tick();
  const one = views.get(callKey('c1')), two = views.get(callKey('c2'));
  assert.equal(one.awaiting, false);
  // Settled after both batches: each batch's new slide at its real page.
  assert.deepEqual([one.touched, one.refs, one.slides.b], [[3], { p: 3 }, 2]);
  assert.deepEqual([two.touched, two.refs], [[4], { p: 4 }]);
  // Without a store nothing is recorded and nothing breaks.
  const plain = bridgeFor(fakeController());
  await plain.bridge.serve('s1', { requestId: 'r9', kind: 'edit', path: 'a.pptx', callId: 'c9', editId: 'e9', ops: [{ op: 'setText', slideId: 's1', shapeId: 't', text: 'z' }] });
  assert.equal(plain.replies[0].ok, true);
});

test('client: a chat row opens the very address the AI tab uses, only for the Session the Sidebar shows', async () => {
  const opened = [], listeners = [];
  let mounted = 's1';
  const sidebar = { openResource: address => opened.push(address), mounted: { getSnapshot: () => mounted, subscribe: listener => { listeners.push(listener); return () => {}; } } };
  const deck = deckOpener(() => sidebar);
  for (const path of ['deck.pptx', './deck.pptx', '/w/decks/a b.pptx']) {
    deck.openDeck('s1', path);
    // The bridge opens the tab for the model's own path the same way.
    const controller = fakeController();
    const { bridge, opened: bridged } = bridgeFor(controller);
    await bridge.serve('s1', { requestId: 'r', kind: 'read', path });
    assert.equal(opened.at(-1), bridged[0], path);
    assert.equal(opened.at(-1), sessionAddress('s1', path));
  }
  assert.equal(deck.canOpenDeck('s1'), true);
  assert.equal(deck.canOpenDeck('s2'), false, 'another Session: no button');
  const off = deck.subscribeDeck(() => {});
  assert.equal(listeners.length, 1);
  assert.equal(typeof off, 'function');
  mounted = undefined;
  assert.equal(deck.canOpenDeck('s1'), false);
  const none = deckOpener(() => undefined);
  assert.equal(none.canOpenDeck('s1'), false, 'no Sidebar: no button');
  assert.equal(typeof none.subscribeDeck(() => {}), 'function');
  assert.doesNotThrow(() => none.openDeck('s1', 'a.pptx'));
});

test('client: a declined or failed approval clears the waiting mark', async () => {
  const views = createToolViewStore();
  const declined = fakeController({ approve: false });
  await bridgeFor(declined, views).bridge.serve('s1', { requestId: 'r1', kind: 'edit', path: 'a.pptx', summary: 's', editId: 'e1', callId: 'c1', ops: [{ op: 'setText', slideId: 's1', shapeId: 't', text: 'x' }] });
  assert.equal(views.get(callKey('c1')).awaiting, false);
  const failing = fakeController();
  failing.propose = async () => { throw Object.assign(new Error('gone'), { code: 'officesdk/cancelled' }); };
  const { bridge, replies } = bridgeFor(failing, views);
  await bridge.serve('s1', { requestId: 'r2', kind: 'edit', path: 'a.pptx', summary: 's', editId: 'e2', callId: 'c2', ops: [{ op: 'setText', slideId: 's1', shapeId: 't', text: 'x' }] });
  assert.deepEqual([replies[0].ok, replies[0].code, views.get(callKey('c2')).awaiting], [false, 'cancelled', false]);
  assert.equal(views.get(callKey('c2')).failure, 'cancelled', 'the row learns the code the tool text may lose');
  assert.equal(views.get(callKey('c2')).summary, 's', 'and keeps what was recorded before');
});

test('client: an edit opens the tab in edit mode, asks the user, applies, and remembers it for undo', async () => {
  const controller = fakeController();
  const { bridge, replies, opened } = bridgeFor(controller);
  await bridge.serve('s1', { requestId: 'r1', kind: 'edit', path: 'dir/My Deck.pptx', summary: 'rename', editId: 'e1', ops: [{ op: 'setText', slideId: 's1', shapeId: 't', text: 'new' }] });
  assert.deepEqual(opened, ['dsh-resource://file/session/s1/dir/My%20Deck.pptx']);
  assert.deepEqual(controller.calls.map(([name]) => name), ['mode', 'inspect', 'propose', 'apply']);
  assert.deepEqual(controller.calls[2][1].rows, [{ op: 'setText', where: '#1 · T', before: 'old', after: 'new' }]);
  assert.equal(controller.calls[3][1].editId, 'e1', 'the frame keys its undo stack by editId');
  assert.deepEqual(replies, [{ workspaceFileScopeId: 's1', requestId: 'r1', ok: true, result: { editRevision: 1, applied: 1, slideIds: ['s1'], undoable: 'e1' } }]);
  assert.deepEqual(controller.state.edits, [{ editId: 'e1', summary: 'rename' }]);
  await bridge.serve('s1', { requestId: 'r2', kind: 'undo', path: 'dir/My Deck.pptx', editId: 'e1' });
  assert.deepEqual(replies[1].result, { editRevision: 2, undone: 'e1', undoable: null, unsaved: true });
  assert.deepEqual(controller.state.edits, []);
  controller.state.dirty = true;
  await bridge.serve('s1', { requestId: 'r3', kind: 'read', path: 'dir/My Deck.pptx' });
  assert.deepEqual([replies[2].result.unsaved, replies[2].result.editRevision], [true, 0]);
});

test('client: the Host\'s save notice reaches the tab without a reply', async () => {
  const controller = fakeController();
  const { bridge, replies, registry } = bridgeFor(controller);
  registry.add('dsh-resource://file/session/s1/a.pptx', controller);
  await bridge.serve('s1', { requestId: 'n1', kind: 'saved', path: 'a.pptx', version: 'v4', editRevision: 2 });
  assert.deepEqual(controller.calls, [['saved', 'v4', 2]]);
  assert.deepEqual(replies, []);
});

test('client: invalid plans and declines never reach apply; other formats are refused', async () => {
  const declined = fakeController({ approve: false });
  let { bridge, replies } = bridgeFor(declined);
  await bridge.serve('s1', { requestId: 'r1', kind: 'edit', path: 'a.pptx', ops: [{ op: 'setText', slideId: 's1', shapeId: 't', text: 'x' }] });
  assert.deepEqual(replies[0].result, { declined: true });
  assert.ok(!declined.calls.some(([name]) => name === 'apply'));
  const invalid = fakeController();
  ({ bridge, replies } = bridgeFor(invalid));
  await bridge.serve('s1', { requestId: 'r2', kind: 'edit', path: 'a.pptx', ops: [{ op: 'setText', slideId: 'gone', shapeId: 't', text: 'x' }] });
  assert.deepEqual([replies[0].ok, replies[0].code], [false, 'invalid-edit']);
  assert.ok(!invalid.calls.some(([name]) => name === 'propose'));
  await bridge.serve('s1', { requestId: 'r3', kind: 'read', path: 'a.csv' });
  assert.deepEqual([replies[1].ok, replies[1].code], [false, 'unsupported-operation']);
  assert.equal(sessionAddress('s 1', '/abs/x.pptx'), 'dsh-resource://file/session/s%201//abs/x.pptx');
});

test('client: a tab that never mounts ends the command as no-editor', async () => {
  const timers = fakeTimers();
  const registry = new ControllerRegistry({ scheduler: timers, timeoutMs: 9 });
  const waiting = registry.wait('dsh-resource://file/session/s1/a.pptx');
  timers.fire(9);
  await assert.rejects(waiting, { code: 'officesdk/no-editor' });
});

test('tools: model-facing tools call the service with the agent Session; save asks through DSH approval', async () => {
  assert.deepEqual(toolInject, ['tools', 'shimoPreview']);
  const tools = new Map(), calls = [], asked = [];
  const service = {
    aiRead: async (...args) => { calls.push(['read', ...args.slice(0, 2), args[3]]); return { deck: {} }; },
    aiEdit: async (...args) => { calls.push(['edit', ...args.slice(0, 2)]); throw new SuiteError('invalid-edit', 'Operation 1: bad'); },
    aiUndo: async (...args) => { calls.push(['undo', ...args.slice(0, 2)]); return { undone: 'e' }; },
    aiSave: async (sessionId, path, { approve }) => { calls.push(['save', sessionId, path]); return { outcome: await approve('write it') }; },
  };
  const approval = { request: async request => { asked.push(request); return 'allowed-once'; } };
  applyTools({ tools: { register: tool => tools.set(tool.name, tool) }, shimoPreview: service, get: name => name === 'approval' ? approval : undefined });
  assert.deepEqual([...tools.keys()], ['presentation_create', 'presentation_suggest_styles', 'presentation_read', 'presentation_edit', 'presentation_undo', 'presentation_save', 'document_create', 'document_read', 'document_edit', 'document_save', 'spreadsheet_create', 'spreadsheet_read', 'spreadsheet_edit', 'spreadsheet_save', 'selection_rewrite']);
  const exec = { agent: { session: { id: 's9' } }, name: 'presentation_save', callId: 'c1', signal: new AbortController().signal };
  assert.deepEqual(await tools.get('presentation_read').execute({ path: 'a.pptx' }, exec), { deck: {} });
  await assert.rejects(tools.get('presentation_edit').execute({ path: 'a.pptx', summary: 'x', operations: [] }, exec), /invalid-edit: Operation 1: bad/);
  assert.deepEqual(await tools.get('presentation_undo').execute({ path: 'a.pptx', editId: 'e' }, exec), { undone: 'e' });
  assert.deepEqual(await tools.get('presentation_save').execute({ path: '/far/a.pptx' }, exec), { outcome: 'allowed-once' });
  // The call id rides along so the Client can label the chat row; the model-facing definitions do not change.
  assert.deepEqual(calls, [['read', 's9', 'a.pptx', { callId: 'c1' }], ['edit', 's9', { path: 'a.pptx', ops: [], summary: 'x', callId: 'c1' }], ['undo', 's9', { path: 'a.pptx', editId: 'e' }], ['save', 's9', '/far/a.pptx']]);
  assert.deepEqual([asked[0].agent, asked[0].toolName, asked[0].callId, asked[0].reason], [exec.agent, 'presentation_save', 'c1', 'write it']);
  assert.match(asked[0].displayReason, /\/far\/a\.pptx/);
  await assert.rejects(tools.get('presentation_read').execute({ path: 'a.pptx' }, { signal: exec.signal }), /agent Session/);
});

test('tools: a SuiteError from the Host\'s own copy of Core still reaches the model as "code: message"', async () => {
  // The release bundles Core separately into lib/host and lib/tools, so the service throws another class.
  const host = await import('../../core/src/errors.mjs?host-copy');
  const tools = new Map();
  const service = {
    aiRead: async () => { throw new host.SuiteError('busy', 'Another edit is running.'); },
    aiUndo: async () => { throw new Error('plain failure'); },
  };
  applyTools({ tools: { register: tool => tools.set(tool.name, tool) }, shimoPreview: service });
  const exec = { agent: { session: { id: 's1' } }, signal: new AbortController().signal };
  await assert.rejects(tools.get('presentation_read').execute({ path: 'a.pptx' }, exec), error => {
    assert.equal(error.message, 'busy: Another edit is running.');
    return true;
  });
  await assert.rejects(tools.get('presentation_undo').execute({ path: 'a.pptx', editId: 'e' }, exec), error => {
    assert.equal(error.message, 'plain failure', 'other errors pass through without a prefix');
    return true;
  });
});

test('hub: a picked command fails once its Client stops reporting, not before', async () => {
  const timers = fakeTimers();
  const hub = new AgentHub({ scheduler: timers, pickupMs: 5, aliveMs: 9 });
  const running = hub.request('s1', { kind: 'edit' });
  const picked = await hub.next('s1');
  timers.fire(5);
  assert.equal(hub.alive('s2', picked.requestId), false, 'another Session cannot keep it alive');
  assert.equal(hub.alive('s1', picked.requestId), true);
  timers.fire(9);
  await assert.rejects(running, { code: 'no-editor', message: /stopped responding/ });
  assert.equal(hub.alive('s1', picked.requestId), false, 'a settled command is gone');
});

test('bridge: reports while a command is pending and stops once it is answered', async () => {
  const timers = fakeTimers();
  const calls = [];
  let decide;
  const rpc = { call: async (_, method, { args }) => { calls.push([method.split('/')[1], args.requestId]); return { status: 'success', value: null }; } };
  const bridge = new AgentBridge({ rpc, openResource: () => {}, controllers: null, mounted: null, scheduler: timers, aliveMs: 3 });
  bridge.run = () => new Promise(resolve => { decide = resolve; });
  // Each beat schedules the next, so fire one pending timer at a time.
  const beat = () => { const timer = timers.timers.find(t => t.ms === 3 && !t.cleared); if (timer) { timer.cleared = true; timer.fn(); } };
  const serving = bridge.serve('s1', { kind: 'read', requestId: 'r1' });
  beat(); beat();
  decide({ ok: 1 });
  await serving;
  beat();
  assert.deepEqual(calls.map(([name]) => name), ['agentAlive', 'agentAlive', 'agentReply']);
});

test('client: on a Windows Host a chat row opens the address the AI tab uses for a backslash path', async () => {
  const opened = [];
  const sidebar = { openResource: address => opened.push(address), mounted: { getSnapshot: () => 's1', subscribe: () => () => {} } };
  const deck = deckOpener(() => sidebar, () => 'win32');
  deck.openDeck('s1', 'C:\\w\\decks\\a.pptx');
  assert.equal(opened[0], sessionAddress('s1', 'C:\\w\\decks\\a.pptx', 'win32'));
  assert.notEqual(opened[0], sessionAddress('s1', 'C:\\w\\decks\\a.pptx'), 'the Host platform changes the address');
});

test('client: the bridge remembers the platform the last Host command named, for chat rows opened with no tab left', async () => {
  const controller = fakeController();
  const { bridge } = bridgeFor(controller);
  assert.equal(bridge.lastPlatform, undefined);
  await bridge.serve('s1', { requestId: 'r', kind: 'read', path: 'C:\\w\\a.pptx', platform: 'win32' });
  assert.equal(bridge.lastPlatform, 'win32');
});
