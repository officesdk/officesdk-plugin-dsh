import test from 'node:test';
import assert from 'node:assert/strict';
import { DocumentController } from '../src/client/controller.mjs';
import { failureKind, fileAddress, parseSource, unwrapRemote } from '../src/client/address.mjs';
import { baseName, isAbsolute, joinPath, parentOf, rootOf, segments } from '../src/client/paths.mjs';
import { AgentBridge, ControllerRegistry, sessionAddress } from '../src/client/agent.mjs';
import { MODULES, UPDATE_CHECK_MS, WITHDRAWN_MODULES, anyUnsaved, cancelKey, copyName, createToolbarPrefs, createUpdateWatch, editButtons, menuKey, nameProblem, pickedPath, savedForSwitch, switchPlan, modeButtons, offeredModules, placeMenu, rovingKey, toolbarBusy, updateEntry, updatePanelView, watchTabEdits } from '../src/client/toolbar.mjs';
import { CSS, mountStyles } from '../src/client/styles.mjs';
import { DEMO_STATES, MOCK_KEY, PREVIEW_KEY, demoState, accountGate, accountHeld, accountMenuKey, createAccountSource, heldButtons, parseAccount } from '../src/client/account.mjs';

const base = 'http://app.local/mount/';
const address = 'dsh-resource://file/session/s1/docs/a.docx';
const tick = () => new Promise(resolve => setImmediate(resolve));

const failure = code => ({ status: 'failure', error: { code: `officesdk/${code}`, message: code } });

/** `host` overrides answers per method, e.g. { rebind: args => failure('source-changed') }. */
function harness({ open, host = {}, locale } = {}) {
  const calls = [];
  let leases = 0;
  const rpc = { call: async (_, method, { args }, signal) => {
    signal?.throwIfAborted();
    calls.push([method.split('/')[1], args]);
    const name = method.split('/')[1];
    if (host[name]) return host[name](args);
    if (name === 'save') return { status: 'success', value: { version: 'v2', editRevision: args.editRevision } };
    if (name === 'rebind') return { status: 'success', value: { revision: args.revision, state: args.diverged ? 'diverged' : 'active', version: 'v1' } };
    if (name === 'open') return { status: 'success', value: open?.(args) ?? { id: `lease${++leases}`, format: 'docx', revision: args.revision, version: 'v1', editRevision: 0, mode: args.mode,
      entry: `/shimo/doc/lease${leases}/index.html`, expiresInMs: 900_000, capabilities: { edit: true, save: true, export: true } } };
    if (name === 'exportTo' || name === 'exportDraft') return { status: 'success', value: { url: '/shimo/doc/x/export/t', name: 'a.docx' } };
    if (name === 'listDrafts') return { status: 'success', value: [] };
    return { status: 'success', value: { expiresInMs: 900_000 } };
  } };
  const listeners = new Set();
  const events = { addEventListener: (_, fn) => listeners.add(fn), removeEventListener: (_, fn) => listeners.delete(fn) };
  const posted = [];
  const frame = { contentWindow: { postMessage: (message, origin) => posted.push([message, origin]) } };
  const downloads = [];
  const timers = [];
  const scheduler = { setTimeout: (fn, ms) => { timers.push({ fn, ms }); return timers.length; }, clearTimeout: id => { if (timers[id - 1]) timers[id - 1].cleared = true; } };
  const tab = new AbortController();
  const states = [];
  const controller = new DocumentController({ rpc, events, base, address, signal: tab.signal, update: state => states.push(state), scheduler, download: (url, name) => downloads.push([url, name]), locale });
  controller.frame = frame;
  const fromFrame = data => { for (const fn of listeners) fn({ source: frame.contentWindow, origin: 'http://app.local', data: { v: 1, ...data } }); };
  return { controller, calls, posted, downloads, timers, tab, states, fromFrame, listeners, names: () => calls.map(([n]) => n) };
}

test('AI edit permission is private to the document and destructive operations still ask', async () => {
  const { controller } = harness();
  const { controller: other } = harness();
  const first = controller.propose({ rows: [] });
  controller.setAiPermission('always');
  assert.equal(await first, true);
  assert.equal(await controller.propose({ rows: [] }), true);
  const dangerous = controller.propose({ rows: [], danger: true });
  assert.equal(controller.state.proposal.danger, true);
  controller.decide(false);
  assert.equal(await dangerous, false);
  const separate = other.propose({ rows: [] });
  assert.ok(other.state.proposal);
  other.decide(false);
  assert.equal(await separate, false);
  controller.setAiPermission('ask');
  const reset = controller.propose({ rows: [] });
  assert.ok(controller.state.proposal);
  controller.decide(false);
  assert.equal(await reset, false);
});

test('session addresses only; absolute addresses never borrow a Session', () => {
  assert.deepEqual(parseSource('dsh-resource://file/session/s1/docs/a%20b.docx'), { sessionId: 's1', path: 'docs/a b.docx' });
  assert.deepEqual(parseSource('dsh-resource://file/session/s1//etc/hosts.docx'), { sessionId: 's1', path: '/etc/hosts.docx' });
  assert.throws(() => parseSource('dsh-resource://file/absolute/Users/a.docx'), /Session/);
  assert.throws(() => parseSource('https://x/a.docx'), /Session/);
  assert.throws(() => parseSource('dsh-resource://file/session/s1/%E0%A4%A'), /encoded/);
  assert.throws(() => parseSource('dsh-resource://file/session//a.docx'), /Session/);
});

test('remote failures keep their code for the UI', () => {
  assert.throws(() => unwrapRemote({ status: 'failure', error: { code: 'officesdk/busy', message: 'm' } }), error => error.code === 'officesdk/busy');
  assert.equal(failureKind({ code: 'officesdk/busy' }).kind, 'retry');
  assert.equal(failureKind({ code: 'officesdk/source-changed', details: { draftId: 'd' } }).draftId, 'd');
  assert.equal(failureKind({ code: 'officesdk/forbidden' }).kind, 'fixed');
});

test('opens, reports loaded only after the frame is ready, renews, and releases on tab close', async () => {
  const h = harness();
  const loaded = [];
  await h.controller.open(1, { loaded: v => loaded.push(v), failed: () => loaded.push('failed') });
  assert.deepEqual(h.calls[0], ['open', { workspaceFileScopeId: 's1', path: 'docs/a.docx', revision: 1, mode: 'view' }]);
  assert.match(h.states.at(-1).url, /^http:\/\/app\.local\/mount\/shimo\/doc\/lease1\/index\.html\?revision=1&parentOrigin=http%3A%2F%2Fapp\.local$/);
  h.fromFrame({ type: 'officesdk:ready', id: 'lease1', revision: 2 });
  h.fromFrame({ type: 'legacy:ready', id: 'lease1', revision: 1 });
  assert.deepEqual(loaded, [], 'stale revision and legacy messages are ignored');
  h.fromFrame({ type: 'officesdk:ready', id: 'lease1', revision: 1 });
  assert.deepEqual(loaded, ['v1']);
  const renew = h.timers.find(timer => timer.ms === 300_000);
  await renew.fn(); await tick();
  assert.ok(h.names().includes('renew'));
  h.tab.abort(); await tick();
  assert.deepEqual(h.calls.at(-1), ['release', { workspaceFileScopeId: 's1', id: 'lease1', autosave: true }]);
  assert.equal(h.listeners.size, 0);
});

test('the interface language the tab opened with reaches the frame as lang', async () => {
  for (const [locale, expected] of [['en', 'en'], ['zh', 'zh'], [undefined, null]]) {
    const h = harness({ locale });
    await h.controller.open(1, { loaded: () => {}, failed: () => {} });
    const url = new URL(h.states.at(-1).url);
    assert.equal(url.searchParams.get('lang'), expected, String(locale));
    assert.equal(url.searchParams.get('revision'), '1');
    h.tab.abort();
  }
});

test('a new revision in view mode releases the old session and reopens', async () => {
  const h = harness();
  await h.controller.open(1, {});
  await h.controller.adopt(2, 'v2', {});
  assert.deepEqual(h.names(), ['open', 'release', 'open']);
  assert.equal(h.calls[1][1].id, 'lease1');
  assert.equal(h.calls[2][1].revision, 2);
});

test('rule 1: our own save keeps the editor and rebinds to the new revision', async () => {
  const h = harness({ host: { rebind: args => ({ status: 'success', value: { revision: args.revision, state: 'active', version: 'v2' } }) } });
  await h.controller.open(1, { mode: 'edit' });
  h.fromFrame({ type: 'officesdk:ready', id: 'lease1', revision: 1 });
  h.fromFrame({ type: 'officesdk:dirty', id: 'lease1', revision: 1, dirty: true, editRevision: 1 });
  assert.equal(h.controller.state.dirty, true);
  await h.controller.save();
  assert.deepEqual(h.calls.at(-1), ['save', { workspaceFileScopeId: 's1', id: 'lease1', editRevision: 1, approved: false, overwrite: false }]);
  assert.equal(h.controller.state.dirty, false);
  const loaded = [];
  await h.controller.adopt(2, 'v2', { loaded: v => loaded.push(v) });
  assert.deepEqual(h.names(), ['open', 'save', 'rebind']);
  assert.deepEqual(h.calls[2][1], { workspaceFileScopeId: 's1', id: 'lease1', revision: 2, diverged: false });
  assert.deepEqual(h.posted.at(-1)[0], { type: 'officesdk:command', v: 1, id: 'lease1', revision: 1, command: 'rebind', newRevision: 2 });
  assert.deepEqual(loaded, ['v2']);
  h.fromFrame({ type: 'officesdk:dirty', id: 'lease1', revision: 2, dirty: true, editRevision: 2 });
  assert.equal(h.controller.state.dirty, true, 'messages now validate against the new revision');
});

test('a save the Host made for the AI is ours too: the Host confirms the version and the editor stays', async () => {
  const h = harness();
  await h.controller.open(1, { mode: 'edit' });
  h.fromFrame({ type: 'officesdk:dirty', id: 'lease1', revision: 1, dirty: true, editRevision: 2 });
  // The revision bump can arrive before the save notice; the Host still knows the file is ours.
  await h.controller.adopt(2, 'v7', {});
  assert.deepEqual(h.names(), ['open', 'rebind']);
  assert.equal(h.calls[1][1].diverged, false);
  h.controller.markSaved('v7', 2);
  assert.equal(h.controller.state.dirty, false);
  h.controller.markSaved('v8', 1);
  assert.equal(h.controller.savedVersion, 'v8', 'an older revision saved does not clear newer edits');
});

test('rule 2: a clean editor reopens on an external change', async () => {
  const h = harness({ host: { rebind: () => failure('source-changed') } });
  await h.controller.open(1, { mode: 'edit' });
  await h.controller.adopt(2, 'v9', {});
  assert.deepEqual(h.names(), ['open', 'rebind', 'release', 'open']);
  assert.equal(h.calls[3][1].mode, 'edit');
});

test('rule 3: unsaved edits survive an external change as a diverged session', async () => {
  const h = harness({ host: { rebind: args => args.diverged ? { status: 'success', value: { revision: args.revision, state: 'diverged', version: 'v1' } } : failure('source-changed') } });
  await h.controller.open(1, { mode: 'edit' });
  h.fromFrame({ type: 'officesdk:dirty', id: 'lease1', revision: 1, dirty: true, editRevision: 3 });
  const loaded = [];
  await h.controller.adopt(2, 'v9', { loaded: v => loaded.push(v) });
  assert.deepEqual(h.names(), ['open', 'rebind', 'rebind']);
  assert.deepEqual(h.calls.slice(1).map(([, args]) => args.diverged), [false, true]);
  assert.deepEqual(loaded, ['v1'], 'loaded reports the version the editor still shows');
  assert.equal(h.controller.state.diverged, true);
});

test('download exports the staged edits without asking the frame first', async () => {
  const h = harness();
  await h.controller.open(1, { mode: 'edit' });
  h.fromFrame({ type: 'officesdk:dirty', id: 'lease1', revision: 1, dirty: true, editRevision: 4 });
  await h.controller.exportDownload();
  assert.equal(h.posted.length, 0);
  assert.deepEqual(h.calls.at(-1), ['exportTo', { workspaceFileScopeId: 's1', id: 'lease1', editRevision: 4, target: 'download' }]);
  assert.deepEqual(h.downloads, [['http://app.local/mount/shimo/doc/x/export/t', 'a.docx']]);
});

test('a save outside the workspace says so, keeps the edits, and an approved retry carries the consent', async () => {
  let outside = true;
  const h = harness({ host: { save: args => outside && !args.approved
    ? { status: 'failure', error: { code: 'officesdk/sandbox-denied', message: 'outside', details: { outside: true } } }
    : { status: 'success', value: { version: 'v2', editRevision: args.editRevision } } } });
  await h.controller.open(1, { mode: 'edit' });
  h.fromFrame({ type: 'officesdk:dirty', id: 'lease1', revision: 1, dirty: true, editRevision: 1 });
  await assert.rejects(h.controller.save(), error => error.code === 'officesdk/sandbox-denied' && error.details.outside === true);
  assert.equal(h.controller.state.dirty, true);
  await h.controller.save({ approved: true });
  assert.equal(h.calls.at(-1)[1].approved, true);
  assert.equal(h.controller.state.dirty, false);
  outside = false;
});

test('a parked session reopens in edit mode with its unsaved changes', async () => {
  const h = harness({ open: args => ({ id: 'lease1', format: 'pptx', revision: args.revision, version: 'v1', editRevision: 3, mode: 'edit', unsaved: true, entry: '/shimo/doc/lease1/index.html', expiresInMs: 900_000, capabilities: { edit: true, save: true, export: true } }) });
  await h.controller.open(1, { mode: 'view' });
  assert.deepEqual([h.controller.state.mode, h.controller.state.dirty, h.controller.editRevision], ['edit', true, 3]);
});

test('a parked session the autosave could not write comes back flagged for a confirmed write-back', async () => {
  const lease = args => ({ id: 'lease1', format: 'pptx', revision: args.revision, version: 'v1', editRevision: 3, mode: 'edit', unsaved: true, saveNeedsApproval: true, entry: '/shimo/doc/lease1/index.html', expiresInMs: 900_000, capabilities: { edit: true, save: true, export: true } });
  const h = harness({ open: lease });
  await h.controller.open(1, { mode: 'view' });
  assert.deepEqual([h.controller.state.dirty, h.controller.state.saveNeedsApproval], [true, true]);
  const plain = harness({ open: args => ({ ...lease(args), saveNeedsApproval: undefined }) });
  await plain.controller.open(1, { mode: 'view' });
  assert.equal(plain.controller.state.saveNeedsApproval, undefined);
});

test('closing an edited tab releases its existing lease even after the tab signal aborts', async () => {
  const h = harness();
  await h.controller.open(1, { mode: 'edit' });
  h.controller.set({ dirty: true });
  h.tab.abort();
  await tick();
  assert.deepEqual(h.calls.filter(([name]) => name === 'release'), [
    ['release', { workspaceFileScopeId: 's1', id: 'lease1', autosave: true }],
  ]);
});

test('a late open after the tab closed is released immediately', async () => {
  let resolveOpen;
  const h = harness();
  h.controller.rpc = { call: (_, method, { args }) => method.endsWith('/open') ? new Promise(resolve => { resolveOpen = () => resolve({ status: 'success', value: { id: 'late', format: 'docx', revision: 1, version: 'v1', entry: '/shimo/doc/late/index.html' } }); }) : (h.calls.push([method.split('/')[1], args]), Promise.resolve({ status: 'success', value: {} })) };
  const opening = h.controller.open(1, {});
  h.tab.abort();
  resolveOpen(); await opening; await tick();
  assert.deepEqual(h.calls.at(-1), ['release', { workspaceFileScopeId: 's1', id: 'late', autosave: true }]);
});

test('a superseded open still learns its lease and releases it', async () => {
  const gates = [], opens = [];
  const h = harness();
  const rpc = h.controller.rpc;
  const original = rpc.call;
  rpc.call = async (path, method, payload, signal) => {
    if (method.endsWith('/open')) {
      await new Promise(resolve => gates.push(resolve));
      opens.push(signal);
    }
    return original(path, method, payload, signal);
  };
  const first = h.controller.open(1, { mode: 'view' });
  const second = h.controller.open(1, { mode: 'edit' });
  await tick();
  for (const release of gates) release();
  await Promise.all([first, second]);
  const opened = h.calls.filter(([name]) => name === 'open').length;
  const released = h.calls.filter(([name]) => name === 'release').map(([, args]) => args.id);
  assert.equal(opened, 2);
  assert.deepEqual(released, ['lease1'], 'the superseded lease is released, the current one kept');
  assert.equal(h.controller.lease.id, 'lease2');
  assert.ok(opens.every(signal => !signal.aborted), 'the open RPC itself is never aborted');
});

test('a stalled open fails with the tab instead of hanging, and its late lease is released', async () => {
  let grant;
  const h = harness();
  const rpc = h.controller.rpc;
  const original = rpc.call;
  rpc.call = async (path, method, payload, signal) => {
    if (method.endsWith('/open')) await new Promise(resolve => { grant = resolve; });
    return original(path, method, payload, signal);
  };
  const ready = h.controller.ensureMode('edit');
  await tick();
  h.timers.find(timer => timer.ms === 120_000 && !timer.cleared).fn();
  await assert.rejects(ready, /timed out/);
  grant();
  await tick(); await tick();
  assert.deepEqual(h.calls.filter(([name]) => name === 'release').map(([, args]) => args.id), ['lease1']);
  assert.equal(h.controller.lease, undefined);
});

test('a hand edit ends the AI undo history in the panel; Cmd/Ctrl+S in the editor asks the panel to save', async () => {
  const h = harness();
  await h.controller.open(1, { mode: 'edit' });
  h.fromFrame({ type: 'officesdk:ready', id: 'lease1', revision: 1 });
  h.controller.set({ edits: [{ editId: 'e1', summary: 'AI edit' }] });
  h.fromFrame({ type: 'officesdk:dirty', id: 'lease1', revision: 1, dirty: true, editRevision: 2 });
  assert.equal(h.controller.state.edits.length, 1, 'an AI stage keeps the undo button');
  h.fromFrame({ type: 'officesdk:dirty', id: 'lease1', revision: 1, dirty: true, editRevision: 3, manual: true });
  assert.deepEqual([h.controller.state.edits, h.controller.state.dirty, h.controller.editRevision], [[], true, 3]);
  let requested = 0;
  h.controller.onSaveRequest = () => requested++;
  h.fromFrame({ type: 'officesdk:save-request', id: 'lease1', revision: 1, editRevision: 3 });
  assert.equal(requested, 1);
  h.controller.set({ mode: 'view' });
  h.fromFrame({ type: 'officesdk:save-request', id: 'lease1', revision: 1, editRevision: 3 });
  assert.equal(requested, 1, 'a view session never saves');
});

test('a PPTX save or export first takes the hand edits still in the editor; a silent editor fails the save', async () => {
  const pptx = args => ({ id: 'lease1', format: 'pptx', revision: args.revision, version: 'v1', editRevision: 0, mode: 'edit', entry: '/shimo/doc/lease1/index.html', expiresInMs: 900_000, capabilities: { edit: true, save: true, export: true } });
  const h = harness({ open: pptx });
  await h.controller.open(1, { mode: 'edit' });
  h.fromFrame({ type: 'officesdk:ready', id: 'lease1', revision: 1 });
  h.fromFrame({ type: 'officesdk:dirty', id: 'lease1', revision: 1, dirty: true, editRevision: 1, manual: true });
  const answer = (result, ok = true) => {
    const [command] = h.posted.at(-1);
    assert.deepEqual([command.command, command.op], ['agent', 'flush']);
    h.fromFrame({ type: 'officesdk:agent-result', id: 'lease1', revision: 1, requestId: command.requestId, ok, ...ok ? { result } : result });
  };
  const saving = h.controller.save();
  await tick();
  assert.equal(h.names().includes('save'), false, 'nothing is saved before the editor answers');
  answer({ editRevision: 3 });
  await saving;
  assert.deepEqual(h.calls.at(-1), ['save', { workspaceFileScopeId: 's1', id: 'lease1', editRevision: 3, approved: false, overwrite: false }], 'the save names the revision the editor handed over');

  const exporting = h.controller.exportDownload();
  await tick();
  answer({ editRevision: 4 });
  await exporting;
  assert.deepEqual(h.calls.at(-1)[1].editRevision, 4);

  const failing = h.controller.save();
  await tick();
  answer({ code: 'stage-failed', message: 'offline' }, false);
  await assert.rejects(failing, error => error.code === 'officesdk/stage-failed');
  const silent = h.controller.save();
  await tick();
  h.timers.filter(timer => !timer.cleared && timer.ms === 5000).at(-1).fn();
  await assert.rejects(silent, error => error.code === 'officesdk/editor-failed');
  assert.equal(h.names().filter(name => name === 'save').length, 1, 'no save after a failed or silent hand-over');

  h.fromFrame({ type: 'officesdk:dirty', id: 'lease1', revision: 1, dirty: true, editRevision: 4, manual: true, unstaged: 'offline' });
  assert.equal(h.controller.state.unstaged, 'offline');
  h.fromFrame({ type: 'officesdk:dirty', id: 'lease1', revision: 1, dirty: true, editRevision: 5, manual: true });
  assert.equal(h.controller.state.unstaged, undefined, 'a later stage clears the warning');
});

test('other formats save without asking the editor', async () => {
  const h = harness();
  await h.controller.open(1, { mode: 'edit' });
  h.fromFrame({ type: 'officesdk:ready', id: 'lease1', revision: 1 });
  await h.controller.save();
  assert.equal(h.posted.some(([command]) => command.op === 'flush'), false);
});

test('Kit views: a rebuild request or a stale renew reopens; a Kit save clears unsaved; a disk change is the user\'s choice', async () => {
  let opens = 0;
  const kitLease = args => ({ id: `view${++opens}`, backend: 'kit', format: 'pptx', revision: args.revision, editRevision: 0, mode: 'edit', unsaved: opens === 1, diskChanged: opens === 1,
    entry: `/shimo/kit/static/page/index.html?view=view${opens}`, expiresInMs: 900_000, capabilities: { edit: true, save: true, export: true } });
  const h = harness({ open: kitLease, host: {
    save: () => ({ status: 'success', value: { saved: true, path: '/w/deck.pptx', editRevision: 0 } }),
    renew: () => ({ status: 'failure', error: { code: 'officesdk/edit-conflict', message: 'restarted', details: { rebuild: true } } }),
  } });
  await h.controller.open(1, { mode: 'edit' });
  assert.deepEqual([h.controller.state.backend, h.controller.state.diskChanged, h.controller.state.dirty], ['kit', true, true]);
  assert.match(h.controller.state.url, /view=view1/);
  h.fromFrame({ type: 'officesdk:ready', id: 'view1', revision: 1 });

  await h.controller.resolveDiskChange(false);
  assert.deepEqual(h.calls.at(-1), ['resolveDiskChange', { workspaceFileScopeId: 's1', id: 'view1', useDisk: false }]);
  assert.equal(h.controller.state.diskChanged, undefined);

  const saving = h.controller.save();
  await tick();
  const [flush] = h.posted.at(-1);
  assert.deepEqual([flush.command, flush.op], ['agent', 'flush']);
  h.fromFrame({ type: 'officesdk:agent-result', id: 'view1', revision: 1, requestId: flush.requestId, ok: true, result: { editRevision: 0 } });
  await saving;
  assert.equal(h.controller.state.dirty, false, 'Kit returns no version; the save is current');

  h.fromFrame({ type: 'officesdk:failed', id: 'view1', revision: 1, message: 'service restarted', rebuild: true });
  await tick();
  assert.equal(h.controller.lease.id, 'view2', 'the page asked for a rebuild: the view is reopened');
  assert.equal(h.controller.state.error, undefined);

  h.timers.filter(timer => !timer.cleared && timer.ms === 300_000).at(-1).fn();
  await tick(); await tick();
  assert.equal(h.controller.lease.id, 'view3', 'a stale renew reopens as well');

  await h.controller.resolveDiskChange(true);
  assert.equal(h.controller.lease.id, 'view4', 'using the disk version reloads the editor');
});

test('a DOCX Kit view: save first takes the editor\'s hand edits; an editor that cannot mount falls back to Core\'s read-only preview', async () => {
  let opens = 0;
  const lease = args => args.mode === 'preview'
    ? { id: `core${++opens}`, format: 'docx', revision: args.revision, version: 'v1', editRevision: 0, mode: 'view', entry: '/shimo/doc/core/index.html', expiresInMs: 900_000, capabilities: { edit: false, save: false, export: true } }
    : { id: `view${++opens}`, backend: 'kit', format: 'docx', revision: args.revision, editRevision: 0, mode: 'edit', unsaved: true,
      entry: `/shimo/kit/static/page/index.html?view=view${opens}`, expiresInMs: 900_000, capabilities: { edit: true, save: true, export: true } };
  const h = harness({ open: lease, host: { save: () => ({ status: 'success', value: { saved: true, path: '/w/a.docx' } }) } });
  await h.controller.open(1, { mode: 'edit' });
  h.fromFrame({ type: 'officesdk:ready', id: 'view1', revision: 1 });
  const saving = h.controller.save();
  await tick();
  const [flush] = h.posted.at(-1);
  assert.deepEqual([flush.command, flush.op], ['agent', 'flush'], 'Kit DOCX hand edits are handed over before the save');
  h.fromFrame({ type: 'officesdk:agent-result', id: 'view1', revision: 1, requestId: flush.requestId, ok: true, result: { editRevision: 2 } });
  await saving;

  await h.controller.open(1, { mode: 'edit' });
  h.fromFrame({ type: 'officesdk:failed', id: 'view2', revision: 1, message: 'Writer 系统字体不可用', unsaved: true });
  await tick(); await tick();
  assert.equal(h.calls.filter(([name]) => name === 'open').at(-1)[1].mode, 'preview');
  assert.deepEqual([h.controller.lease.id, h.controller.state.mode, h.controller.state.fallback, h.controller.state.error], ['core3', 'view', 'unsaved', undefined]);
  await h.controller.open(1, { mode: 'edit' });
  assert.equal(h.controller.state.fallback, undefined, 'a fresh open tries the editor again');
});

test('a Kit view whose Core preview is missing too keeps the editor mode, and retry opens the editor again', async () => {
  let opens = 0;
  const h = harness({ host: { open: args => args.mode === 'preview' ? failure('unavailable')
    : { status: 'success', value: { id: `view${++opens}`, backend: 'kit', format: 'docx', revision: args.revision, editRevision: 0, mode: args.mode,
      entry: `/shimo/kit/static/page/index.html?view=view${opens}`, expiresInMs: 900_000, capabilities: { edit: true, save: true, export: true } } } } });
  await h.controller.open(1, { mode: 'edit' });
  h.fromFrame({ type: 'officesdk:failed', id: 'view1', revision: 1, message: 'Kit restarted' });
  await tick(); await tick();
  assert.equal(h.calls.filter(([name]) => name === 'open').at(-1)[1].mode, 'preview');
  assert.equal(h.controller.lease, undefined);
  assert.ok(h.controller.state.error);
  assert.equal(h.controller.state.mode, 'edit', 'a failed preview fallback returns the tab to its editing mode, not preview or view');

  await h.controller.adopt(2, 'v1', {});
  assert.equal(h.calls.filter(([name]) => name === 'open').at(-1)[1].mode, 'edit');
  assert.deepEqual([h.controller.lease?.id, h.controller.state.mode, h.controller.state.error], ['view2', 'edit', undefined], 'retry reopens the Kit editor for editing');
});

test('a Kit view and the disk: its own save is no change; a file changed by someone else reopens the view', async () => {
  let opens = 0;
  const lease = args => ({ id: `view${++opens}`, backend: 'kit', format: 'xlsx', revision: args.revision, version: 'v1', editRevision: 0, mode: 'edit', unsaved: false,
    entry: `/shimo/kit/static/page/index.html?view=view${opens}`, expiresInMs: 900_000, capabilities: { edit: true, save: true, export: true } });
  // The Host decides: revision 3 is someone else's write (source-changed), 5 a restarted document service.
  const rebind = args => args.revision === 3 ? { status: 'failure', error: { code: 'officesdk/source-changed', message: 'changed' } }
    : args.revision === 5 ? { status: 'failure', error: { code: 'officesdk/edit-conflict', message: 'restarted', details: { rebuild: true } } }
    : { status: 'success', value: { revision: args.revision, state: 'active', version: 'v2' } };
  const h = harness({ open: lease, host: { rebind, save: () => ({ status: 'success', value: { saved: true, path: '/w/book.xlsx', version: 'v2' } }) } });
  await h.controller.open(1, { mode: 'edit' });
  h.fromFrame({ type: 'officesdk:ready', id: 'view1', revision: 1 });
  const saving = h.controller.save();
  await tick();
  const [flush] = h.posted.at(-1);
  h.fromFrame({ type: 'officesdk:agent-result', id: 'view1', revision: 1, requestId: flush.requestId, ok: true, result: { editRevision: 0 } });
  await saving;
  await h.controller.adopt(2, 'v2', {});
  assert.deepEqual([h.controller.lease.id, h.names().filter(name => name === 'open').length], ['view1', 1], 'our own write: the editor stays');
  const adopting = h.controller.adopt(3, 'v3-from-another-program', {});
  await tick(); await tick();
  const [handOver] = h.posted.at(-1);
  assert.deepEqual([handOver.command, handOver.op], ['agent', 'flush'], 'what the editor still holds reaches Kit before the reopen');
  h.fromFrame({ type: 'officesdk:agent-result', id: 'view1', revision: 2, requestId: handOver.requestId, ok: true, result: { editRevision: 0 } });
  await adopting;
  assert.equal(h.controller.lease.id, 'view2', 'another program wrote the file: the view reopens and Kit decides (reload or ask)');
  assert.equal(h.calls.filter(([name]) => name === 'open').at(-1)[1].revision, 3);
  h.controller.markSaved('v4', undefined);
  await h.controller.adopt(4, 'v4', {});
  assert.equal(h.controller.lease.id, 'view2', 'an AI save the Host reported is ours too');
  await h.controller.adopt(5, 'v4', {});
  assert.deepEqual([h.controller.lease.id, h.controller.state.error], ['view3', undefined], 'a stale view reopens instead of failing');
});

test('a save whose reply arrives after the view was reopened does not touch the new view', async () => {
  let opens = 0, answer;
  const lease = args => ({ id: `view${++opens}`, backend: 'kit', format: 'docx', revision: args.revision, version: 'v1', editRevision: 0, mode: 'edit', unsaved: true,
    entry: `/shimo/kit/static/page/index.html?view=view${opens}`, expiresInMs: 900_000, capabilities: { edit: true, save: true, export: true } });
  const h = harness({ open: lease, host: { save: () => new Promise(resolve => { answer = resolve; }), rebind: () => ({ status: 'failure', error: { code: 'officesdk/source-changed', message: 'changed' } }) } });
  await h.controller.open(1, { mode: 'edit' });
  const saving = h.controller.save();
  await tick();
  await h.controller.adopt(2, 'v9', {});
  answer({ status: 'success', value: { saved: true, path: '/w/a.docx', version: 'v2' } });
  const result = await saving;
  assert.deepEqual([result.saved, h.controller.lease.id, h.controller.lease.version, h.controller.state.dirty], [true, 'view2', 'v1', true]);
});

/** A Kit view, ready, with a scripted Host save; `answerFlush` plays the page's hand-over. */
async function kitTab(save, extra = {}, host = {}) {
  const lease = args => ({ id: 'view1', backend: 'kit', format: 'xlsx', revision: args.revision, version: 'v1', editRevision: 0, mode: 'edit', unsaved: false, ...extra,
    entry: '/shimo/kit/static/page/index.html?view=view1', expiresInMs: 900_000, capabilities: { edit: true, save: true, export: true } });
  const saves = [];
  const h = harness({ open: lease, host: { ...host, save: args => { saves.push(args); return save(args); } } });
  await h.controller.open(1, { mode: 'edit' });
  h.fromFrame({ type: 'officesdk:ready', id: 'view1', revision: 1 });
  const answerFlush = async () => {
    await tick();
    const [flush] = h.posted.at(-1);
    h.fromFrame({ type: 'officesdk:agent-result', id: 'view1', revision: 1, requestId: flush.requestId, ok: true, result: { editRevision: h.controller.editRevision } });
    await tick(); await tick();
    return flush;
  };
  const pause = () => h.timers.filter(timer => !timer.cleared && timer.ms === 1_000).at(-1);
  return { h, saves, answerFlush, pause };
}

test('autosave: hand edits are written back once the user pauses, without ending their input; AI edits wait for a save', async () => {
  const { h, saves, answerFlush, pause } = await kitTab(() => ({ status: 'success', value: { saved: true, path: '/w/book.xlsx', version: 'v2' } }));
  h.fromFrame({ type: 'officesdk:dirty', id: 'view1', revision: 1, dirty: true, editRevision: 1 });
  assert.equal(pause(), undefined, 'an AI edit does not start an autosave');
  h.fromFrame({ type: 'officesdk:dirty', id: 'view1', revision: 1, dirty: true, editRevision: 2, manual: true });
  const first = pause();
  h.fromFrame({ type: 'officesdk:dirty', id: 'view1', revision: 1, dirty: true, editRevision: 3, manual: true });
  assert.equal(first.cleared, true, 'typing again restarts the pause');
  pause().fn();
  const flush = await answerFlush();
  assert.deepEqual(flush.args, { settle: false }, 'the hand-over does not settle an IME composition or an edited cell');
  assert.deepEqual([saves.length, saves[0].editRevision, saves[0].approved, saves[0].overwrite], [1, 3, false, false]);
  assert.deepEqual([h.controller.state.dirty, h.controller.savedVersion], [false, 'v2']);
});

test('autosave: a file outside the writable roots waits for authorization, then keeps saving; a read-only Session stops it', async () => {
  const outsideAnswer = args => args.approved
    ? { status: 'success', value: { saved: true, path: '/elsewhere/book.xlsx', version: `v${args.editRevision + 1}` } }
    : { status: 'failure', error: { code: 'officesdk/sandbox-denied', message: 'outside', details: { outside: true } } };
  const outside = await kitTab(outsideAnswer, { access: 'outside' });
  assert.deepEqual([outside.h.controller.state.access, outside.h.controller.state.authorized], ['outside', false]);
  outside.h.fromFrame({ type: 'officesdk:dirty', id: 'view1', revision: 1, dirty: true, editRevision: 1 });
  assert.equal(outside.h.controller.state.handEdited, undefined, 'an editor\'s own change is not a hand edit');
  outside.h.fromFrame({ type: 'officesdk:dirty', id: 'view1', revision: 1, dirty: true, editRevision: 1, manual: true });
  outside.pause().fn(); await tick(); await tick();
  assert.equal(outside.h.controller.state.handEdited, true);
  assert.deepEqual([outside.saves.length, outside.h.controller.state.dirty, outside.h.controller.state.saveOutside], [0, true, undefined], 'no save and no prompt before authorization');
  const authorizing = outside.h.controller.authorize();
  await outside.answerFlush(); await authorizing;
  assert.deepEqual([outside.saves.at(-1).approved, outside.h.controller.state.authorized, outside.h.controller.state.dirty], [true, true, false], 'authorizing writes what the editor holds');
  outside.h.fromFrame({ type: 'officesdk:dirty', id: 'view1', revision: 1, dirty: true, editRevision: 2, manual: true });
  outside.pause().fn(); await outside.answerFlush();
  assert.deepEqual([outside.saves.length, outside.saves.at(-1).approved, outside.h.controller.state.dirty], [2, true, false], 'later pauses save with the authorization');

  const allowed = await kitTab(() => ({ status: 'success', value: { saved: true, path: '/w/book.xlsx', version: 'v2' } }), { access: 'allowed' });
  assert.equal(allowed.h.controller.state.authorized, true, 'a writable file is authorized from the start');

  const unreported = await kitTab(outsideAnswer);
  assert.equal(unreported.h.controller.state.authorized, true, 'an unreported access is tried');
  unreported.h.fromFrame({ type: 'officesdk:dirty', id: 'view1', revision: 1, dirty: true, editRevision: 1, manual: true });
  unreported.pause().fn(); await unreported.answerFlush();
  assert.deepEqual([unreported.h.controller.state.access, unreported.h.controller.state.authorized, unreported.h.controller.state.saveOutside], ['outside', false, undefined], 'a refusal turns the tab to asking for authorization, without a prompt');

  const readOnly = await kitTab(() => ({ status: 'failure', error: { code: 'officesdk/sandbox-denied', message: 'read only' } }));
  readOnly.h.fromFrame({ type: 'officesdk:dirty', id: 'view1', revision: 1, dirty: true, editRevision: 1, manual: true });
  readOnly.pause().fn(); await readOnly.answerFlush();
  readOnly.h.fromFrame({ type: 'officesdk:dirty', id: 'view1', revision: 1, dirty: true, editRevision: 2, manual: true });
  assert.equal(readOnly.pause()?.cleared ?? true, true, 'a read-only Session stops autosaving');
  assert.deepEqual([readOnly.saves.length, readOnly.h.controller.state.saveOutside], [1, undefined]);
  assert.deepEqual([readOnly.h.controller.state.access, readOnly.h.controller.state.authorized, readOnly.h.controller.state.autosaveFailed], ['denied', false, undefined],
    'the refusal turns the tab read only, without an autosave failure');
});

/** A Kit tab whose Host answers `access` with `now.access` (a function may defer the answer). */
async function accessTab(save, { format = 'docx', access = 'allowed', answer } = {}) {
  const now = { access, asked: 0 };
  const tab = await kitTab(save, { format, access }, { access: () => { now.asked += 1; return answer ? answer(now) : { status: 'success', value: now.access ? { access: now.access } : {} }; } });
  const poll = async () => {
    const timer = tab.h.timers.filter(t => !t.cleared && t.ms === 3_000).at(-1);
    assert.ok(timer, 'an editing Kit tab keeps asking for its access');
    timer.fn(); await tick(); await tick();
  };
  const flushes = () => tab.h.posted.filter(([message]) => message.op === 'flush').length;
  return { ...tab, now, poll, flushes };
}
const edit = (h, editRevision) => h.fromFrame({ type: 'officesdk:dirty', id: 'view1', revision: 1, dirty: true, editRevision, manual: true });
const saved = () => ({ status: 'success', value: { saved: true, path: '/w/a', version: 'v2' } });

test('read only mid-tab (DOCX, XLSX, PPTX): the tab says so within a poll, autosave neither hands over nor reports a failure; writable again saves the held edits', async () => {
  for (const format of ['docx', 'xlsx', 'pptx']) {
    const { h, saves, now, poll, flushes, pause, answerFlush } = await accessTab(saved, { format });
    assert.deepEqual([h.controller.state.access, h.controller.state.authorized], ['allowed', true], format);
    now.access = 'denied';
    await poll();
    assert.deepEqual([h.controller.state.access, h.controller.state.authorized, h.controller.state.autosaveFailed], ['denied', false, undefined], `${format}: read only`);
    edit(h, 1);
    pause()?.fn(); await tick(); await tick();
    assert.deepEqual([flushes(), saves.length, h.controller.state.autosaveFailed, h.controller.state.dirty], [0, 0, undefined, true], `${format}: nothing handed over, written or reported`);
    now.access = 'allowed';
    await poll();
    assert.deepEqual([h.controller.state.access, h.controller.state.authorized], ['allowed', true], `${format}: writable again`);
    pause().fn(); await answerFlush();
    assert.deepEqual([saves.length, h.controller.state.dirty, h.controller.state.autosaveFailed], [1, false, undefined], `${format}: the held edits are saved without reopening`);
  }
});

test('read only between polls: the autosave asks first and holds; a failure while the Session turned read only is not reported; a real failure still is', async () => {
  const between = await accessTab(saved);
  between.now.access = 'denied';
  edit(between.h, 1);
  between.pause().fn(); await tick(); await tick();
  assert.deepEqual([between.flushes(), between.saves.length, between.h.controller.state.access, between.h.controller.state.autosaveFailed], [0, 0, 'denied', undefined]);

  // The hand-over fails (it timed out in DSH) and by then the Session is read only: no 「未能自动保存」.
  const late = await accessTab(() => failure('editor-failed'));
  edit(late.h, 1);
  late.pause().fn();
  await tick(); await tick();
  late.now.access = 'denied';
  await late.answerFlush(); await tick();
  assert.deepEqual([late.saves.length, late.h.controller.state.autosaveFailed, late.h.controller.state.access], [1, undefined, 'denied']);

  // The same, with a poll asking after the autosave's own question and answering first.
  const held = [];
  const overtaken = await accessTab(() => failure('editor-failed'), { answer: () => new Promise(resolve => held.push(resolve)) });
  edit(overtaken.h, 1);
  overtaken.pause().fn(); await tick();
  held.at(-1)({ status: 'success', value: { access: 'allowed' } });
  await overtaken.answerFlush(); await tick();
  assert.equal(held.length, 2, 'the failed autosave asks again');
  overtaken.h.timers.filter(t => !t.cleared && t.ms === 3_000).at(-1).fn(); await tick();
  held[2]({ status: 'success', value: { access: 'denied' } }); await tick(); await tick();
  held[1]({ status: 'success', value: { access: 'denied' } }); await tick(); await tick();
  assert.deepEqual([overtaken.h.controller.state.autosaveFailed, overtaken.h.controller.state.access], [undefined, 'denied'], 'an overtaken question still sees the read-only state');

  // Writable throughout: a real failure is shown and retried as before.
  const real = await accessTab(() => failure('unavailable'));
  edit(real.h, 1);
  real.pause().fn(); await real.answerFlush(); await tick();
  assert.deepEqual([real.h.controller.state.autosaveFailed, real.h.controller.state.access], ['unavailable', 'allowed']);
  assert.ok(real.h.timers.some(timer => !timer.cleared && timer.ms === 5_000), 'retried');
});

test('access refresh: outside keeps the authorization the user gave; a held autosave stays held; a stale answer or an unknown one changes nothing; closing stops the polls', async () => {
  const approved = await accessTab(saved, { access: 'outside' });
  approved.h.controller.autosaveApproved = true;
  approved.now.access = 'denied'; await approved.poll();
  approved.now.access = 'outside'; await approved.poll();
  assert.deepEqual([approved.h.controller.state.access, approved.h.controller.state.authorized, approved.h.controller.autosaveApproved], ['outside', true, true], 'no second authorization after a read-only spell');
  const asking = await accessTab(saved, { access: 'outside' });
  asking.now.access = 'denied'; await asking.poll();
  asking.now.access = 'outside'; await asking.poll();
  assert.deepEqual([asking.h.controller.state.access, asking.h.controller.state.authorized], ['outside', false], 'not authorized before: still asks');

  const paused = await accessTab(saved);
  paused.now.access = 'denied'; await paused.poll();
  edit(paused.h, 1);
  paused.h.controller.pauseAutosave(true);
  paused.now.access = 'allowed'; await paused.poll();
  assert.equal(paused.pause(), undefined, 'Save As holds the autosave even when writable again');
  paused.h.controller.pauseAutosave(false);
  paused.pause().fn(); await paused.answerFlush();
  assert.deepEqual([paused.saves.length, paused.h.controller.state.dirty], [1, false], 'saved once after the hold');

  let release;
  const stale = await accessTab(saved, { answer: () => new Promise(resolve => { release = resolve; }) });
  const timer = stale.h.timers.filter(t => !t.cleared && t.ms === 3_000).at(-1);
  timer.fn(); await tick();
  stale.h.controller.lease = { ...stale.h.controller.lease, id: 'view2' };
  release({ status: 'success', value: { access: 'denied' } }); await tick(); await tick();
  assert.deepEqual([stale.h.controller.state.access, stale.h.controller.state.authorized], ['allowed', true], 'an answer for a replaced lease is dropped');

  const unknown = await accessTab(saved, { access: 'allowed' });
  unknown.now.access = undefined; await unknown.poll();
  assert.deepEqual([unknown.h.controller.state.access, unknown.h.controller.state.authorized], ['allowed', true], 'an unknown answer keeps the access');
  unknown.h.controller.stop();
  assert.equal(unknown.h.timers.filter(t => !t.cleared && t.ms === 3_000).length, 0, 'closing the tab stops the polls');

  const banner = await accessTab(saved);
  banner.now.access = 'denied'; await banner.poll();
  banner.h.controller.set({ saveState: 'failed', saveFailed: 'read only' });
  banner.now.access = 'allowed'; await banner.poll();
  assert.equal(banner.h.controller.state.saveState, undefined, 'a read-only notice goes once the Session may write');
});

test('autosave: a rewritten file is never overwritten; busy tries again', async () => {
  const rewritten = await kitTab(() => ({ status: 'failure', error: { code: 'officesdk/source-changed', message: 'changed', details: { overwrite: true } } }));
  rewritten.h.fromFrame({ type: 'officesdk:dirty', id: 'view1', revision: 1, dirty: true, editRevision: 1, manual: true });
  rewritten.pause().fn(); await rewritten.answerFlush();
  assert.deepEqual([rewritten.saves.length, rewritten.saves[0].overwrite, rewritten.h.controller.state.overwritePrompt, rewritten.h.controller.state.dirty], [1, false, undefined, true]);

  let calls = 0;
  const busy = await kitTab(() => ++calls === 1 ? { status: 'failure', error: { code: 'officesdk/busy', message: 'AI running' } } : { status: 'success', value: { saved: true, path: '/w/book.xlsx', version: 'v2' } });
  busy.h.fromFrame({ type: 'officesdk:dirty', id: 'view1', revision: 1, dirty: true, editRevision: 1, manual: true });
  busy.pause().fn(); await busy.answerFlush();
  busy.pause().fn(); await busy.answerFlush();
  assert.deepEqual([busy.saves.length, busy.h.controller.state.dirty], [2, false]);
});

test('autosave: a failed write is shown, never looks saved; a failure that may pass is retried and the notice clears on success', async () => {
  let calls = 0;
  const flaky = await kitTab(() => ++calls === 1 ? { status: 'failure', error: { code: 'officesdk/unavailable', message: 'document service down' } } : { status: 'success', value: { saved: true, path: '/w/book.xlsx', version: 'v2' } });
  flaky.h.fromFrame({ type: 'officesdk:dirty', id: 'view1', revision: 1, dirty: true, editRevision: 1, manual: true });
  flaky.pause().fn(); await flaky.answerFlush();
  assert.deepEqual([flaky.h.controller.state.dirty, flaky.h.controller.state.autosaveFailed], [true, 'document service down']);
  const retry = flaky.h.timers.filter(timer => !timer.cleared && timer.ms === 5_000).at(-1);
  assert.ok(retry, 'a transient failure schedules a retry');
  retry.fn(); await flaky.answerFlush();
  assert.deepEqual([flaky.saves.length, flaky.h.controller.state.dirty, flaky.h.controller.state.autosaveFailed], [2, false, undefined]);

  const conflict = await kitTab(() => ({ status: 'failure', error: { code: 'officesdk/conflict', message: 'kept as a draft' } }));
  conflict.h.fromFrame({ type: 'officesdk:dirty', id: 'view1', revision: 1, dirty: true, editRevision: 1, manual: true });
  conflict.pause().fn(); await conflict.answerFlush();
  assert.equal(conflict.h.controller.state.autosaveFailed, 'kept as a draft');
  assert.equal(conflict.h.timers.filter(timer => !timer.cleared && timer.ms === 5_000).length, 0, 'a lasting failure waits for the next hand edit');
});

test('DOCX, XLSX and PPTX: an editor that cannot mount falls back to the preview and the Kit view is released without autosave', async () => {
  for (const format of ['docx', 'xlsx', 'pptx']) {
    let opens = 0;
    const lease = args => args.mode === 'preview'
      ? { id: `core${++opens}`, format, revision: args.revision, version: 'v1', editRevision: 0, mode: 'view', entry: '/shimo/doc/core/index.html', expiresInMs: 900_000, capabilities: { edit: false, save: false, export: true } }
      : { id: `view${++opens}`, backend: 'kit', format, revision: args.revision, version: 'v1', editRevision: 0, mode: 'edit', unsaved: false,
        entry: `/shimo/kit/static/page/index.html?view=view${opens}`, expiresInMs: 900_000, capabilities: { edit: true, save: true, export: true } };
    const h = harness({ open: lease });
    await h.controller.open(1, { mode: 'edit' });
    h.fromFrame({ type: 'officesdk:failed', id: 'view1', revision: 1, message: 'the editor bundle failed to load' });
    await tick(); await tick();
    assert.deepEqual([h.controller.lease.id, h.controller.state.mode, h.controller.state.fallback], ['core2', 'view', 'preview'], format);
    assert.deepEqual(h.calls.find(([name]) => name === 'release'), ['release', { workspaceFileScopeId: 's1', id: 'view1', autosave: false }], format);
    h.tab.abort();
    assert.deepEqual(h.calls.filter(([name]) => name === 'release').at(-1)[1], { workspaceFileScopeId: 's1', id: 'core2', autosave: true }, `${format}: later releases save as usual`);
  }
});

test('the update panel follows edits, saves and closing other tabs while the current tab stays clean', () => {
  const current = harness().controller;
  const other = harness().controller;
  const tabs = new Map([['current', current]]);
  const watchers = new Set();
  const states = [];
  const off = watchTabEdits(tabs, watchers, () => states.push(anyUnsaved(tabs)));
  tabs.set('other', other);
  for (const notify of watchers) notify();
  other.set({ dirty: true });
  other.set({ dirty: false });
  other.set({ unstaged: 'pending' });
  tabs.delete('other');
  for (const notify of watchers) notify();
  other.set({ dirty: true });
  off();
  current.set({ dirty: true });
  assert.deepEqual(states, [false, false, true, false, true, false]);
  assert.equal(watchers.size, 0);
});

test('the update entry shows only for a confirmed newer release the owner may install, while it installs, and once ready (V3.0)', () => {
  const ready = { current: '0.0.28', latest: '0.0.29', available: true, canUpdate: true };
  assert.equal(updateEntry(ready), 'available');
  assert.equal(updateEntry(undefined), 'hidden', 'still checking');
  assert.equal(updateEntry({ current: '0.0.28', available: false, canUpdate: true }), 'hidden', 'latest, a failed check or a fixed install');
  assert.equal(updateEntry({ current: '0.0.28', latest: undefined, available: false, canUpdate: true }), 'hidden', 'a failed check is never an update');
  assert.equal(updateEntry({ ...ready, canUpdate: false }), 'hidden', 'only the owner can install');
  assert.equal(updateEntry({ ...ready, canUpdate: undefined }), 'hidden', 'an older Host that does not say');
  assert.equal(updateEntry({ current: '0.0.28', available: false, installed: '0.0.29', canUpdate: true }), 'ready', 'installed, waiting to take effect: the entry keeps the task state');
  assert.equal(updateEntry({ current: '0.0.28', available: false, installed: '0.0.29', canUpdate: false }), 'hidden', 'not for a caller who may not install');
  assert.equal(updateEntry({ ...ready, updating: true }), 'updating');
  assert.equal(updateEntry({ ...ready, failed: 'x' }), 'available', 'a failed install can be retried');
});

test('the update bubble: Host values only, notes before installation, off while unsaved, then progress and ready (V3.0)', () => {
  const check = { current: '0.0.28', latest: '0.0.29', releasedAt: '2026-11-12T03:00:00.000Z', available: true, canUpdate: true };
  assert.equal(updatePanelView(check, 'available').action.disabled, true);
  Object.assign(check, { highlights: ['功能一'], highlightsEn: ['Feature one'] });
  assert.deepEqual(updatePanelView(check, 'available', false, 'en').notes, ['Feature one']);
  assert.deepEqual(updatePanelView({ ...check, installed: '0.0.28' }, 'ready').notes, []);
  const open = updatePanelView(check, 'available', false);
  assert.deepEqual([open.version, open.date, open.notes, open.action, open.message], ['0.0.29', '2026-11-12T03:00:00.000Z', ['功能一'], { label: 'updateNow', disabled: false, install: true }, undefined]);
  const unsaved = updatePanelView(check, 'available', true);
  assert.deepEqual([unsaved.action.disabled, unsaved.message], [true, { key: 'updateUnsaved', tone: 'warning' }]);
  const busy = updatePanelView({ ...check, updating: true }, 'updating');
  assert.deepEqual([busy.action.disabled, busy.action.busy, busy.message.key], [true, true, 'updatePreparingNote']);
  const restart = updatePanelView({ ...check, available: false, installed: '0.0.29', restartRequired: true }, 'ready');
  assert.deepEqual(restart.notes, [], 'release notes are shown before installation only');
  assert.deepEqual([restart.version, restart.action, restart.message.key], ['0.0.29', { label: 'updateGotIt', close: true }, 'updateReadyRestart']);
  assert.equal(updatePanelView({ ...check, installed: '0.0.29', restartRequired: false }, 'ready').message.key, 'updateReadyLoaded');
  // No date without one from the Host (a git install), junk dropped; highlights only when the Host has them.
  assert.equal(updatePanelView({ ...check, releasedAt: undefined }, 'available').date, undefined);
  assert.equal(updatePanelView({ ...check, releasedAt: 'soon' }, 'available').date, undefined);
  assert.deepEqual(updatePanelView({ ...check, highlights: ['a', '', 3, 'b'] }, 'available').notes, ['a', 'b']);
  // The date belongs to `latest`: a ready copy older than a newer release shows no date.
  assert.equal(updatePanelView({ ...check, latest: '0.0.30', installed: '0.0.29' }, 'ready').date, undefined);
  // A refused install says why inside the bubble.
  assert.deepEqual(updatePanelView({ ...check, failed: '仍有未保存修改' }, 'available').message, { text: '仍有未保存修改', tone: 'warning' });
});

test('an update waits while any open tab has edits not yet written or not yet handed over; a failed save keeps it waiting', () => {
  const tabs = new Map([['a', { state: { dirty: false } }], ['b', { state: undefined }]]);
  assert.equal(anyUnsaved(tabs), false);
  tabs.set('c', { state: { dirty: true } });
  assert.equal(anyUnsaved(tabs), true);
  tabs.set('c', { state: { dirty: true, saveState: 'failed' } });
  assert.equal(anyUnsaved(tabs), true, 'a save that failed keeps the edits and the block');
  tabs.set('c', { state: { dirty: false, unstaged: 'retrying' } });
  assert.equal(anyUnsaved(tabs), true, 'edits not yet handed to DSH');
});

test('pen/eye: a Kit tab leaves edit only once every edit is on disk, else it stays editing and says so', async () => {
  const opened = h => h.calls.filter(([name]) => name === 'open').map(([, args]) => args.mode);
  const saved = () => ({ status: 'success', value: { saved: true, path: '/w/book.xlsx', version: 'v2' } });

  const clean = await kitTab(saved, { access: 'allowed' });
  const leaving = clean.h.controller.switchMode('view');
  await clean.answerFlush();
  assert.equal(await leaving, true);
  assert.deepEqual([opened(clean.h), clean.saves.length], [['edit', 'view'], 0], 'nothing to write: reopens read only');

  const dirty = await kitTab(saved, { access: 'allowed' });
  dirty.h.fromFrame({ type: 'officesdk:dirty', id: 'view1', revision: 1, dirty: true, editRevision: 1, manual: true });
  const pending = dirty.pause();
  const writing = dirty.h.controller.switchMode('view');
  assert.equal(dirty.h.controller.state.switching, true);
  assert.equal(await dirty.h.controller.switchMode('view'), false, 'one switch at a time');
  await dirty.answerFlush(); await dirty.answerFlush();
  assert.equal(await writing, true);
  assert.equal(pending.cleared, true, 'the pending autosave gives way to the save');
  assert.deepEqual([dirty.saves.length, dirty.saves[0].editRevision, opened(dirty.h), dirty.h.controller.state.switching], [1, 1, ['edit', 'view'], false]);

  const refused = await kitTab(() => ({ status: 'success', value: { saved: false } }), { access: 'allowed' });
  refused.h.fromFrame({ type: 'officesdk:dirty', id: 'view1', revision: 1, dirty: true, editRevision: 1, manual: true });
  const staying = refused.h.controller.switchMode('view');
  await refused.answerFlush(); await refused.answerFlush();
  assert.equal(await staying, false);
  assert.deepEqual([opened(refused.h), refused.h.controller.state.mode, refused.h.controller.state.modeBlocked], [['edit'], 'edit', true], 'a save that wrote nothing keeps the edits in edit');

  const outside = await kitTab(saved, { access: 'outside' });
  outside.h.fromFrame({ type: 'officesdk:dirty', id: 'view1', revision: 1, dirty: true, editRevision: 1, manual: true });
  const blocked = outside.h.controller.switchMode('view');
  await outside.answerFlush();
  assert.equal(await blocked, false);
  assert.deepEqual([outside.saves.length, opened(outside.h), outside.h.controller.state.modeBlocked], [0, ['edit'], true], 'unauthorized edits are never written to switch');

  const failing = await kitTab(() => ({ status: 'failure', error: { code: 'officesdk/unavailable', message: 'disk full' } }), { access: 'allowed' });
  failing.h.fromFrame({ type: 'officesdk:dirty', id: 'view1', revision: 1, dirty: true, editRevision: 1, manual: true });
  const thrown = assert.rejects(failing.h.controller.switchMode('view'), /disk full/);
  await failing.answerFlush(); await failing.answerFlush();
  await thrown;
  assert.deepEqual([opened(failing.h), failing.h.controller.state.switching], [['edit'], false], 'a failed save keeps the tab editing');
});

test('pen/eye: read to edit reopens editing; the selected mode does nothing', async () => {
  const h = harness();
  await h.controller.open(1, { mode: 'view' });
  assert.equal(await h.controller.switchMode('view'), false);
  assert.equal(await h.controller.switchMode('edit'), true);
  assert.deepEqual(h.calls.filter(([name]) => name === 'open').map(([, args]) => args.mode), ['view', 'edit']);
  assert.equal(h.controller.state.mode, 'edit');
});

test('pen/eye buttons: Kit tabs only, mutually pressed, and the eye waits for edits that cannot be written', () => {
  const kit = { backend: 'kit', mode: 'edit', loading: false, authorized: true, access: 'allowed' };
  assert.deepEqual(modeButtons(kit), { edit: { pressed: true, disabled: false }, view: { pressed: false, disabled: false } });
  assert.deepEqual(modeButtons({ ...kit, mode: 'view' }), { edit: { pressed: false, disabled: false }, view: { pressed: true, disabled: false } });
  assert.equal(modeButtons({ ...kit, backend: undefined }), undefined, 'Core tabs keep their own flow');
  assert.equal(modeButtons({ ...kit, mode: 'view', fallback: 'preview' }), undefined, 'a fallback preview has no editor to switch to');
  assert.equal(modeButtons({ ...kit, mode: 'preview' }), undefined);
  assert.equal(modeButtons({ ...kit, dirty: true }).view.disabled, false, 'authorized edits are saved on the way');
  assert.deepEqual(modeButtons({ ...kit, dirty: true, authorized: false, access: 'outside' }).view, { pressed: false, disabled: true, reason: 'unsaved' });
  assert.equal(modeButtons({ ...kit, dirty: true, access: 'denied' }).view.reason, 'unsaved', 'a read-only Session never writes');
  assert.equal(modeButtons({ ...kit, switching: true }).edit.disabled, true);
  assert.equal(modeButtons({ ...kit, loading: true }).view.disabled, true);
});

test('pen/eye: the old view is released before the file opens again, and a failed switch can be retried', async () => {
  let release;
  const h = harness({ host: { release: () => new Promise(resolve => { release = resolve; }) } });
  await h.controller.open(1, { mode: 'edit' });
  const switching = h.controller.switchMode('view');
  await tick();
  assert.deepEqual(h.names().slice(-2), ['open', 'release'], 'no open while Kit is still closing the old view');
  release({ status: 'success', value: {} });
  assert.equal(await switching, true);
  assert.deepEqual(h.names().slice(-2), ['release', 'open']);

  let fail = true;
  const g = harness({ host: { open: args => fail && args.mode === 'view'
    ? { status: 'failure', error: { code: 'officesdk/unavailable', message: 'aborted' } }
    : { status: 'success', value: { id: `l${args.mode}`, format: 'xlsx', revision: args.revision, version: 'v1', editRevision: 0, mode: args.mode, entry: `/shimo/doc/l${args.mode}/index.html`, expiresInMs: 900_000, capabilities: { edit: true } } } } });
  await g.controller.open(1, { mode: 'edit' });
  await g.controller.switchMode('view');
  assert.deepEqual([g.controller.state.mode, g.controller.state.error, g.controller.lease], ['view', 'aborted', undefined]);
  fail = false;
  assert.equal(await g.controller.switchMode('edit'), true, 'the pen still works without a session');
  assert.deepEqual([g.controller.state.mode, g.controller.state.error, g.controller.lease.id], ['edit', undefined, 'ledit']);
  await g.controller.switchMode('view');
  assert.equal(g.controller.state.mode, 'view');
});

test('pen/eye: a held-back switch is said only until the edits are on disk', async () => {
  const outside = await kitTab(() => ({ status: 'success', value: { saved: true, path: '/w/book.xlsx', version: 'v2' } }), { access: 'outside' });
  outside.h.fromFrame({ type: 'officesdk:dirty', id: 'view1', revision: 1, dirty: true, editRevision: 1, manual: true });
  const blocked = outside.h.controller.switchMode('view');
  await outside.answerFlush(); await blocked;
  assert.equal(outside.h.controller.state.modeBlocked, true);
  const authorizing = outside.h.controller.authorize();
  await outside.answerFlush(); await authorizing;
  assert.deepEqual([outside.h.controller.state.dirty, outside.h.controller.state.modeBlocked], [false, undefined], 'saved: the notice goes');
  outside.h.fromFrame({ type: 'officesdk:dirty', id: 'view1', revision: 1, dirty: true, editRevision: 2, manual: true });
  assert.equal(outside.h.controller.state.modeBlocked, undefined, 'a later edit does not bring it back');
});

test('pen/eye: the closing view is hidden while Kit releases it, and AI requests and source changes wait for the switch', async () => {
  let release;
  const h = harness({ host: { release: () => new Promise(resolve => { release = resolve; }) } });
  await h.controller.open(1, { mode: 'edit' });
  const switching = h.controller.switchMode('view');
  await tick();
  assert.deepEqual([h.controller.state.loading, h.controller.state.url], [true, undefined], 'no editor to type into while it closes');
  const ai = h.controller.ensureMode('edit', new AbortController().signal).catch(() => {});
  const adopting = h.controller.adopt(1, 'v1', {});
  await tick();
  assert.equal(h.names().filter(name => name === 'open').length, 1, 'nothing else opens before the switch is done');
  release({ status: 'success', value: {} });
  await switching; await tick();
  assert.equal(h.names().filter(name => name === 'open').length, 2, 'the AI waits for the new view to be ready');
  h.fromFrame({ type: 'officesdk:ready', id: 'lease2', revision: 1 });
  await adopting; await tick(); await tick();
  assert.deepEqual(h.calls.filter(([name]) => name === 'open').map(([, args]) => args.mode), ['edit', 'view', 'edit'], 'the AI then gets its editing view');
  void ai;
});

test('pen/eye buttons wait while the header saves or authorizes', () => {
  const kit = { backend: 'kit', mode: 'edit', authorized: true, access: 'allowed' };
  assert.deepEqual([modeButtons(kit, true).edit.disabled, modeButtons(kit, true).view.disabled], [true, true]);
});

test('pen/eye: a source change from the switch\'s own save waits until the new view is ready', async () => {
  const { h, answerFlush } = await kitTab(() => ({ status: 'success', value: { saved: true, path: '/w/book.xlsx', version: 'v2' } }), { access: 'allowed' });
  h.fromFrame({ type: 'officesdk:dirty', id: 'view1', revision: 1, dirty: true, editRevision: 1 });
  const switching = h.controller.switchMode('view');
  await answerFlush(); await answerFlush();
  // The save rewrote the file: DSH bumps the revision while the view reopens.
  const adopting = h.controller.adopt(2, 'v2', {});
  await tick(); await tick();
  assert.equal(h.names().includes('rebind'), false, 'no rebind for a frame that has not loaded');
  h.fromFrame({ type: 'officesdk:ready', id: 'view1', revision: 1 });
  assert.equal(await switching, true);
  await adopting;
  assert.deepEqual([h.controller.state.loading, h.controller.state.ready, h.controller.revision], [false, true, 2]);
});

test('pen/eye: a tab closed mid-switch opens nothing, and a cancelled AI request stops waiting for the switch', async () => {
  let release;
  const h = harness({ host: { release: () => new Promise(resolve => { release = resolve; }) } });
  await h.controller.open(1, { mode: 'edit' });
  const switching = h.controller.switchMode('view');
  await tick();
  const ai = new AbortController();
  const waiting = h.controller.ensureMode('edit', ai.signal);
  ai.abort(new Error('cancelled'));
  await assert.rejects(waiting, /cancelled/);
  h.tab.abort();
  release({ status: 'success', value: {} });
  assert.equal(await switching, false);
  assert.deepEqual(h.names().filter(name => name === 'open').length, 1, 'no session is opened for a closed tab');
});

test('P toolbar: modules keep the fixed order and only those offered here appear', () => {
  assert.deepEqual(MODULES, ['undo', 'redo', 'save', 'saveAs', 'folder', 'print', 'layout', 'focus', 'slideshow', 'mode', 'update', 'account']);
  assert.deepEqual(offeredModules({ update: true, mode: true }), ['mode', 'update']);
  // A Core tab has no pen/eye, so the menu does not offer them.
  assert.deepEqual(offeredModules({ mode: false, update: true }), ['update']);
});

test('V3.1 toolbar: Open, Undo and Redo are withdrawn from quick access and customization', () => {
  assert.deepEqual(WITHDRAWN_MODULES, ['folder', 'undo', 'redo', 'account', 'skills']);
  const all = { undo: true, redo: true, save: true, saveAs: true, folder: true, print: true, focus: true, mode: true, update: true, account: true, skills: true };
  assert.deepEqual(offeredModules(all), ['save', 'saveAs', 'print', 'focus', 'mode', 'update']);
  const hidden = createToolbarPrefs(() => undefined).hidden();
  assert.deepEqual(offeredModules(all).filter(id => !hidden.includes(id)), ['save', 'focus']);
  // The tab still supports it (the open flow stays); only the toolbar and its menu no longer show it.
  const kit = { backend: 'kit', format: 'docx', mode: 'edit', capabilities: { save: true } };
  assert.deepEqual(editButtons(kit).folder, { disabled: false });
  assert.equal(offeredModules({ folder: Boolean(editButtons(kit).folder) }).includes('folder'), false);
});

test('the slideshow shortcut follows Focus and is available for ready, editable or read-only PPT files', () => {
  const state = { backend: 'kit', format: 'pptx', ready: true, mode: 'view', capabilities: { save: false } };
  for (const mode of ['view', 'edit']) {
    assert.deepEqual(editButtons({ ...state, mode }).slideshow, { disabled: false });
  }
  assert.equal(editButtons({ ...state, ready: false }).slideshow.disabled, true);
  for (const patch of [{ format: 'docx' }, { format: 'xlsx' }, { fallback: 'preview' }, { backend: 'core' }]) {
    assert.equal(editButtons({ ...state, ...patch }).slideshow, undefined);
  }
  const prefs = createToolbarPrefs(() => undefined);
  assert.deepEqual(offeredModules({ save: true, focus: true, slideshow: true, mode: true }).filter(id => !prefs.hidden().includes(id)), ['save', 'focus', 'slideshow']);
});

test('P toolbar: busy is the header\'s own work, a mode switch or a reload, never a Kit autosave', () => {
  assert.equal(toolbarBusy(false, { dirty: true, authorized: true }), false);
  assert.equal(toolbarBusy(true, {}), true);
  assert.equal(toolbarBusy(false, { switching: true }), true);
  assert.equal(toolbarBusy(false, { loading: true }), true);
});

test('P toolbar: shown/hidden choices are shared, persist, survive refusing storage, and follow other windows', () => {
  const store = new Map();
  const storage = { getItem: key => store.get(key) ?? null, setItem: (key, value) => store.set(key, value) };
  const prefs = createToolbarPrefs(() => storage);
  let changes = 0;
  prefs.subscribe(() => { changes += 1; });
  assert.deepEqual(prefs.hidden(), ['saveAs', 'mode', 'print', 'update']);
  prefs.toggle('update');
  assert.deepEqual(prefs.hidden(), ['saveAs', 'mode', 'print']);
  assert.deepEqual(JSON.parse(store.get(prefs.key)), { hidden: ['saveAs', 'mode', 'print'] });
  // A second header (another tab or suite) reads the same choice.
  assert.deepEqual(createToolbarPrefs(() => storage).hidden(), ['saveAs', 'mode', 'print']);
  prefs.toggle('update');
  assert.deepEqual(prefs.hidden(), ['saveAs', 'mode', 'print', 'update']);
  // Another window wrote the stored value; ids from a newer version survive this one's writes.
  store.set(prefs.key, JSON.stringify({ hidden: ['mode', 'futureModule', 7] }));
  prefs.reload();
  assert.deepEqual(prefs.hidden(), ['mode', 'futureModule']);
  prefs.toggle('update');
  assert.deepEqual(JSON.parse(store.get(prefs.key)).hidden, ['mode', 'futureModule', 'update']);
  prefs.toggle('update');
  assert.equal(changes, 5);
  // Private mode, quota or a refusal: the page keeps working in memory.
  const refusing = createToolbarPrefs(() => ({ getItem: () => { throw new Error('denied'); }, setItem: () => { throw new Error('denied'); } }));
  refusing.toggle('mode');
  assert.deepEqual(refusing.hidden(), ['saveAs', 'print', 'update']);
  assert.deepEqual(createToolbarPrefs(() => ({ getItem: () => '{not json' })).hidden(), ['saveAs', 'mode', 'print', 'update']);
  assert.deepEqual(createToolbarPrefs(() => undefined).hidden(), ['saveAs', 'mode', 'print', 'update']);
});

test('compact toolbar migrates implicit visibility while keeping previous customization', () => {
  const old = new Map([['officesdk:toolbar:v3.1', JSON.stringify({ hidden: ['print', 'futureModule'] })]]);
  const storage = { getItem: key => old.get(key), setItem: (key, value) => old.set(key, value) };
  const prefs = createToolbarPrefs(() => storage);
  assert.deepEqual(prefs.hidden(), ['saveAs', 'mode', 'print', 'futureModule']);
  prefs.toggle('saveAs');
  assert.deepEqual(createToolbarPrefs(() => storage).hidden(), ['mode', 'print', 'futureModule']);
});

test('P toolbar: a popover sits under its trigger, right edges aligned, 8px inside the viewport, scrolling when short of room', () => {
  const viewport = { width: 800, height: 600 };
  const trigger = { left: 700, right: 725, top: 10, bottom: 34 };
  assert.deepEqual(placeMenu({ trigger, width: 179, height: 80, viewport }), { left: 546, top: 39, maxWidth: 784, maxHeight: 553, height: 80 });
  // Near the left edge it moves right to keep 8px; near the right edge it stays 8px inside.
  assert.equal(placeMenu({ trigger: { right: 100, bottom: 34 }, width: 242, height: 80, viewport }).left, 8);
  assert.equal(placeMenu({ trigger: { right: 798, bottom: 34 }, width: 242, height: 80, viewport }).left, 800 - 8 - 242);
  // Wider than the viewport allows: capped at the viewport less both margins.
  const narrow = placeMenu({ trigger: { right: 300, bottom: 34 }, width: 400, height: 80, viewport: { width: 320, height: 600 } });
  assert.deepEqual([narrow.left, narrow.maxWidth], [8, 304]);
  // Little room below: the popover keeps its top and scrolls inside.
  assert.equal(placeMenu({ trigger: { right: 725, bottom: 560 }, width: 179, height: 200, viewport }).maxHeight, 600 - 8 - 565);
  // An ancestor transform shifted the fixed box by (+30, +12): the next placement corrects for it.
  const place = placeMenu({ trigger, width: 179, height: 80, viewport });
  const fixed = placeMenu({ trigger, width: 179, height: 80, viewport, target: place, measured: { left: place.left + 30, top: place.top + 12 } });
  assert.deepEqual([fixed.left, fixed.top], [place.left - 30, place.top - 12]);
});

test('P toolbar: three-dot menu keys move, wrap, toggle and close; toolbar arrows move between buttons', () => {
  assert.deepEqual(menuKey('ArrowDown', 1, 2), { index: 0 });
  assert.deepEqual(menuKey('ArrowUp', 0, 2), { index: 1 });
  assert.deepEqual(menuKey('Home', 1, 3), { index: 0 });
  assert.deepEqual(menuKey('End', 0, 3), { index: 2 });
  assert.deepEqual(menuKey('Enter', 1, 2), { index: 1, toggle: true, close: 'restore' });
  assert.deepEqual(menuKey(' ', 0, 2), { index: 0, toggle: true, close: 'restore' });
  assert.deepEqual(menuKey('Escape', 0, 2), { index: 0, close: 'restore' });
  assert.deepEqual(menuKey('Tab', 0, 2), { index: 0, close: 'leave' });
  assert.equal(menuKey('a', 0, 2), undefined);
  assert.equal(rovingKey('ArrowRight', 3, 4), 0);
  assert.equal(rovingKey('ArrowLeft', 0, 4), 3);
  assert.equal(rovingKey('Home', 2, 4), 0);
  assert.equal(rovingKey('End', 0, 4), 3);
  assert.equal(rovingKey('ArrowDown', 0, 4), undefined);
  assert.equal(cancelKey({ key: 'Escape' }), true);
  assert.equal(cancelKey({ key: '.', metaKey: true }), true);
  assert.equal(cancelKey({ key: '.' }), false);
});

test('P toolbar: hiding a module hides only its entry — pen/eye keep the mode, update stays absent without a release', () => {
  const kit = { backend: 'kit', mode: 'view' };
  // The mode buttons' state does not depend on whether the module is shown.
  assert.deepEqual(modeButtons(kit).view, { pressed: true, disabled: false });
  // Restoring the update entry with no newer release still shows nothing.
  assert.equal(updateEntry({ available: false, canUpdate: true }), 'hidden');
});

test('P toolbar theme: DSH\'s effective theme decides; the system scheme only when no DSH theme presenter runs; popovers stay white', () => {
  assert.match(CSS, /\[data-ds-dark-theme\] \.sbp\{--sbp-icon:#d5d8dc;/);
  assert.match(CSS, /@media \(prefers-color-scheme:dark\)\{:root:not\(\[data-ds-theme-source\]\) body:not\(\[data-ds-dark-theme\]\) \.sbp\{--sbp-icon:#d5d8dc;/);
  assert.match(CSS, /\.sbp\{--sbp-icon:#3c4046;/);
  assert.doesNotMatch(CSS, /invert\(/);
  // Tool rows: own classes only; beside the right Sidebar the open button drops its label, not the row's words.
  assert.match(CSS, /\.sbt\{[^}]*container-type:inline-size/);
  assert.match(CSS, /@container \(max-width:640px\)\{\.sbt-open\{padding:0 5px\}\.sbt-open span\{display:none\}\}/);
  assert.match(CSS, /\.sbp-pop\{[^}]*background:#fff;[^}]*border-radius:10px;[^}]*font:400 13px\/20px/);
});

test('P toolbar stylesheet: added once per document, removed by the last user, skipped without a DOM', () => {
  const children = [];
  const element = () => {
    const attributes = new Map();
    const node = { textContent: '', parentNode: null, getAttribute: name => attributes.get(name) ?? null, setAttribute: (name, value) => attributes.set(name, value),
      remove() { children.splice(children.indexOf(node), 1); node.parentNode = null; } };
    return node;
  };
  const head = { querySelector: () => children[0] ?? null, appendChild(node) { children.push(node); node.parentNode = head; } };
  const document = { head, createElement: element };
  const first = mountStyles(document);
  const second = mountStyles(document);
  assert.equal(children.length, 1);
  assert.equal(children[0].textContent, CSS);
  first();
  first();
  assert.equal(children.length, 1);
  second();
  assert.equal(children.length, 0);
  assert.doesNotThrow(() => mountStyles(undefined)());
});

test('toolbar Save: what the editor holds is taken first; nothing new is not written; clicks and Cmd+S share one save', async () => {
  const { h, saves, answerFlush } = await kitTab(args => ({ status: 'success', value: { saved: true, path: '/w/book.xlsx', version: `v${args.editRevision + 1}` } }));
  // Clean: the flush finds nothing new, so the file is not written and the tick shows.
  const clean = h.controller.saveNow();
  await answerFlush();
  await clean;
  assert.deepEqual([saves.length, h.controller.state.saveState], [0, 'saved']);
  const tick = h.timers.filter(timer => !timer.cleared && timer.ms === 1_200).at(-1);
  tick.fn();
  assert.equal(h.controller.state.saveState, undefined, 'the tick goes after a moment');
  // Dirty: two clicks while one save runs write once.
  h.fromFrame({ type: 'officesdk:dirty', id: 'view1', revision: 1, dirty: true, editRevision: 2, manual: true });
  const first = h.controller.saveNow();
  const second = h.controller.saveNow();
  assert.equal(first, second, 'the second click joins the running save');
  assert.equal(h.controller.state.saveState, 'saving');
  await answerFlush();
  await first;
  assert.deepEqual([saves.length, saves[0].editRevision, h.controller.state.dirty, h.controller.state.saveState], [1, 2, false, 'saved']);
});

test('toolbar Save: a failure keeps the edits and says why until the next save; an autosave already running is waited for', async () => {
  let fail = true;
  const { h, saves, answerFlush } = await kitTab(args => fail ? { status: 'failure', error: { code: 'officesdk/unavailable', message: 'disk full' } }
    : { status: 'success', value: { saved: true, path: '/w/book.xlsx', version: `v${args.editRevision + 1}` } });
  h.fromFrame({ type: 'officesdk:dirty', id: 'view1', revision: 1, dirty: true, editRevision: 2, manual: true });
  const failing = assert.rejects(h.controller.saveNow(), /disk full/);
  await answerFlush();
  await failing;
  assert.deepEqual([h.controller.state.saveState, h.controller.state.saveFailed, h.controller.state.dirty], ['failed', 'disk full', true]);
  h.controller.set({ saveFailedSeen: true });
  fail = false;
  // An autosave is writing when Save is clicked: Save waits, and with nothing new since, does not write again.
  let finishAutosave;
  h.controller.saving = new Promise(resolve => { finishAutosave = resolve; });
  const retry = h.controller.saveNow();
  assert.deepEqual([h.controller.state.saveState, h.controller.state.saveFailed, h.controller.state.saveFailedSeen], ['saving', undefined, undefined]);
  h.controller.set({ dirty: false });
  finishAutosave();
  h.controller.saving = undefined;
  // Then the editor is asked for input the autosave left alone (an IME composition, an edited cell).
  const flush = await answerFlush();
  assert.deepEqual(flush.args, {}, 'a settling flush');
  await retry;
  assert.deepEqual([saves.length, h.controller.state.saveState], [1, 'saved'], 'only the failed attempt reached the Host');
  // A failed toolbar save, then an autosave that succeeds: the failure notice goes.
  h.controller.set({ saveState: 'failed', saveFailed: 'disk full', dirty: true });
  h.fromFrame({ type: 'officesdk:dirty', id: 'view1', revision: 1, dirty: true, editRevision: 3, manual: true });
  const auto = h.controller.autosave();
  await answerFlush();
  await auto;
  assert.deepEqual([h.controller.state.saveState, h.controller.state.saveFailed, h.controller.state.dirty], [undefined, undefined, false]);
});

test('a press inside the editor page reaches the header (it closes the update bubble and menus); a stale view is not heard', async () => {
  const { h } = await kitTab(() => ({ status: 'success', value: { saved: true } }));
  let presses = 0;
  const off = h.controller.onPress(() => { presses += 1; });
  h.fromFrame({ type: 'officesdk:press', id: 'view1', revision: 1 });
  h.fromFrame({ type: 'officesdk:press', id: 'other', revision: 1 });
  assert.equal(presses, 1);
  off();
  h.fromFrame({ type: 'officesdk:press', id: 'view1', revision: 1 });
  assert.equal(presses, 1, 'unsubscribed');
});

test('PPT history: the frame announces undo/redo availability, a reopen forgets it, and toolbar undo/redo/print are agent requests', async () => {
  const { h } = await kitTab(() => ({ status: 'success', value: { saved: true } }), { format: 'pptx' });
  h.fromFrame({ type: 'officesdk:history', id: 'view1', revision: 1, canUndo: true, canRedo: false });
  assert.deepEqual(h.controller.state.history, { canUndo: true, canRedo: false });
  // The reopen below replaces the session: requests still waiting are told so.
  const pending = [h.controller.editHistory('undo'), h.controller.print()].map(request => request.catch(error => error.message));
  await tick();
  const sent = h.posted.slice(-2).map(([message]) => [message.command, message.op, message.args]);
  assert.deepEqual(sent, [['agent', 'history', { action: 'undo' }], ['agent', 'print', {}]]);
  h.controller.set({ saveState: 'failed', saveFailed: 'x', authorizePrompt: true });
  await h.controller.open(1, { mode: 'edit' });
  assert.deepEqual(h.controller.state.history, { canUndo: false, canRedo: false }, 'a reopened PPT keeps its buttons, unavailable');
  assert.deepEqual([h.controller.state.saveState, h.controller.state.saveFailed, h.controller.state.authorizePrompt], [undefined, undefined, undefined]);
  assert.deepEqual(await Promise.all(pending), ['The document session was replaced', 'The document session was replaced']);
});

test('an AI edit in flight is busy for the toolbar; its flushes are not', async () => {
  const { h } = await kitTab(() => ({ status: 'success', value: { saved: true } }), { format: 'pptx' });
  const apply = h.controller.agent('apply', { ops: [] });
  assert.equal(h.controller.state.aiBusy, true);
  void h.controller.agent('flush', { settle: false });
  const [applied] = h.posted.at(-2);
  h.fromFrame({ type: 'officesdk:agent-result', id: 'view1', revision: 1, requestId: applied.requestId, ok: true, result: {} });
  await apply;
  assert.equal(h.controller.state.aiBusy, false, 'only apply/undo count');
  assert.equal(toolbarBusy(false, { aiBusy: true }), true);
  assert.equal(toolbarBusy(false, { saveState: 'saving' }), true);
});

test('undo/redo, Save and Print buttons: PPT-only history, read mode and read-only Sessions say why', () => {
  const kit = { backend: 'kit', capabilities: { save: true }, mode: 'edit', format: 'pptx', authorized: true, history: { canUndo: true, canRedo: false } };
  const buttons = editButtons(kit);
  assert.deepEqual([buttons.undo, buttons.redo], [{ disabled: false }, { disabled: true }]);
  assert.deepEqual(buttons.save, { disabled: false, look: 'clean' });
  assert.equal(buttons.print, undefined, 'PPT and Excel print blank pages through the browser: not offered');
  assert.deepEqual(editButtons({ ...kit, format: 'docx' }).print, { disabled: false });
  assert.equal(editButtons({ ...kit, format: 'xlsx' }).print, undefined);
  // Word/Excel: no toolbar history (their editors do not expose it); a PPT without an announced history neither.
  assert.equal(editButtons({ ...kit, format: 'docx' }).undo, undefined);
  assert.equal(editButtons({ ...kit, history: undefined }).undo, undefined);
  assert.deepEqual(editButtons({ ...kit, dirty: true }).save.look, 'dirty');
  assert.deepEqual(editButtons({ ...kit, saveState: 'failed', dirty: true }).save.look, 'failed');
  assert.deepEqual(editButtons({ ...kit, saveState: 'saving' }).save, { disabled: true, look: 'saving' });
  assert.deepEqual(editButtons({ ...kit, saveState: 'saving' }).undo, { disabled: true, reason: 'busy' });
  const reading = editButtons({ ...kit, mode: 'view', history: { canUndo: false, canRedo: false } });
  assert.deepEqual([reading.undo.disabled, reading.save], [true, { disabled: true, reason: 'view', look: 'clean' }]);
  assert.deepEqual(editButtons({ ...kit, format: 'docx', mode: 'view' }).print, { disabled: false }, 'printing works in read mode');
  assert.deepEqual(editButtons({ ...kit, access: 'denied', authorized: false }).save, { disabled: true, reason: 'readOnly', look: 'clean' });
  assert.deepEqual(editButtons({ ...kit, aiBusy: true }).redo, { disabled: true, reason: 'busy' });
  // A fallback preview or a failed tab offers none of them; a Core tab saves but neither prints nor has history.
  // A fallback preview offers only Open (to leave the file); a failed tab without a lease, nothing.
  assert.deepEqual(editButtons({ ...kit, fallback: 'preview' }), { slideshow: undefined, saveAs: undefined, folder: { disabled: false }, undo: undefined, redo: undefined, save: undefined, print: undefined });
  assert.deepEqual(editButtons({ ...kit, mode: 'view' }).saveAs, { disabled: false }, 'Save As copies what read mode shows too');
  assert.deepEqual(editButtons({ ...kit, saveState: 'saving' }).folder, { disabled: true, reason: 'busy' });
  const core = editButtons({ backend: 'core', capabilities: { save: true }, mode: 'edit', format: 'docx' });
  assert.deepEqual([core.undo, core.save?.look, core.print], [undefined, 'clean', undefined]);
});

test('Save As names: a copy by default, the extension kept apart; names a file cannot take are refused in place', () => {
  assert.equal(copyName('报告.docx'), '报告 - 副本.docx');
  assert.equal(copyName('deck.v2.pptx', ' copy'), 'deck.v2 copy.pptx');
  assert.equal(nameProblem('  '), 'empty');
  assert.equal(nameProblem('a/b'), 'characters');
  assert.equal(nameProblem('what?'), 'characters');
  assert.equal(nameProblem('.hidden'), 'dot');
  assert.equal(nameProblem('季度 汇报'), undefined);
});

test('Open: only an absolute Office path from the desktop chooser is used; addresses keep an absolute path\'s leading slash', () => {
  assert.equal(pickedPath('/Users/a/b.xlsx'), '/Users/a/b.xlsx');
  assert.equal(pickedPath('relative/b.xlsx'), undefined);
  assert.equal(pickedPath('/Users/a/notes.txt'), undefined);
  assert.equal(pickedPath('/Users/a/\u0007.docx'), undefined);
  assert.equal(pickedPath(undefined), undefined);
  assert.equal(fileAddress('s1', 'docs/a b.docx'), 'dsh-resource://file/session/s1/docs/a%20b.docx');
  const outside = fileAddress('s1', '/Users/me/x#1.pptx');
  assert.equal(outside, 'dsh-resource://file/session/s1//Users/me/x%231.pptx');
  assert.deepEqual(parseSource(outside), { sessionId: 's1', path: '/Users/me/x#1.pptx' });
});

test('paths follow the Host\'s platform: POSIX without one; on Windows drives, shares and \\ too', () => {
  for (const platform of [undefined, 'darwin', 'linux', 'nonsense']) {
    assert.deepEqual(['/Users/a.docx', 'a.docx', 'C:\\a.docx', 'C:/a.docx', '\\\\srv\\share\\a.docx'].map(path => isAbsolute(path, platform)), [true, false, false, false, false]);
    assert.deepEqual([parentOf('/a.docx', platform), parentOf('a.docx', platform), parentOf('sub/a.docx', platform), parentOf('/x/y/a.docx', platform), parentOf('sub\\a.docx', platform)], ['/', '', 'sub', '/x/y', '']);
    assert.deepEqual([baseName('/x/a.docx', platform), baseName('sub\\a.docx', platform)], ['a.docx', 'sub\\a.docx']);
    assert.deepEqual([joinPath('', 'a.docx', platform), joinPath('/', 'a.docx', platform), joinPath('/x/', 'a.docx', platform), joinPath('sub', 'a.docx', platform)], ['a.docx', '/a.docx', '/x/a.docx', 'sub/a.docx']);
    assert.deepEqual(segments('/x//y/a.docx', platform), ['x', 'y', 'a.docx']);
  }
  const win = 'win32';
  assert.deepEqual(['C:\\a.docx', 'c:/a.docx', '\\\\srv\\share\\a.docx', '/a.docx', 'a.docx', 'sub\\a.docx', 'C:a.docx'].map(path => isAbsolute(path, win)), [true, true, true, true, false, false, false]);
  assert.deepEqual([rootOf('C:\\x\\a.docx', win), rootOf('\\\\srv\\share\\a.docx', win), rootOf('sub\\a.docx', win)], ['C:\\', '\\\\srv\\share\\', '']);
  assert.deepEqual([parentOf('C:\\a.docx', win), parentOf('C:\\x\\y\\a.docx', win), parentOf('\\\\srv\\share\\a.docx', win), parentOf('sub\\a.docx', win), parentOf('sub/deep/a.docx', win), parentOf('a.docx', win)],
    ['C:\\', 'C:\\x\\y', '\\\\srv\\share\\', 'sub', 'sub/deep', '']);
  assert.deepEqual([baseName('C:\\x\\a.docx', win), baseName('sub/a.docx', win)], ['a.docx', 'a.docx']);
  assert.deepEqual([joinPath('C:\\', 'a.docx', win), joinPath('C:\\x\\', 'a.docx', win), joinPath('C:\\x', 'a.docx', win), joinPath('sub/deep', 'a.docx', win), joinPath('\\\\srv\\share', 'a.docx', win)],
    ['C:\\a.docx', 'C:\\x\\a.docx', 'C:\\x\\a.docx', 'sub/deep/a.docx', '\\\\srv\\share\\a.docx']);
  assert.deepEqual([segments('C:\\x/y\\a.docx', win), segments('C:\\', win).at(-1)], [['C:', 'x', 'y', 'a.docx'], 'C:']);
});

test('Windows: the chooser\'s drive and share paths open; addresses take \\ as a separator there and nowhere else', () => {
  assert.equal(pickedPath('C:\\Users\\me\\b.xlsx', 'win32'), 'C:\\Users\\me\\b.xlsx');
  assert.equal(pickedPath('\\\\srv\\share\\b.pptx', 'win32'), '\\\\srv\\share\\b.pptx');
  assert.equal(pickedPath('C:\\Users\\me\\notes.txt', 'win32'), undefined);
  assert.equal(pickedPath('relative\\b.xlsx', 'win32'), undefined);
  assert.equal(pickedPath('C:\\Users\\me\\b.xlsx'), undefined, 'without the Host saying Windows, only POSIX paths count');
  assert.equal(pickedPath('/Users/a/b.xlsx', 'darwin'), '/Users/a/b.xlsx');
  assert.equal(fileAddress('s1', 'docs\\a b.docx', 'win32'), fileAddress('s1', 'docs/a b.docx', 'win32'));
  assert.equal(fileAddress('s1', 'docs/a b.docx', 'win32'), 'dsh-resource://file/session/s1/docs/a%20b.docx');
  assert.equal(fileAddress('s1', 'C:\\Users\\me\\x#1.pptx', 'win32'), 'dsh-resource://file/session/s1/C%3A/Users/me/x%231.pptx');
  assert.equal(fileAddress('s1', 'docs\\a.docx'), 'dsh-resource://file/session/s1/docs%5Ca.docx', 'on POSIX a backslash is part of a name');
  assert.equal(fileAddress('s1', 'docs\\a.docx', 'darwin'), fileAddress('s1', 'docs\\a.docx'));
});

test('AI commands address tabs as the tabs do: one implementation, the Host\'s platform from a lease or the command', async () => {
  assert.equal(sessionAddress, fileAddress);
  assert.equal(sessionAddress('s 1', '/abs/x.pptx'), 'dsh-resource://file/session/s%201//abs/x.pptx');
  const registry = new ControllerRegistry();
  let platform;
  const bridge = new AgentBridge({ rpc: null, openResource: () => {}, controllers: registry, mounted: null, platform: () => platform });
  assert.equal(bridge.address('s1', 'sub\\deck.pptx'), 'dsh-resource://file/session/s1/sub%5Cdeck.pptx');
  platform = 'win32';
  assert.equal(bridge.address('s1', 'sub\\deck.pptx'), fileAddress('s1', 'sub/deck.pptx'));
  assert.equal(bridge.address('s1', 'sub\\deck.pptx', { platform: 'darwin' }), 'dsh-resource://file/session/s1/sub%5Cdeck.pptx', 'a command that names the platform wins');
  const marked = [];
  registry.add(fileAddress('s1', 'sub/deck.pptx'), { markSaved: (...args) => marked.push(args), state: {} });
  await bridge.serve('s1', { kind: 'saved', path: 'sub\\deck.pptx', version: 'v2', editRevision: 3 });
  assert.deepEqual(marked, [['v2', 3]]);
  // A bare value works as well as a getter.
  assert.equal(new AgentBridge({ platform: 'win32' }).address('s1', 'a\\b.pptx'), fileAddress('s1', 'a/b.pptx'));
});

test('the lease says the Host\'s platform; the tab keeps it after the lease is gone', async () => {
  const h = harness({ open: args => ({ id: 'view1', backend: 'kit', format: 'docx', revision: args.revision, version: 'v1', editRevision: 0, mode: 'edit', platform: 'win32',
    entry: '/shimo/kit/static/page/index.html?view=view1', expiresInMs: 900_000, capabilities: { edit: true, save: true, export: true } }) });
  assert.equal(h.controller.platform, undefined, 'nothing is guessed before the Host answers');
  await h.controller.open(1, { mode: 'edit' });
  assert.equal(h.controller.platform, 'win32');
  h.tab.abort();
  assert.deepEqual([h.controller.lease, h.controller.platform], [undefined, 'win32']);
  const mac = harness();
  await mac.controller.open(1, { mode: 'edit' });
  assert.equal(mac.controller.platform, undefined, 'a lease without one leaves POSIX rules');
});

test('OPEN-01 last save: saved only with nothing new; the editor is locked and autosave held throughout, released on every way out', async () => {
  assert.equal(savedForSwitch({ dirty: false }, 3, 3), true);
  assert.equal(savedForSwitch({ dirty: false }, 3, 4), false);
  assert.equal(savedForSwitch({ unstaged: 'x' }, 3, 3), false);
  const { h } = await kitTab(() => ({ status: 'success', value: { saved: true } }));
  const states = [];
  let rounds = 0;
  // Every save "succeeds", yet a new edit lands each time: three tries, then the tab stays and says so.
  h.controller.saveNow = async () => { states.push([h.controller.state.locked, h.controller.autosavePaused]); rounds += 1; h.controller.editRevision += 1; };
  assert.equal(await h.controller.saveForSwitch(), false);
  assert.equal(rounds, 3);
  assert.deepEqual(states, [[true, true], [true, true], [true, true]]);
  assert.deepEqual([h.controller.state.locked, h.controller.autosavePaused, h.controller.state.saveState], [undefined, false, 'failed']);
  // A failure releases them too.
  h.controller.saveNow = async () => { throw Object.assign(new Error('disk full'), { code: 'officesdk/unavailable' }); };
  await assert.rejects(h.controller.saveForSwitch(), /disk full/);
  assert.deepEqual([h.controller.state.locked, h.controller.autosavePaused], [undefined, false]);
  // Nothing new: saved at once.
  h.controller.saveNow = async () => {};
  assert.equal(await h.controller.saveForSwitch(), true);
});

test('while the Host answers the access question, Save As, the account hold or a reopen still stop the autosave; a stalled answer lets it go on; only the latest answer counts', async () => {
  const held = [];
  const deferred = () => ({ answer: () => new Promise(resolve => held.push(resolve)) });
  for (const begin of [h => h.controller.pauseAutosave(true), h => h.controller.setAccountHold(true), h => { h.controller.lease = { ...h.controller.lease, id: 'view2' }; }]) {
    held.length = 0;
    const tab = await accessTab(saved, deferred());
    edit(tab.h, 1);
    tab.pause().fn(); await tick();
    begin(tab.h);
    held.at(-1)({ status: 'success', value: { access: 'allowed' } }); await tick(); await tick();
    assert.deepEqual([tab.flushes(), tab.saves.length, tab.h.controller.state.autosaveFailed], [0, 0, undefined], 'nothing handed over, written or reported');
  }

  held.length = 0;
  const stalled = await accessTab(saved, deferred());
  edit(stalled.h, 1);
  stalled.pause().fn(); await tick();
  stalled.h.timers.filter(t => !t.cleared && t.ms === 4_000).at(-1).fn();
  await stalled.answerFlush();
  assert.deepEqual([stalled.saves.length, stalled.h.controller.state.dirty], [1, false], 'a Host that does not answer in time does not hold the save');

  held.length = 0;
  const order = await accessTab(saved, deferred());
  order.h.timers.filter(t => !t.cleared && t.ms === 3_000).at(-1).fn(); await tick();
  void order.h.controller.refreshAccess(); await tick();
  held[1]({ status: 'success', value: { access: 'denied' } }); await tick(); await tick();
  held[0]({ status: 'success', value: { access: 'allowed' } }); await tick(); await tick();
  assert.deepEqual([order.h.controller.state.access, order.h.controller.state.authorized], ['denied', false], 'an older answer arriving late does not undo a newer one');
});

test('toolbar Save after the Session turned read only, before the next poll: the tab turns read only and says so, not the Host\'s refusal', async () => {
  const { h, now } = await accessTab(() => failure('sandbox-denied'));
  h.controller.messages = { readOnlyHint: '当前会话是只读的，这里的修改不会保存到原文件。' };
  now.access = 'denied';
  edit(h, 1);
  const saving = h.controller.saveNow();
  await tick();
  const [flush] = h.posted.at(-1);
  h.fromFrame({ type: 'officesdk:agent-result', id: 'view1', revision: 1, requestId: flush.requestId, ok: true, result: { editRevision: 1 } });
  await assert.rejects(saving, error => error.code === 'officesdk/sandbox-denied');
  assert.deepEqual([h.controller.state.access, h.controller.state.authorized, h.controller.state.saveState, h.controller.state.saveFailed],
    ['denied', false, 'failed', '当前会话是只读的，这里的修改不会保存到原文件。']);
});

test('holding autosave for Save As never lifts a lasting stop; Save As and probes ask the Host with their options', async () => {
  const { h, pause } = await kitTab(() => ({ status: 'success', value: { saved: true } }));
  h.controller.pauseAutosave(true);
  h.fromFrame({ type: 'officesdk:dirty', id: 'view1', revision: 1, dirty: true, editRevision: 2, manual: true });
  assert.equal(pause(), undefined, 'held: no autosave scheduled');
  h.controller.pauseAutosave(false);
  assert.ok(pause(), 'resumed: the pending edits autosave at the next pause');
  h.controller.readOnly();
  h.controller.pauseAutosave(true);
  h.controller.pauseAutosave(false);
  assert.deepEqual([h.controller.state.access, h.controller.state.authorized], ['denied', false], 'a read-only stop stays');
  void h.controller.probe('/Users/me/a.docx');
  void h.controller.browse('docs');
  await tick();
  const asked = h.calls.slice(-2).map(([name, args]) => [name, args.path, args.kind]);
  assert.deepEqual(asked, [['probe', '/Users/me/a.docx', 'file'], ['browse', 'docs', undefined]]);
});

test('OPEN-01 plan: no editing session switches at once; Kit without write access asks (or offers Save As only when read-only); otherwise a real save first', () => {
  assert.equal(switchPlan({ mode: 'view' }, true), 'switch');
  assert.equal(switchPlan({ mode: 'edit', dirty: true }, false), 'switch', 'a failed or released tab has nothing to save');
  const kit = { backend: 'kit', mode: 'edit', authorized: false, access: 'outside' };
  assert.equal(switchPlan(kit, true), 'switch', 'nothing to write');
  assert.equal(switchPlan({ ...kit, dirty: true }, true), 'authorize');
  assert.equal(switchPlan({ ...kit, unstaged: 'x' }, true), 'authorize');
  assert.equal(switchPlan({ ...kit, dirty: true, access: 'denied' }, true), 'readOnly');
  assert.equal(switchPlan({ ...kit, authorized: true }, true), 'save', 'a clean-looking tab still hands over its input');
  assert.equal(switchPlan({ backend: 'core', mode: 'edit' }, true), 'save');
});

test('a pending switch belongs to the view it was asked from: a reopen drops it; the own-file lookup retries after a failure', async () => {
  let fail = true;
  const { h } = await kitTab(() => ({ status: 'success', value: { saved: true } }));
  h.controller.call = async (method, args) => {
    if (method === 'probe') { if (fail) throw new Error('busy'); return { realPath: '/w/book.xlsx', path: 'book.xlsx' }; }
    return { status: 'success', value: {} };
  };
  assert.equal(await h.controller.ownRealPath(), undefined);
  fail = false;
  assert.equal(await h.controller.ownRealPath(), '/w/book.xlsx', 'a failure is not remembered');
  h.controller.set({ pendingOpen: { address: 'dsh-resource://file/session/s1/other.docx', source: h.controller.address, tabId: 't', awaiting: true } });
  delete h.controller.call;
  await h.controller.open(1, { mode: 'edit' });
  assert.equal(h.controller.state.pendingOpen, undefined);
});

test('account rules: not connected stops editing and saving whatever the credits; 0 credits still saves; unknown usage is not 0', () => {
  assert.deepEqual(accountGate({ status: 'connected', credits: 670 }), { edit: true, save: true, paid: 'yes' });
  assert.deepEqual(accountGate({ status: 'connected', credits: 0 }), { edit: true, save: true, paid: 'no' });
  assert.deepEqual(accountGate({ status: 'connected', credits: null }), { edit: true, save: true, paid: 'unknown' });
  assert.deepEqual(accountGate({ status: 'disconnected', credits: 0 }), { edit: false, save: false, paid: 'no', reason: 'account' });
  assert.deepEqual(accountGate({ status: 'expired', credits: 500 }).reason, 'account');
});

test('account mock: parsed defensively; the shared source tells this window\'s tabs at once and other windows through storage', () => {
  const warnings = [];
  assert.deepEqual(parseAccount(null), { status: 'connected', email: 'user@example.com', credits: 670 });
  assert.equal(parseAccount('{oops', text => warnings.push(text)).status, 'connected');
  assert.equal(warnings.length, 1);
  assert.deepEqual(parseAccount('{"status":"weird","credits":-3}'), { status: 'disconnected', email: undefined, credits: null });
  assert.equal(parseAccount('{"status":"connected","credits":"5"}').credits, null);
  const store = new Map();
  const storage = { getItem: key => store.get(key) ?? null, setItem: (key, value) => store.set(key, value) };
  const source = createAccountSource(() => storage);
  assert.equal(accountHeld(source.snapshot()), false, 'preview off: no hold');
  const seen = [];
  source.subscribe(snapshot => seen.push([snapshot.preview, snapshot.account.status]));
  store.set(PREVIEW_KEY, '1');
  store.set(MOCK_KEY, JSON.stringify({ status: 'expired', email: 'a@b.c', credits: 3 }));
  source.reload();
  assert.equal(accountHeld(source.snapshot()), true);
  source.connect();
  assert.equal(JSON.parse(store.get(MOCK_KEY)).status, 'connected', 'connect is written for other windows');
  assert.deepEqual(seen, [[true, 'expired'], [true, 'connected']]);
  assert.equal(accountHeld(source.snapshot()), false);
  const refusing = createAccountSource(() => { throw new Error('denied'); });
  assert.deepEqual([refusing.snapshot().preview, refusing.snapshot().account.status], [false, 'connected']);
});

test('account hold on the toolbar: write buttons off with the reason, the eye only while there are edits; nothing changes without a hold', () => {
  const state = { backend: 'kit', capabilities: { save: true }, mode: 'edit', format: 'pptx', authorized: true, history: { canUndo: true, canRedo: true } };
  const edits = editButtons(state);
  const modes = modeButtons(state);
  const same = heldButtons(edits, modes, state, false);
  assert.equal(same.edits, edits);
  assert.equal(same.modes, modes);
  const clean = heldButtons(edits, modes, state, true);
  assert.deepEqual([clean.edits.undo.reason, clean.edits.save.reason, clean.edits.saveAs.reason], ['account', 'account', 'account']);
  assert.deepEqual(clean.edits.folder, edits.folder, 'Open stays');
  assert.deepEqual(clean.modes.view, modes.view, 'a clean tab may still go to read mode');
  assert.equal(clean.modes.edit, modes.edit, 'the pressed pen stays as it is');
  const dirty = heldButtons(editButtons({ ...state, dirty: true }), modeButtons({ ...state, dirty: true }), { ...state, dirty: true }, true);
  assert.deepEqual(dirty.modes.view, { pressed: false, disabled: true, reason: 'account' });
  const reading = heldButtons(edits, modeButtons({ ...state, mode: 'view' }), { ...state, mode: 'view' }, true);
  assert.deepEqual(reading.modes.edit, { pressed: false, disabled: true, reason: 'account' });
  // Open: a tab with edits stays; a clean one goes. Without a hold the plan is what it was.
  assert.equal(switchPlan({ ...state, dirty: true }, true, true), 'account');
  assert.equal(switchPlan(state, true, true), 'switch');
  assert.equal(switchPlan({ ...state, dirty: true }, true), 'save');
  assert.deepEqual(accountMenuKey('Enter', 1, 3), { index: 1, activate: true });
  assert.deepEqual(accountMenuKey('ArrowUp', 0, 3), { index: 2 });
  assert.deepEqual(accountMenuKey('Escape', 0, 3), { index: 0, close: 'restore' });
});

test('account hold in the controller: nothing is written — no save, autosave, authorization, AI edit or autosave on release; lifting it saves at the next pause', async () => {
  const { h, saves, pause } = await kitTab(() => ({ status: 'success', value: { saved: true } }));
  h.controller.setAccountHold(true);
  assert.equal(h.controller.state.accountHold, true);
  h.fromFrame({ type: 'officesdk:dirty', id: 'view1', revision: 1, dirty: true, editRevision: 2, manual: true });
  assert.equal(pause(), undefined, 'no autosave while held');
  await assert.rejects(h.controller.saveNow(), { code: 'officesdk/account' });
  await assert.rejects(h.controller.save(), { code: 'officesdk/account' });
  await assert.rejects(h.controller.authorize(), { code: 'officesdk/account' });
  await assert.rejects(h.controller.agent('apply', { ops: [] }), { code: 'officesdk/account' });
  await assert.rejects(h.controller.exportCopy('copy.xlsx'), { code: 'officesdk/account' }, 'Save As is a save too');
  await assert.rejects(h.controller.ensureMode('edit'), { code: 'officesdk/account' }, 'an AI edit is refused before any reopen');
  // Save As / a switch's last save end their own pause: the hold stays.
  h.controller.pauseAutosave(true);
  h.controller.pauseAutosave(false);
  assert.equal(pause(), undefined);
  assert.equal(saves.length, 0);
  // Lifting the hold: the edits (still there) autosave at the next pause.
  h.controller.setAccountHold(false);
  assert.ok(pause(), 'autosave resumes');
  assert.equal(h.controller.state.dirty, true, 'the edits are still there');
  // Held again: closing the tab releases without autosaving, and a reopen is read-only.
  h.controller.setAccountHold(true);
  await h.controller.open(1, { mode: 'edit' });
  const releases = h.calls.filter(([name]) => name === 'release');
  assert.equal(releases.at(-1)[1].autosave, false);
  assert.equal(h.calls.filter(([name]) => name === 'open').at(-1)[1].mode, 'view');
});

test('the editor frame\'s sandbox comes from the address, before any lease: print needs allow-modals from the first load', async () => {
  const { sandboxFormat } = await import('../src/client/plugin.mjs');
  assert.equal(sandboxFormat('dsh-resource://file/session/s1/docs/Report.DOCX'), 'docx');
  assert.equal(sandboxFormat('dsh-resource://file/session/s1//Users/me/deck.pptx'), 'pptx');
  assert.equal(sandboxFormat('not an address'), undefined);
});

test('the three-dot checkbox turns the account preview on and off; the menu\'s demo states write the mock and take effect at once', () => {
  const store = new Map();
  const storage = { getItem: key => store.get(key) ?? null, setItem: (key, value) => store.set(key, value), removeItem: key => store.delete(key) };
  const source = createAccountSource(() => storage);
  const seen = [];
  source.subscribe(snapshot => seen.push(`${snapshot.preview}:${demoState(snapshot.account)}`));
  source.setPreview(true);
  assert.equal(store.get(PREVIEW_KEY), '1');
  for (const state of DEMO_STATES) source.setMock(state.account);
  assert.deepEqual(seen, ['true:connected', 'true:connected', 'true:disconnected', 'true:expired', 'true:zero', 'true:unknown']);
  assert.equal(accountHeld(source.snapshot()), false, 'usage unavailable is not a hold');
  source.setMock(DEMO_STATES[1].account);
  assert.equal(accountHeld(source.snapshot()), true);
  source.setPreview(false);
  assert.equal(store.has(PREVIEW_KEY), false);
  assert.equal(accountHeld(source.snapshot()), false, 'off: no hold, whatever the mock');
});

test('release check: one for all open tabs, again every ten minutes while any is open; a failure keeps the last answer; the last tab stops it', async () => {
  const timers = [];
  const scheduler = { setTimeout: (fn, ms) => { timers.push({ fn, ms }); return timers.length; }, clearTimeout: id => { if (timers[id - 1]) timers[id - 1].cleared = true; } };
  const asked = [];
  let answer = { current: '0.0.33', latest: '0.0.34', available: true };
  const watch = createUpdateWatch({ scheduler, fetch: async ({ refresh }) => { asked.push(refresh); if (answer instanceof Error) throw answer; return answer; } });
  const offA = watch.subscribe(() => {});
  const offB = watch.subscribe(() => {});
  await tick();
  assert.deepEqual(asked, [false], 'the first tab checks (the Host may answer from its ten-minute cache); a second tab shares it');
  assert.equal(watch.snapshot().latest, '0.0.34');
  assert.deepEqual(timers.filter(timer => !timer.cleared).map(timer => timer.ms), [UPDATE_CHECK_MS]);
  answer = new Error('offline');
  timers.at(-1).fn();
  await tick();
  assert.deepEqual(asked, [false, true], 'the background check asks the branch again');
  assert.equal(watch.snapshot().latest, '0.0.34', 'a failed check is not "up to date"');
  answer = { current: '0.0.33', latest: '0.0.35', available: true };
  await watch.check(true);
  assert.equal(watch.snapshot().latest, '0.0.35');
  offA();
  assert.equal(timers.at(-1).cleared, undefined, 'one tab is still open');
  offB();
  assert.equal(timers.at(-1).cleared, true, 'no tab: no more checks');
  assert.equal(UPDATE_CHECK_MS, 600_000);
});

test('selection-update from the page reaches the Host with the selection id, not the frame envelope id', async () => {
  const { createFrameChannel } = await import('../../core/src/browser/frame.mjs');
  const h = harness();
  await h.controller.open(1, { mode: 'edit' });
  h.fromFrame({ type: 'officesdk:ready', id: 'lease1', revision: 1 });
  // The page's own envelope: `id` is the view, so the selection id must travel beside it.
  let sent;
  const page = { parent: { postMessage: message => { sent = message; } }, location: { origin: 'http://app.local' }, addEventListener() {} };
  createFrameChannel({ window: page, id: 'lease1', revision: 1, parentOrigin: 'http://app.local' }).notify('selection-update', { selection: 'sel-0a0b0c0d', state: 'cancelled' });
  assert.equal(sent.id, 'lease1');
  h.fromFrame(sent);
  await new Promise(resolve => setImmediate(resolve));
  assert.deepEqual(h.calls.filter(([name]) => name === 'selectionUpdate'), [['selectionUpdate', { workspaceFileScopeId: 's1', id: 'sel-0a0b0c0d', selectionState: 'cancelled' }]]);
  // Anything that is not a selection id is ignored.
  createFrameChannel({ window: page, id: 'lease1', revision: 1, parentOrigin: 'http://app.local' }).notify('selection-update', { selection: 'lease1', state: 'cancelled' });
  h.fromFrame(sent);
  await new Promise(resolve => setImmediate(resolve));
  assert.equal(h.calls.filter(([name]) => name === 'selectionUpdate').length, 1);
});

test('while open waits on the document engine download the tab shows its progress and the rendering timeout restarts', async () => {
  let finish;
  const opening = new Promise(resolve => { finish = resolve; });
  let status = { state: 'downloading', received: 10 * 1048576, total: 100 * 1048576 };
  const h = harness({ host: { open: () => opening.then(lease => ({ status: 'success', value: lease })), engineStatus: () => ({ status: 'success', value: status }) } });
  const pending = h.controller.open(1);
  await tick();
  const watch = h.timers.find(timer => timer.ms === 800);
  assert.ok(watch, 'the engine watch starts 800 ms after open');
  watch.fn();
  await tick(); await tick();
  assert.deepEqual(h.controller.state.engine, { received: 10 * 1048576, total: 100 * 1048576 });
  assert.ok(h.timers.filter(timer => timer.ms === 120_000).length >= 2, 'the rendering timeout was started again');
  assert.equal(h.timers.at(-1).ms, 1000, 'it keeps watching every second');
  // The editor fell back from npm to a later source (a drill profile); this platform's engine is in test (spec 261111-kit-0-1-3).
  status = { state: 'downloading', received: 20 * 1048576, total: 100 * 1048576, source: 'release', fallback: ['npm'], platformStatus: 'test' };
  h.timers.at(-1).fn();
  await tick(); await tick();
  assert.deepEqual(h.controller.state.engine, { received: 20 * 1048576, total: 100 * 1048576, fallback: true, testing: true });
  status = { state: 'ready' };
  h.timers.at(-1).fn();
  await tick(); await tick();
  assert.equal(h.controller.state.engine, undefined);
  finish({ id: 'lease1', format: 'docx', revision: 1, version: 'v1', editRevision: 0, mode: 'edit', entry: '/shimo/doc/lease1/index.html', expiresInMs: 900_000, capabilities: { edit: true } });
  await pending;
  assert.equal(h.controller.lease.id, 'lease1');
});

test('project table commit defers file revision notifications until the refreshed editor is ready', async () => {
  const h=harness({open:args=>({id:'view1',backend:'kit',format:'docx',revision:args.revision,version:'v1',editRevision:0,mode:'edit',entry:'/shimo/kit/static/page/index.html?view=view1',expiresInMs:900000})});
  await h.controller.open(1,{mode:'edit'});h.fromFrame({type:'officesdk:ready',id:'view1',revision:1});
  h.controller.beginTableCommit();await h.controller.adopt(2,'v2',{});await h.controller.adopt(3,'v3',{});
  assert.equal(h.calls.filter(([name])=>name==='rebind').length,0);assert.equal(h.controller.revision,1);
  h.controller.endTableCommit();await tick();await tick();
  assert.equal(h.controller.revision,3);assert.equal(h.calls.filter(([name])=>name==='rebind').length,1);assert.equal(h.controller.autosavePaused,false);h.tab.abort();
});

test('an incoming PPT revision waits for the opening frame to acknowledge its original revision', async () => {
  const h = harness({open:args=>({id:'deck',revision:args.revision,version:'v1',backend:'kit',format:'pptx',mode:'edit',entry:'/shimo/kit/static/page/index.html?view=deck',capabilities:{},expiresInMs:900000}),
    host:{rebind:args=>({status:'success',value:{revision:args.revision,state:'active',version:'v2'}})}});
  await h.controller.open(1,{mode:'edit'});
  await h.controller.adopt(2,'v2',{});
  assert.equal(h.controller.revision,1);
  h.fromFrame({type:'officesdk:ready',id:'deck',revision:1});
  await tick();
  assert.equal(h.controller.revision,2);
  assert.equal(h.controller.state.ready,true);
  assert.equal(h.names().filter(n=>n==='open').length,1);
  h.tab.abort();
});

test('the first agent joins the PPT tab while the initial open RPC is outstanding',async()=>{
  let grant;
  const h=harness({host:{open:()=>new Promise(resolve=>{grant=resolve;})}});
  const opening=h.controller.open(1,{mode:'edit'});
  const joining=h.controller.ensureMode('edit',new AbortController().signal);
  await tick();
  assert.deepEqual(h.names(),['open']);
  grant({status:'success',value:{id:'only',revision:1,version:'v1',backend:'kit',format:'pptx',mode:'edit',entry:'/shimo/kit/static/page/index.html?view=only',capabilities:{},expiresInMs:900000}});
  await opening;
  h.fromFrame({type:'officesdk:ready',id:'only',revision:1});
  await joining;
  h.tab.abort();
});
