import test from 'node:test';
import assert from 'node:assert/strict';
import { attachmentKind, createChangeQueue, fitWriterPageWidth, pageLocale, startKitPage, suiteMountOptions, ACT_EVENTS } from '../src/browser/kit-page.mjs';
import { injectWorkerSource, installSheetMediaBridge, sheetMediaReference } from '../src/browser/kit-sheet-assets.mjs';
import { frameCommand } from '../src/browser/frame.mjs';

const user = { uid: 'local:me', name: 'me' };

test('mount options per suite: sheet is zh-CN with its own React and no revision; Writer and Sheet name the user', () => {
  const common = { mode: 'edit', user, language: 'en-US', nextChangeSetRevision: 4 };
  // Sheet always turns its new/old UI switch off and passes no feedbackUrl (no feedback entry).
  const ui = { versionSwitch: { disabled: true } };
  assert.deepEqual(suiteMountOptions({ suite: 'sheet', ...common }), { language: 'zh-CN', user, ui });
  const document = suiteMountOptions({ suite: 'document', ...common });
  assert.deepEqual([document.language, document.nextChangeSetRevision, document.user, document.runtimeScripts.length], ['en-US', 4, user, 3]);
  const presentation = suiteMountOptions({ suite: 'presentation', ...common });
  assert.deepEqual([presentation.nextChangeSetRevision, presentation.user], [4, undefined]);
  for (const suite of ['document', 'presentation']) for (const mode of ['edit', 'view']) assert.equal('ui' in suiteMountOptions({ suite, ...common, mode }), false, `${suite} ${mode}: office-kit refuses ui outside Sheet`);
  assert.equal(suiteMountOptions({ suite: 'document', ...common, mode: 'view' }).nextChangeSetRevision, undefined, 'read only: no change log to continue');
  assert.deepEqual(suiteMountOptions({ suite: 'sheet', ...common, mode: 'view' }), { language: 'zh-CN', ui }, 'read-only Sheet refuses edit options such as user, but still turns the UI switch off');
  assert.deepEqual(suiteMountOptions({ suite: 'document', ...common, mode: 'view' }).user, user);
});

test('inserted media kinds per suite', () => {
  const editor = { presentationResourceCategory: (bucket) => bucket === 'images' ? 'media' : 'embeddings' };
  assert.deepEqual([attachmentKind(editor, 'presentation', 'images'), attachmentKind(editor, 'presentation', 'files')], ['media', 'embeddings']);
  assert.deepEqual([attachmentKind(editor, 'document', 'images'), attachmentKind(editor, 'document', 'files')], ['image', 'embedding']);
  assert.equal(attachmentKind(editor, 'sheet', 'images'), undefined);
});

test('sheet media references: only modoc-assets media names, normalized to Kit\'s form', () => {
  assert.equal(sheetMediaReference('modoc-assets:/media/a1.png'), 'modoc-assets:/media/a1.png');
  assert.equal(sheetMediaReference('modoc-assets://media/a1.png?x=1'), 'modoc-assets:/media/a1.png');
  for (const bad of ['modoc-assets:/media/../x', 'modoc-assets:/media/.env', 'modoc-assets:/other/a.png', 'https://x/media/a.png', 'modoc-assets:/media/a%2Fb.png', undefined]) assert.equal(sheetMediaReference(bad), undefined, String(bad));
});

test('only the Sheet image worker gets the shim, once', () => {
  const worker = 'self.addEventListener("message", e => fetch(e.data.requestKey))';
  const shimmed = injectWorkerSource(worker, 'application/javascript');
  assert.notEqual(shimmed, worker);
  assert.equal(injectWorkerSource(shimmed, 'application/javascript'), shimmed);
  assert.equal(injectWorkerSource('postMessage(1)', 'application/javascript'), 'postMessage(1)');
  assert.equal(injectWorkerSource(worker, 'text/css'), worker);
});

test('the page answers modoc-assets fetches through the Host and leaves every other request alone', async () => {
  const reads = [], passed = [];
  const window = { URL, Headers, Response, Blob, Worker: class {}, fetch: async input => { passed.push(String(input)); return new Response('net'); } };
  const uninstall = installSheetMediaBridge(window, async (reference, purpose) => { reads.push([reference, purpose]); return { ok: true, status: 200, contentType: 'image/png', bytes: new Uint8Array([1, 2]).buffer }; });
  const picture = await window.fetch('modoc-assets://media/p.png');
  assert.deepEqual([picture.status, picture.headers.get('content-type'), [...new Uint8Array(await picture.arrayBuffer())]], [200, 'image/png', [1, 2]]);
  assert.deepEqual(reads, [['modoc-assets:/media/p.png', 'display']]);
  assert.equal((await window.fetch('modoc-assets:/media/../etc')).status, 404, 'an invalid reference never reaches the Host');
  assert.equal(await (await window.fetch('/shimo/kit/v/x/meta')).text(), 'net');
  assert.deepEqual(passed, ['/shimo/kit/v/x/meta']);
  uninstall();
  assert.equal(await (await window.fetch('modoc-assets:/media/p.png')).text(), 'net');
});

test('the change queue still numbers changes from 0 after a restart', async () => {
  const sent = [];
  const queue = createChangeQueue({ send: async (seq, change) => { sent.push([seq, change.format]); } });
  queue.push({ format: 'mow-v1', body: 'a' }); queue.push({ format: 'mow-v1', body: 'b' });
  await queue.idle();
  queue.restart();
  queue.push({ format: 'mow-v1', body: 'c' });
  await queue.idle();
  assert.deepEqual(sent, [[0, 'mow-v1'], [1, 'mow-v1'], [0, 'mow-v1']]);
});

test('page width: the SDK\'s own status-bar command is used; without it the document keeps its default zoom', async () => {
  const clicks = [];
  const zoom = { click: () => clicks.push('zoom') }, item = { click: () => clicks.push('fit') };
  const document = { querySelector: selector => selector === '.editor-statusbar__zoom-value' ? zoom : selector.includes('root:fit-page-width') && clicks.includes('zoom') ? item : null };
  assert.equal(await fitWriterPageWidth(document, { wait: async () => {} }), true);
  assert.deepEqual(clicks, ['zoom', 'fit']);
  const warn = console.warn; console.warn = () => {};
  try { assert.equal(await fitWriterPageWidth({ querySelector: () => null }, { timeoutMs: 300, wait: async () => {} }), false); }
  finally { console.warn = warn; }
});

/** A Kit page in a fake frame: the parent sees `posted`, the Host sees `requests`. */
function fakePage({ meta = { v: 1, id: 'v1', suite: 'document', mode: 'edit', name: 'a.docx' }, abandon = { closed: true, unsaved: true } } = {}) {
  const posted = [], requests = [];
  const parent = { postMessage: message => posted.push(message) };
  const window = { location: { href: 'http://app.local/shimo/kit/static/page/index.html?view=v1&revision=1&parentOrigin=http%3A%2F%2Fapp.local', origin: 'http://app.local' },
    parent, addEventListener() {}, removeEventListener() {} };
  const root = { replaceChildren() { requests.push('replaceChildren'); } };
  const document = { getElementById: () => root, documentElement: { lang: 'zh' }, head: { append() {} }, createElement: () => ({}) };
  const json = body => ({ ok: true, status: 200, headers: new Headers({ 'x-next-change-set-revision': '0' }), json: async () => body, text: async () => JSON.stringify(body) });
  const fetch = async (url, init = {}) => {
    const where = new URL(url).pathname;
    requests.push(`${init.method ?? 'GET'} ${where}`);
    if (where.endsWith('/meta')) return json(meta);
    if (where.endsWith('/catalog.json')) return json({});
    if (where.endsWith('/content')) return json({ blocks: [] });
    if (where.endsWith('/abandon')) return json(abandon);
    return json({});
  };
  return { window, document, fetch, posted, requests };
}

test('a page that cannot load its editor abandons the view, then reports failed with the unsaved flag', async () => {
  const page = fakePage();
  await assert.rejects(startKitPage({ ...page, loadEditor: async () => { throw new Error('bundle missing'); } }), /bundle missing/);
  assert.ok(page.requests.includes('POST /shimo/kit/v/v1/abandon'));
  const failed = page.posted.find(message => message.type === 'officesdk:failed');
  assert.deepEqual([failed.message, failed.unsaved, failed.rebuild], ['bundle missing', true, undefined]);
});

test('a change fault while the editor mounts fails the mount once and never starts a rebuild', async () => {
  const page = fakePage();
  const listeners = {};
  const editorModule = {
    parseResourceCatalog: value => value, parseSystemFontCatalog: value => value, presentationResourceCategory: () => 'media',
    createOfflineEditor: async () => ({
      on: (event, listener) => { listeners[event] = listener; },
      start: async () => { listeners.transferFault?.(); throw new Error('装载期间已发生变化交付故障'); },
      destroy: async () => {},
    }),
  };
  await assert.rejects(startKitPage({ ...page, loadEditor: async () => editorModule }), /装载期间/);
  await new Promise(resolve => setTimeout(resolve, 20));
  assert.equal(page.requests.filter(request => request.endsWith('/content')).length, 1, 'no second content fetch: no rebuild');
  assert.equal(page.requests.includes('replaceChildren'), false);
  assert.equal(page.posted.filter(message => message.type === 'officesdk:failed').length, 1);
});

/** A PPT page whose parent can send commands: `send(op, args)` resolves with the agent result. */
async function pptPage({ mode = 'edit' } = {}) {
  const page = fakePage({ meta: { v: 1, id: 'v1', suite: 'presentation', mode, name: 'a.pptx' } });
  let receive;
  const afterprint = new Set();
  page.window.addEventListener = (type, listener) => { if (type === 'message') receive = listener; if (type === 'afterprint') afterprint.add(listener); };
  page.window.removeEventListener = (type, listener) => { if (type === 'afterprint') afterprint.delete(listener); };
  page.window.print = () => { page.printed = (page.printed ?? 0) + 1; };
  page.closePrint = () => { for (const listener of [...afterprint]) listener(); };
  const calls = [];
  const listeners = {};
  const doc = { canUndo: false, canRedo: false };
  const internals = {
    getDocument: () => doc,
    on: (event, listener) => { (listeners[event] ??= []).push(listener); return () => {}; },
    undo: () => calls.push('undo'), redo: () => calls.push('redo'),
    dispatch: intent => calls.push(intent.commandId),
  };
  const editorModule = {
    parseResourceCatalog: value => value, parseSystemFontCatalog: value => value, presentationResourceCategory: () => 'media',
    createOfflineEditor: async () => ({ suite: 'presentation', on() {}, start: async () => {}, prepare: async () => {}, destroy: async () => {},
      api: { presentation: { editorInternals: () => internals, getJsApi: () => ({}) } } }),
  };
  await startKitPage({ ...page, loadEditor: async () => editorModule });
  const send = (op, args) => new Promise(resolve => {
    const requestId = `r${Math.random().toString(36).slice(2, 8)}`;
    receive({ source: page.parent ?? page.window.parent, origin: 'http://app.local', data: frameCommand({ id: 'v1', revision: 1, command: 'agent', requestId, op, args }) });
    const poll = () => { const result = page.posted.find(message => message.type === 'officesdk:agent-result' && message.requestId === requestId); result ? resolve(result) : setTimeout(poll, 1); };
    poll();
  });
  const fire = event => { for (const listener of listeners[event] ?? []) listener(); };
  const history = () => page.posted.filter(message => message.type === 'officesdk:history').map(({ canUndo, canRedo }) => ({ canUndo, canRedo }));
  return { page, send, fire, doc, calls, history };
}

test('PPT toolbar undo/redo: the editor history is announced on change and drives undo/redo; read only refuses', async () => {
  const ppt = await pptPage();
  assert.deepEqual(ppt.history(), [{ canUndo: false, canRedo: false }]);
  ppt.doc.canUndo = true;
  ppt.fire('state-change');
  ppt.fire('commit');
  assert.deepEqual(ppt.history(), [{ canUndo: false, canRedo: false }, { canUndo: true, canRedo: false }], 'announced once per change');
  const undone = await ppt.send('history', { action: 'undo' });
  assert.equal(undone.ok, true);
  assert.deepEqual(ppt.calls, ['undo']);
  assert.equal((await ppt.send('history', { action: 'redo' })).ok, true);
  assert.equal((await ppt.send('history', { action: 'nope' })).code, 'unsupported-operation');
  assert.deepEqual(ppt.calls, ['undo', 'redo']);
  const viewing = await pptPage({ mode: 'readonly' });
  assert.deepEqual(viewing.history(), [{ canUndo: false, canRedo: false }]);
  assert.equal((await viewing.send('history', { action: 'undo' })).code, 'read-only');
});

test('print opens the browser\'s print dialog after replying, one at a time; modals stay quiet', async () => {
  const ppt = await pptPage();
  const [first, second] = await Promise.all([ppt.send('print'), ppt.send('print')]);
  assert.deepEqual([first.result, second.result], [{}, { ignored: true }], 'a second click while the dialog is coming is ignored');
  await new Promise(resolve => setTimeout(resolve, 5));
  assert.equal(ppt.page.printed, 1);
  assert.deepEqual((await ppt.send('print')).result, { ignored: true }, 'the dialog is still open');
  ppt.page.closePrint();
  assert.deepEqual((await ppt.send('print')).result, {}, 'once the dialog closed, printing again works');
  assert.equal(ppt.page.window.confirm('leave?'), false);
  assert.equal(ppt.page.window.prompt('name?'), null);
  assert.equal(ppt.page.window.onbeforeunload, null);
});

/** A page whose editor makes changes on demand: `edit()` emits one; `input(event)` plays a page event (any suite). */
async function changingPage({ createAgent } = {}) {
  const page = fakePage({ meta: { v: 1, id: 'v1', suite: 'presentation', mode: 'edit', name: 'a.pptx' } });
  const listeners = {}, events = {}, sent = [];
  let receive;
  page.window.addEventListener = (type, listener) => { if (type === 'message') receive = listener; (events[type] ??= []).push(listener); };
  const fetch = page.fetch;
  page.fetch = async (url, init = {}) => { if (new URL(url).pathname.endsWith('/change')) sent.push(JSON.parse(init.body)); return fetch(url, init); };
  const editorModule = {
    parseResourceCatalog: value => value, parseSystemFontCatalog: value => value, presentationResourceCategory: () => 'media',
    createOfflineEditor: async () => ({ suite: 'presentation', on: (event, listener) => { listeners[event] = listener; }, start: async () => {}, prepare: async () => {}, destroy: async () => {} }),
  };
  await startKitPage({ ...page, loadEditor: async () => editorModule, createAgent });
  const settle = () => new Promise(resolve => setTimeout(resolve, 5));
  const edit = async body => { listeners.change({ format: 'modoc-v1', body }); await settle(); };
  const input = (type, isTrusted = true) => { for (const listener of events[type] ?? []) listener({ type, isTrusted }); };
  const send = (op, args) => new Promise(resolve => {
    const requestId = `r${Math.random().toString(36).slice(2, 8)}`;
    receive({ source: page.window.parent, origin: 'http://app.local', data: frameCommand({ id: 'v1', revision: 1, command: 'agent', requestId, op, args }) });
    const poll = () => { const result = page.posted.find(message => message.type === 'officesdk:agent-result' && message.requestId === requestId); result ? resolve(result) : setTimeout(poll, 1); };
    poll();
  });
  const dirty = () => page.posted.filter(message => message.type === 'officesdk:dirty').map(({ editRevision, manual }) => ({ editRevision, manual }));
  const fault = async () => { listeners.transferFault(); await new Promise(resolve => setTimeout(resolve, 20)); };
  return { page, sent, edit, input, send, dirty, events, fault };
}

test('a change the editor makes before anyone acts reaches Kit as inert and leaves the document clean', async () => {
  const sheet = await changingPage();
  for (const type of ACT_EVENTS) assert.ok(sheet.events[type]?.length, `${type} is watched`);
  await sheet.edit('calc-csr reset');
  sheet.input('keydown', false); // a synthetic event is not the user
  await sheet.edit('more housekeeping');
  assert.deepEqual(sheet.sent, [{ seq: 0, format: 'modoc-v1', body: 'calc-csr reset', inert: true }, { seq: 1, format: 'modoc-v1', body: 'more housekeeping', inert: true }]);
  assert.deepEqual(sheet.dirty(), [], 'no dirty: nothing to autosave');
  sheet.input('keydown');
  await sheet.edit('typed');
  assert.deepEqual(sheet.sent[2], { seq: 2, format: 'modoc-v1', body: 'typed' });
  assert.deepEqual(sheet.dirty(), [{ editRevision: 3, manual: true }], 'a hand edit is dirty and autosaved as before');
});

test('pointer input or an AI operation also ends the editor\'s own changes', async () => {
  const clicked = await changingPage();
  clicked.input('pointerdown');
  await clicked.edit('cell');
  assert.deepEqual([clicked.sent[0].inert, clicked.dirty()], [undefined, [{ editRevision: 1, manual: true }]]);
  const ai = await changingPage();
  assert.equal((await ai.send('apply', { ops: [] })).code, 'unsupported-operation');
  await ai.edit('after an AI request');
  assert.equal(ai.sent[0].inert, undefined);
  assert.equal(ai.dirty().length, 1);
});

test('the toolbar history and an AI edit end the editor\'s own changes; an AI change is dirty but not a hand edit', async () => {
  const toolbar = await changingPage();
  await toolbar.send('history', { action: 'undo' });
  await toolbar.edit('undone');
  assert.deepEqual([toolbar.sent[0].inert, toolbar.dirty()], [undefined, [{ editRevision: 1, manual: true }]]);
  const agent = { applying: true, inspect: async () => ({}) };
  const ai = await changingPage({ createAgent: () => agent });
  await ai.edit('AI writes');
  assert.deepEqual([ai.sent[0].inert, ai.dirty()], [undefined, [{ editRevision: 1, manual: undefined }]]);
});

test('a reload before anyone acted keeps the document clean; after a hand edit it reports the recovered edits', async () => {
  const untouched = await changingPage();
  await untouched.edit('calc-csr reset');
  await untouched.fault();
  assert.equal(untouched.page.posted.filter(message => message.type === 'officesdk:dirty').length, 0);
  const touched = await changingPage();
  touched.input('keydown');
  await touched.fault();
  assert.deepEqual(touched.page.posted.filter(message => message.type === 'officesdk:dirty').map(message => message.recovered), [true]);
});

test('261109: an AI edit written cell by cell is a real change in every write, after the editor\'s own inert one', async () => {
  let emit;
  const agent = {
    applying: false, inspect: async () => ({}),
    // One change per cell, as the Sheet agent's cell-by-cell write produces them.
    async apply() { this.applying = true; try { for (const cell of ['D4', 'E4', 'F4', 'D5']) await emit(cell); return { applied: 1 }; } finally { this.applying = false; } },
  };
  const sheet = await changingPage({ createAgent: () => agent });
  emit = sheet.edit;
  await sheet.edit('calc-csr reset');
  const result = await sheet.send('apply', { ops: [] });
  assert.equal(result.type, 'officesdk:agent-result');
  assert.deepEqual(sheet.sent.map(change => [change.body, change.inert]), [['calc-csr reset', true], ['D4', undefined], ['E4', undefined], ['F4', undefined], ['D5', undefined]]);
  assert.ok(sheet.dirty().length >= 1 && sheet.dirty().every(entry => entry.manual === undefined), 'the AI writes make the document unsaved without counting as hand edits');
  assert.equal(sheet.dirty().at(-1).editRevision, 5);
});

test('STRATEGY-6984: the page language is the tab\'s lang, else <html lang>, else Chinese', () => {
  const params = query => new URLSearchParams(query);
  const html = lang => ({ documentElement: { lang } });
  assert.equal(pageLocale(params('lang=en'), html('zh')), 'en');
  assert.equal(pageLocale(params('lang=en-US'), html('zh')), 'en');
  assert.equal(pageLocale(params('lang=zh'), html('en')), 'zh');
  assert.equal(pageLocale(params(''), html('en')), 'en', 'without lang the page keeps following <html lang>');
  assert.equal(pageLocale(params(''), html('zh')), 'zh');
  assert.equal(pageLocale(params(''), html('')), 'zh');
  assert.equal(pageLocale(params('lang=ja'), html('en')), 'zh', 'only English has its own labels');
  assert.equal(pageLocale(undefined, undefined), 'zh');
});

test('STRATEGY-6984: the tab\'s lang reaches the editor language and the AI overlay; without it nothing changes', async () => {
  for (const [query, editorLanguage, locale] of [['&lang=en', 'en-US', 'en'], ['&lang=zh', 'zh-CN', 'zh'], ['', 'zh-CN', 'zh']]) {
    const page = fakePage({ meta: { v: 1, id: 'v1', suite: 'presentation', mode: 'edit', name: 'a.pptx' } });
    page.window.location.href += query;
    const created = [], agents = [];
    const editorModule = {
      parseResourceCatalog: value => value, parseSystemFontCatalog: value => value, presentationResourceCategory: () => 'media',
      createOfflineEditor: async options => { created.push(options); return { suite: 'presentation', on() {}, start: async () => {}, prepare: async () => {}, destroy: async () => {} }; },
    };
    await startKitPage({ ...page, loadEditor: async () => editorModule, createAgent: options => { agents.push(options.locale); return undefined; } });
    assert.equal(created[0].language, editorLanguage, query || 'no lang');
    assert.deepEqual(agents, [locale], query || 'no lang');
  }
});
