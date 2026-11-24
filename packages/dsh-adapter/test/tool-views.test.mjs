import test from 'node:test';
import assert from 'node:assert/strict';
import { TOOL_VIEW_MESSAGES, TOOL_VIEW_NS, TOOL_VIEW_TOOLS, callKey, createToolView, createToolViewStore, editContext, editKey, editPages, editRows, failureOf, pageRanges, pagesLabel, registerToolViews, settleContext, toolViewModel } from '../src/client/tool-views.mjs';

const zh = key => TOOL_VIEW_MESSAGES.zh[key];
const en = key => TOOL_VIEW_MESSAGES.en[key];
const text = value => [{ type: 'text', text: typeof value === 'string' ? value : JSON.stringify(value) }];
const started = (args, extra = {}) => ({ phase: 'start', block: { callId: 'c1', name: 'x', argsRaw: JSON.stringify(args), args: reader(args), subCalls: [] }, ...extra });
const settled = (args, content, extra = {}) => ({ phase: 'result', block: { kind: 'result', callId: 'c1', isError: false, content, call: { name: 'x', argsRaw: JSON.stringify(args) }, subCalls: [], ...extra } });
/** A failed call as the Client sees it: text `Error: <message>`; a plugin throw carries no `error` object. */
const failed = (args, message, error) => ({ phase: 'result', block: { kind: 'result', callId: 'c1', isError: true, content: text(`Error: ${message}`), ...(error && { error }), call: { argsRaw: JSON.stringify(args) }, subCalls: [] } });
/** DSH's lazy preparing reader: `text` grows, `complete` says the string closed. */
function reader(values, complete = Object.keys(values)) {
  return { text: key => typeof values[key] === 'string' ? values[key] : undefined, complete: key => complete.includes(key) };
}
const props = (toolName, call, extra = {}) => ({ t: zh, callId: 'c1', toolName, ...call, ...extra });

const DECK_OPS = [
  { op: 'addSlide', ref: 'p3', index: 2, layout: 'L2', placeholders: { title: '洱海' } },
  { op: 'addTextBox', slideRef: 'p3', text: '三段正文\n第二行', left: 1, top: 1, width: 10, height: 10 },
  { op: 'setText', slideId: 's2', shapeId: 't1', text: '新的标题' },
  { op: 'replaceText', slideId: 's5', shapeId: 'b', find: '旧', replace: '新' },
  { op: 'setFill', slideId: 's9', shapeId: 'r', color: '#2826C9' },
];
const DECK = { slides: [{ id: 's1', shapes: [] }, { id: 's2', shapes: [{ id: 't1', name: '标题 1' }] }, { id: 's3', shapes: [] }, { id: 's4', shapes: [] }, { id: 's5', shapes: [{ id: 'b', name: '' }] }] };

test('only the plugin\'s presentation, document and spreadsheet tools get a view, under their own names, inside effects, without a priority', () => {
  assert.deepEqual([...TOOL_VIEW_TOOLS], ['presentation_create', 'presentation_read', 'presentation_edit', 'presentation_undo', 'presentation_save', 'document_create', 'document_read', 'document_edit', 'document_save', 'spreadsheet_create', 'spreadsheet_read', 'spreadsheet_edit', 'spreadsheet_save']);
  const registered = [], locales = [], disposed = [];
  let effects = 0;
  const ctx = {
    effect: fn => { effects += 1; disposed.push(fn()); },
    locale: { register: (ns, messages) => { locales.push([ns, messages]); return () => {}; } },
    slots: { inject: (name, fn) => { assert.equal(name, 'tool.call.toolview'); return fn(); }, register: (options, component) => { registered.push([options, component]); return () => {}; } },
  };
  registerToolViews(ctx, fakeReact(), { store: createToolViewStore() });
  assert.deepEqual(registered.map(([{ inject: _inject, ...options }]) => options), TOOL_VIEW_TOOLS.map(key => ({ name: 'tool.call.toolview', key, locale: TOOL_VIEW_NS })));
  assert.ok(registered.every(([options, component]) => !('priority' in options) && typeof component === 'function'));
  assert.deepEqual(registered[0][0].inject('s7'), { shimoSessionId: 's7' }, 'a plugin-specific name: owner props win clashes');
  assert.deepEqual(locales, [[TOOL_VIEW_NS, TOOL_VIEW_MESSAGES]]);
  assert.equal(effects, 14, 'every registration is undone with the plugin');
  assert.ok(disposed.every(dispose => typeof dispose === 'function'));
});

test('zh and en dictionaries have the same keys', () => {
  assert.deepEqual(Object.keys(TOOL_VIEW_MESSAGES.en).sort(), Object.keys(TOOL_VIEW_MESSAGES.zh).sort());
});

test('pages: new slide positions, refs to them, recorded slide ids; unknown ones are left out', () => {
  assert.deepEqual(editPages(DECK_OPS), { pages: [3], unplaced: 0 });
  const context = editContext(DECK_OPS, DECK);
  assert.deepEqual(context, { slides: { s2: 2, s5: 5 }, shapes: { 's2/t1': '标题 1', 's5/b': 'b' }, refs: {} });
  assert.deepEqual(editPages(DECK_OPS, context), { pages: [2, 3, 5], unplaced: 0 });
  assert.equal(pageRanges([2, 3, 4, 5, 7], '、'), '2–5、7');
  assert.equal(pageRanges([2, 3, 7], ', '), '2, 3, 7');
  assert.equal(pageRanges([4], '、'), '4');
  assert.deepEqual(editPages('nope'), { pages: [], unplaced: 0 });
  assert.deepEqual(editPages([null, { op: 'addSlide', index: -1 }, { op: 'setText', slideRef: 'missing' }]), { pages: [], unplaced: 1 });
  // Appended slides (no index) without a record: counted, not placed.
  const appended = [{ op: 'addSlide', ref: 'a' }, { op: 'addTextBox', slideRef: 'a', text: 'x' }, { op: 'addSlide' }];
  assert.equal(pagesLabel(zh, editPages(appended)), '新增 2 页');
  assert.equal(pagesLabel(en, editPages([...DECK_OPS, ...appended], context)), 'slides 2, 3, 5 · 2 new slides');
  // Settled after the edit: actual pages of the touched slides win, deleted ones keep their page.
  const before = editContext([{ op: 'deleteSlide', slideId: 's4' }, ...appended], DECK);
  const after = editContext(appended, { slides: [{ id: 's1' }, { id: 's2' }, { id: 'n1', shapes: [] }, { id: 's3' }, { id: 'n2' }] }, { slideIds: ['s4', 'n1', 'n2'], refs: { a: 'n1' } });
  const settled = settleContext(before, after, ['s4', 'n1', 'n2']);
  assert.deepEqual([settled.touched, settled.refs, settled.awaiting], [[3, 4, 5], { a: 3 }, false]);
  assert.equal(pagesLabel(zh, editPages(appended, settled)), '第 3–5 页');
  assert.deepEqual(editRows(zh, appended, settled).map(row => row.where), ['新增第 3 页', '第 3 页', '']);
  // A moved slide's row keeps the page it left; its target is in the text; the summary pages are where it is now.
  const move = [{ op: 'moveSlide', slideId: 's2', index: 4 }];
  const moved = settleContext(editContext(move, DECK), editContext(move, { slides: [{ id: 's1' }, { id: 's3' }, { id: 's4' }, { id: 's5' }, { id: 's2' }] }, { slideIds: ['s2'] }), ['s2']);
  assert.deepEqual(editRows(zh, move, moved).map(row => [row.where, row.text]), [['第 2 页', '移到第 5 页']]);
  assert.equal(pagesLabel(zh, editPages(move, moved)), '第 5 页');
});

test('operation rows: what, where, content in short', () => {
  const rows = editRows(zh, DECK_OPS, editContext(DECK_OPS, DECK));
  assert.deepEqual(rows, [
    { op: 'addSlide', label: '新增页', where: '新增第 3 页', text: '洱海' },
    { op: 'addTextBox', label: '加文本框', where: '第 3 页', text: '三段正文' },
    { op: 'setText', label: '改文字', where: '第 2 页 · 标题 1', text: '新的标题' },
    { op: 'replaceText', label: '替换文字', where: '第 5 页 · b', text: '旧 → 新' },
    { op: 'setFill', label: '改填充色', where: '', text: '#2826C9' },
  ]);
  assert.equal(editRows(en, [{ op: 'moveSlide', slideId: 'x', index: 0 }])[0].text, 'to slide 1');
  assert.equal(editRows(zh, [{ op: 'addSlide', index: 0 }, { op: 'deleteSlide', slideId: 's1' }], { slides: { s1: 1 } })[1].where, '第 1 页');
});

test('edit row: running, then applied with the recorded pages', () => {
  const args = { path: 'decks/洱海介绍.pptx', summary: '新增三段正文\n细节', operations: DECK_OPS };
  const running = toolViewModel(zh, props('presentation_edit', started(args)));
  assert.equal(running.state, 'running');
  assert.equal(running.title, '正在修改演示文稿');
  assert.equal(running.file, '洱海介绍.pptx');
  assert.equal(running.pages, '第 3 页');
  assert.equal(running.summary, '新增三段正文');
  const store = createToolViewStore();
  store.record([callKey('c1'), editKey('e1')], { ...editContext(DECK_OPS, DECK), summary: args.summary, ops: DECK_OPS, awaiting: true });
  assert.equal(toolViewModel(zh, props('presentation_edit', started(args)), key => store.get(key)).title, '等你确认修改');
  store.record([callKey('c1'), editKey('e1')], { ...editContext(DECK_OPS, DECK), summary: args.summary, ops: DECK_OPS, awaiting: false });
  const done = toolViewModel(zh, props('presentation_edit', settled(args, text({ applied: 5, slideIds: ['s2'], editId: 'e1', unsaved: true }))), key => store.get(key));
  assert.equal(done.state, 'ok');
  assert.equal(done.title, '修改了演示文稿');
  assert.equal(done.pages, '第 2、3、5 页');
  assert.equal(done.rows.length, 5);
  assert.deepEqual(done.lines, ['已应用 5 步 · 还没保存到文件 · 可以撤销']);
});

test('stopped and not approved are told apart from failures', () => {
  const args = { path: 'a.pptx', summary: 's', operations: [] };
  const outcome = (call, tool = 'presentation_edit') => { const model = toolViewModel(zh, props(tool, call)); return [model.state, model.tone, model.title]; };
  assert.deepEqual(outcome(settled(args, text({ declined: true }))), ['declined', 'muted', '没有修改演示文稿']);
  assert.deepEqual(toolViewModel(zh, props('presentation_edit', settled(args, text({ declined: true })))).lines, ['你没有同意这次修改']);
  // DSH's own: interrupted (no result), stopped before dispatch, its approval rejected or cancelled.
  assert.deepEqual(outcome({ phase: 'result', block: { isError: true, content: [], error: { name: 'Interrupted', code: 'interrupted' } } }), ['stopped', 'warn', '已停止']);
  assert.deepEqual(outcome(failed(args, 'tool call aborted before dispatch', { name: 'AbortError', code: 'ABORTED_BEFORE_DISPATCH' })), ['stopped', 'warn', '已停止']);
  assert.deepEqual(outcome(failed(args, 'tool call aborted before dispatch')), ['stopped', 'warn', '已停止']);
  assert.deepEqual(outcome(failed(args, 'the user rejected tool "presentation_edit"')), ['declined', 'muted', '没有修改演示文稿']);
  assert.deepEqual(outcome(failed(args, 'approval for tool "presentation_edit" was cancelled')), ['stopped', 'warn', '已停止']);
  // The plugin's: Stop pressed (cancelled), save outside the workspace refused / cancelled / no way to ask.
  assert.deepEqual(outcome(failed(args, 'cancelled: The document command was cancelled.')), ['stopped', 'warn', '已停止']);
  const outside = reason => failed({ path: '/far/a.pptx' }, `sandbox-denied: The user did not allow writing /far/a.pptx outside the workspace (${reason}). The changes stay in the editor; they can save from the panel.`);
  assert.deepEqual(outcome(outside('rejected'), 'presentation_save'), ['declined', 'muted', '没有保存演示文稿']);
  assert.deepEqual(toolViewModel(zh, props('presentation_save', outside('rejected'))).lines, ['你没有同意写入工作区外的文件，修改还在编辑器里']);
  assert.deepEqual(outcome(outside('cancelled'), 'presentation_save'), ['stopped', 'warn', '已停止']);
  assert.deepEqual(outcome(outside('no approval channel'), 'presentation_save'), ['error', 'error', '保存演示文稿失败']);
  assert.deepEqual(outcome(failed({ path: 'a.pptx' }, 'sandbox-denied: This Session cannot write /w/a.pptx (read only).'), 'presentation_save'), ['error', 'error', '保存演示文稿失败']);
});

test('the packaged plugin\'s errors carry no code: recorded codes and the Host\'s fixed wording still classify them', () => {
  const args = { path: 'a.pptx', summary: 's', operations: [] };
  const outcome = (call, tool = 'presentation_edit', recall) => { const model = toolViewModel(zh, props(tool, call), recall); return [model.state, model.title, model.summary]; };
  for (const wording of ['The document command was cancelled.', 'The edit was cancelled; nothing was changed.', 'The document operation was cancelled.', 'Cancelled.']) {
    assert.deepEqual(outcome(failed(args, wording)), ['stopped', '已停止', 's'], wording);
  }
  assert.deepEqual(outcome(failed({ path: '/a/(rejected)/x.pptx' }, 'The user did not allow writing /a/(rejected)/x.pptx outside the workspace (cancelled). The changes stay in the editor.'), 'presentation_save'), ['stopped', '已停止', ''], 'the outcome is read at its fixed place, not anywhere in the path');
  assert.deepEqual(outcome(failed({ path: '/far/a.pptx' }, 'The user did not allow writing /far/a.pptx outside the workspace (rejected). The changes stay in the editor; they can save from the panel.'), 'presentation_save'), ['declined', '没有保存演示文稿', '']);
  assert.deepEqual(outcome(failed({ path: '/far/a.pptx' }, 'The user did not allow writing /far/a.pptx outside the workspace (cancelled). The changes stay in the editor.'), 'presentation_save'), ['stopped', '已停止', '']);
  const store = createToolViewStore();
  store.record([callKey('c1')], { slides: {}, shapes: {}, failure: 'busy' });
  assert.deepEqual(outcome(failed(args, 'Another edit is still running.'), 'presentation_edit', key => store.get(key)), ['error', '修改演示文稿失败', '文档正忙：Another edit is still running.']);
  store.record([callKey('c1')], { failure: 'cancelled' });
  assert.deepEqual(outcome(failed(args, 'Cancelled.'), 'presentation_edit', key => store.get(key)), ['stopped', '已停止', 's']);
  // Without a record: the reason alone, still a failure.
  assert.deepEqual(outcome(failed(args, 'Another edit is still running.')), ['error', '修改演示文稿失败', 'Another edit is still running.']);
});

test('failures name the reason inline and keep the full text', () => {
  const args = { path: 'a.pptx', summary: 's', operations: [] };
  const model = toolViewModel(zh, props('presentation_edit', failed(args, 'invalid-edit: Operation 2: shape t9 not found.\nmore')));
  assert.deepEqual([model.state, model.tone, model.title, model.summary], ['error', 'error', '修改演示文稿失败', '操作无效：Operation 2: shape t9 not found.']);
  assert.equal(model.error, 'Error: invalid-edit: Operation 2: shape t9 not found.\nmore');
  assert.equal(toolViewModel(en, props('presentation_save', failed({ path: 'a.pptx' }, 'read-only: This Session is read-only.'), { t: en })).summary, 'Read-only session: This Session is read-only.');
  assert.equal(failureOf({ isError: true, content: text('Error: officesdk/no-editor: gone') }).code, 'no-editor');
  // Codes the Host never reports get no category (it turns unknown Client codes into editor-failed).
  assert.deepEqual(toolViewModel(zh, props('presentation_undo', failed({ path: 'a', editId: 'e' }, 'undo-unavailable: Only the most recent AI edit can be undone.'))).summary, 'Only the most recent AI edit can be undone.');
  assert.equal(toolViewModel(zh, props('presentation_undo', failed({ path: 'a', editId: 'e' }, 'editor-failed: That edit is no longer in this editor.'))).summary, '编辑器出错：That edit is no longer in this editor.');
  assert.equal(toolViewModel(zh, props('presentation_read', failed({ path: 'a.pptx' }, 'something odd'))).summary, 'something odd');
  assert.equal(toolViewModel(zh, props('presentation_read', { phase: 'result', block: { isError: true, content: [], error: { name: 'Error', code: 'E1' } } })).summary, 'E1', 'no text: the error object, prefix dropped');
});

test('read, undo and save rows', () => {
  const deck = { slides: [{ id: 'a', index: 0, shapes: [{ id: '1', text: '  \n' }, { id: '2', text: '标题\n副标题' }] }, { id: 'b', index: 1, shapes: [] }] };
  const read = toolViewModel(zh, props('presentation_read', settled({ path: '/abs/x.pptx' }, text({ format: 'pptx', path: '/abs/x.pptx', unsaved: true, deck }))));
  assert.deepEqual([read.title, read.file, read.summary], ['读取了演示文稿', 'x.pptx', '共 2 页 · 有未保存的修改']);
  assert.deepEqual(read.slides, [{ page: '第 1 页', shapes: '2 个元素', text: '标题' }, { page: '第 2 页', shapes: '0 个元素', text: '' }]);
  const store = createToolViewStore();
  store.record([callKey('c0'), editKey('e7')], { slides: { s2: 2 }, shapes: {}, summary: '改标题', ops: [{ op: 'setText', slideId: 's2', shapeId: 't', text: 'x' }] });
  const undo = toolViewModel(zh, props('presentation_undo', settled({ path: 'x.pptx', editId: 'e7' }, text({ undone: 'e7', unsaved: true }))), key => store.get(key));
  assert.deepEqual([undo.title, undo.pages, undo.summary, undo.lines], ['撤销了修改', '第 2 页', '改标题', ['已在编辑器里撤销，还没保存到文件']]);
  // A read result cut short by DSH: the outline recorded for this call stands in.
  store.record([callKey('c1')], { outline: [{ shapes: 3, text: '封面' }], unsaved: false });
  const cut = toolViewModel(zh, props('presentation_read', settled({ path: 'x.pptx' }, text('{"format":"pptx","deck":{"slides":[{"id"'))), key => store.get(key));
  assert.deepEqual([cut.summary, cut.slides], ['共 1 页', [{ page: '第 1 页', shapes: '3 个元素', text: '封面' }]]);
  const unknownUndo = toolViewModel(zh, props('presentation_undo', settled({ path: 'x.pptx', editId: 'zz' }, text({ undone: 'zz' }))));
  assert.deepEqual([unknownUndo.pages, unknownUndo.summary], ['', '']);
  const saved = toolViewModel(zh, props('presentation_save', settled({ path: 'x.pptx' }, text({ saved: true, version: 'v2', path: '/w/x.pptx' }))));
  assert.deepEqual([saved.title, saved.lines], ['保存了演示文稿', ['已写回 /w/x.pptx']]);
  const same = toolViewModel(zh, props('presentation_save', settled({ path: 'x.pptx' }, text({ saved: false, unchanged: true }))));
  assert.deepEqual([same.state, same.title], ['unchanged', '没有要保存的修改']);
});

test('create row: file, style, generating; refused or existing files', () => {
  const running = toolViewModel(zh, props('presentation_create', started({ path: 'out/洱海.pptx' })));
  assert.deepEqual([running.title, running.file, running.summary], ['正在新建演示文稿', '洱海.pptx', ''], 'no style named: the Host picks the workspace default, not known yet');
  assert.equal(toolViewModel(zh, props('presentation_create', started({ path: 'out/洱海.pptx', style: 'second' }))).summary, '风格 second');
  assert.equal(toolViewModel(zh, props('presentation_create', settled({ path: 'out/洱海.pptx' }, text({ created: true, style: { name: 'second' }, styleSource: 'workspace' })))).summary, '风格 second', 'the style the Host chose');
  const done = toolViewModel(zh, props('presentation_create', settled({ path: 'out/洱海.pptx', style: 'blue' }, text({ created: true, style: { name: 'blue', palette: {} }, format: 'pptx', path: 'out/洱海.pptx', deck: { slides: [] }, generating: true }))));
  assert.deepEqual([done.title, done.summary, done.lines], ['新建了演示文稿', '风格 blue', ['在右侧幻灯片浏览视图中逐页生成']]);
  const refused = toolViewModel(zh, props('presentation_create', failed({ path: '/far/x.pptx' }, 'sandbox-denied: The user did not allow creating /far/x.pptx outside the workspace (rejected). Create it inside the workspace instead.')));
  assert.deepEqual([refused.state, refused.title, refused.lines], ['declined', '没有新建演示文稿', ['你没有同意']]);
  const streaming = reader({ path: 'a.pptx', style: 'blu' }, ['path']);
  assert.equal(toolViewModel(zh, props('presentation_create', { phase: 'preparing', block: { args: streaming } })).summary, '', 'style not written yet: left out');
  assert.equal(toolViewModel(zh, props('presentation_create', { phase: 'preparing', block: { args: reader({ path: 'a.pptx', style: 'blue' }) } })).summary, '风格 blue');
  assert.equal(toolViewModel(zh, props('presentation_create', failed({ path: 'x.pptx', style: 'x' }, 'invalid-edit: Unknown style "x"; available: default.'))).summary, 'Unknown style "x"; available: default.');
  const exists = toolViewModel(zh, props('presentation_create', failed({ path: 'x.pptx' }, 'exists: A file with that name already exists.')));
  assert.deepEqual([exists.state, exists.summary], ['error', '同名文件已存在：A file with that name already exists.']);
  const Row = createToolView({ React: fakeReact(), store: createToolViewStore(), openDeck: () => {}, canOpenDeck: () => true });
  const tree = Row(props('presentation_create', settled({ path: 'out/洱海.pptx' }, text({ created: true, style: { name: 'default' } })), { shimoSessionId: 's1', useDisclosure: () => ({ expanded: false, toggle() {} }) }));
  assert.equal(words(byClass(tree, 'sbt-row')[0]), '新建了演示文稿洱海.pptx风格 default');
  assert.equal(byClass(tree, 'sbt-open').length, 1);
});

test('incomplete or malformed input never throws', () => {
  const preparing = toolViewModel(zh, props('presentation_edit', { phase: 'preparing', block: { callId: 'c1', name: 'presentation_edit', args: reader({ path: 'dir/a.pp', summary: 'half' }, ['summary']) } }));
  assert.deepEqual([preparing.state, preparing.path, preparing.summary], ['preparing', '', 'half']);
  for (const call of [
    { phase: 'start', block: { argsRaw: '{"path": "a.pptx", "operations": [' } },
    { phase: 'start', block: {} },
    { phase: 'result', block: { isError: false, content: text('not json'), call: { argsRaw: '[]' } } },
    { phase: 'result', block: { isError: false, content: 'weird', call: {} } },
    { phase: 'result', block: undefined },
    settled({ path: 7, operations: 'x', summary: {} }, text({ applied: 'many', deck: { slides: 'no' } })),
  ]) for (const tool of TOOL_VIEW_TOOLS) assert.doesNotThrow(() => toolViewModel(zh, props(tool, call)), `${tool} ${JSON.stringify(call)}`);
  assert.equal(toolViewModel(undefined, props('presentation_read', started({ path: 'a.pptx' })), undefined).title, 'read.running', 'no locale: keys');
});

// A React stand-in: elements as plain objects, hooks as plain values.
function fakeReact({ expanded = false } = {}) {
  const toggles = [];
  return {
    toggles,
    createElement: (type, props, ...children) => ({ type, props: props ?? {}, children: children.flat(Infinity).filter(child => child !== null && child !== undefined && child !== false) }),
    useState: initial => [expanded || initial, () => toggles.push('local')],
    useSyncExternalStore: (subscribe, get) => { subscribe(() => {})(); return get(); },
  };
}
const all = node => (node && typeof node === 'object' ? [node, ...node.children.flatMap(all)] : []);
const words = node => (typeof node === 'string' ? node : node && typeof node === 'object' ? node.children.map(words).join('') : '');
const byClass = (tree, name) => all(tree).filter(node => (node.props.className ?? '').split(' ').includes(name));

test('rendered row: summary line, open beside, expand to operations, inline error', () => {
  const store = createToolViewStore();
  store.record([callKey('c1')], editContext(DECK_OPS, DECK));
  const opened = [];
  const args = { path: 'decks/洱海介绍.pptx', summary: '新增三段正文', operations: DECK_OPS };
  const React = fakeReact();
  const mounted = { id: 's9' };
  const Row = createToolView({ React, store, openDeck: (sessionId, path) => opened.push([sessionId, path]), canOpenDeck: sessionId => sessionId === mounted.id });
  const disclosure = { expanded: false, toggle: () => opened.push('toggle') };
  const collapsed = Row(props('presentation_edit', settled(args, text({ applied: 5, editId: 'e1', unsaved: true })), { shimoSessionId: 's9', openFile: () => opened.push('dsh openFile'), useDisclosure: () => disclosure }));
  assert.equal(words(byClass(collapsed, 'sbt-row')[0]), '修改了演示文稿洱海介绍.pptx第 2、3、5 页新增三段正文');
  assert.equal(byClass(collapsed, 'sbt-file')[0].props.title, 'decks/洱海介绍.pptx');
  assert.equal(byClass(collapsed, 'sbt-body').length, 0);
  const open = byClass(collapsed, 'sbt-open')[0];
  assert.equal(open.props['aria-label'], '在右侧打开', 'named even when a narrow column shows only the icon');
  let stopped = false;
  open.props.onClick({ stopPropagation: () => { stopped = true; } });
  assert.deepEqual(opened, [['s9', 'decks/洱海介绍.pptx']], 'the plugin address, not DSH openFile');
  assert.ok(stopped, 'opening does not toggle the row');
  const row = byClass(collapsed, 'sbt-row')[0];
  assert.equal(row.props.role, 'button');
  row.props.onKeyDown({ key: 'Enter', preventDefault() {} });
  assert.deepEqual(opened.at(-1), 'toggle');

  disclosure.expanded = true;
  const expanded = Row(props('presentation_edit', settled(args, text({ applied: 5, editId: 'e1', unsaved: true })), { shimoSessionId: 's9', useDisclosure: () => disclosure, inspect: () => {} }));
  const items = all(byClass(expanded, 'sbt-list')[0]).filter(node => node.type === 'li').map(words);
  assert.deepEqual(items, ['新增页新增第 3 页洱海', '加文本框第 3 页三段正文', '改文字第 2 页 · 标题 1新的标题', '替换文字第 5 页 · b旧 → 新', '改填充色#2826C9']);
  assert.equal(words(byClass(expanded, 'sbt-result')[0]), '已应用 5 步 · 还没保存到文件 · 可以撤销');
  assert.equal(byClass(expanded, 'sbt-inspect').length, 1);

  const error = Row(props('presentation_save', failed({ path: 'x.pptx' }, 'sandbox-denied: nope'), { shimoSessionId: 'other', useDisclosure: () => ({ expanded: true, toggle() {} }) }));
  assert.equal(error.props['data-state'], 'error');
  assert.equal(words(byClass(error, 'sbt-summary')[0]), '没有写入权限：nope');
  assert.ok(byClass(error, 'sbt-summary')[0].props.className.includes('sbt-error'));
  assert.equal(byClass(error, 'sbt-open').length, 0, 'not the Session the right Sidebar shows → no button');
  assert.match(words(all(error).find(node => node.type === 'pre')), /sandbox-denied: nope/);
  const noSession = Row(props('presentation_read', settled({ path: 'x.pptx' }, text({ deck: { slides: [] } })), { useDisclosure: () => disclosure }));
  assert.equal(byClass(noSession, 'sbt-open').length, 0, 'no Session id → no button');
});

test('rendered row: a stop or failure colours the title; the summary only when it is the reason', () => {
  const Row = createToolView({ React: fakeReact(), store: createToolViewStore() });
  const args = { path: 'a.pptx', summary: '加三页', operations: [] };
  const stopped = Row(props('presentation_edit', failed(args, 'cancelled: The document command was cancelled.'), { useDisclosure: () => ({ expanded: false, toggle() {} }) }));
  assert.ok(byClass(stopped, 'sbt-title')[0].props.className.includes('sbt-warn'));
  assert.ok(!byClass(stopped, 'sbt-summary')[0].props.className.includes('sbt-warn'), 'the model\'s own summary is not a warning');
  const error = Row(props('presentation_read', failed({ path: 'a.pptx' }, 'busy: wait'), { useDisclosure: () => ({ expanded: false, toggle() {} }) }));
  assert.ok(byClass(error, 'sbt-title')[0].props.className.includes('sbt-error'));
  assert.ok(byClass(error, 'sbt-summary')[0].props.className.includes('sbt-error'));
  const declined = Row(props('presentation_edit', settled(args, text({ declined: true })), { useDisclosure: () => ({ expanded: false, toggle() {} }) }));
  assert.equal(byClass(declined, 'sbt-title')[0].props.className, 'sbt-title');
});

test('the open button follows the Session the right Sidebar shows', () => {
  let shown = 's1';
  const subscribed = [];
  const Row = createToolView({ React: fakeReact(), store: createToolViewStore(), openDeck: () => {}, canOpenDeck: sessionId => sessionId === shown, subscribeDeck: listener => { subscribed.push(listener); return () => {}; } });
  const row = () => Row(props('presentation_read', settled({ path: 'a.pptx' }, text({ deck: { slides: [] } })), { shimoSessionId: 's1', useDisclosure: () => ({ expanded: false, toggle() {} }) }));
  assert.equal(byClass(row(), 'sbt-open').length, 1);
  shown = 's2';
  assert.equal(byClass(row(), 'sbt-open').length, 0);
  assert.ok(subscribed.length > 0, 'subscribes to the Sidebar\'s Session');
});

test('rendered row while preparing: no expansion, no open button until the path is complete', () => {
  const Row = createToolView({ React: fakeReact({ expanded: true }), store: createToolViewStore(), openDeck: () => {} });
  const tree = Row(props('presentation_edit', { phase: 'preparing', block: { args: reader({ path: 'a.pp' }, []) } }, { shimoSessionId: 's1' }));
  assert.equal(byClass(tree, 'sbt-row')[0].props.role, undefined);
  assert.equal(byClass(tree, 'sbt-open').length, 0);
  assert.equal(byClass(tree, 'sbt-body').length, 0);
  assert.equal(byClass(tree, 'sbt-row')[0].props['data-busy'], true);
});

test('store keeps the newest entries and tells subscribers', () => {
  const store = createToolViewStore({ max: 2 });
  let calls = 0;
  const off = store.subscribe(() => calls++);
  store.record(['a', '', undefined, 'b'], 1);
  store.record(['c'], 2);
  assert.deepEqual([store.get('a'), store.get('b'), store.get('c')], [undefined, 1, 2]);
  off();
  store.record(['d'], 3);
  assert.equal(calls, 2);
});
