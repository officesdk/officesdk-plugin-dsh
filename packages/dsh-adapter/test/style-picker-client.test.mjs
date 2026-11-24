import test from 'node:test';
import assert from 'node:assert/strict';
import { STYLE_PICKER_ID, STYLE_PICKER_MESSAGES, STYLE_PICKER_NS, appendText, caretAtStart, composerOf, createStylePicker, createStylePickerStores, focusComposer, plainBackspace, registerStylePicker } from '../src/client/style-picker.mjs';
import { InspirationLibrary } from '../src/client/inspiration-library.mjs';
import { PROMPTS } from '../src/client/inspiration.mjs';
import { CSS } from '../src/client/styles.mjs';

const zh = key => STYLE_PICKER_MESSAGES.zh[key];
const en = key => STYLE_PICKER_MESSAGES.en[key];
const DEFAULT = { id: 'default', name: { zh: '蓝色品牌', en: 'Blue brand' }, aliases: [], summary: { zh: '企业汇报', en: 'Business' }, scenarios: { zh: ['业务汇报'], en: [] }, thumb: 'data:image/png;base64,AA==' };
const LIST = { styles: [DEFAULT], defaultStyle: 'default', remembered: null, canRemember: true };
const PPT = { kind: 'ppt', id: 'default', title: '蓝色品牌' };
const WORD = { kind: 'word', id: 'gongwen', title: '公文版式' };

function draftBox(draft = '') {
  const box = { draft, calls: [] };
  box.actions = { setDraft: text => { box.calls.push(['setDraft', text]); box.draft = text; }, submit: () => box.calls.push(['submit']) };
  return box;
}
function rpcStub(answers = {}) {
  const calls = [];
  return { calls, call: async (_path, method, { args }) => {
    const name = method.split('/')[1];
    calls.push([name, args]);
    const answer = answers[name];
    const value = typeof answer === 'function' ? await answer(args) : answer;
    return value instanceof Error ? { status: 'failure', error: { code: 'officesdk/forbidden', message: value.message } } : { status: 'success', value: value ?? null };
  } };
}
const tick = async (n = 4) => { for (let i = 0; i < n; i++) await new Promise(resolve => setImmediate(resolve)); };
function fakeReact(overrides = {}) {
  return {
    Fragment: 'fragment',
    useState: initial => [typeof initial === 'function' ? initial() : initial, () => {}],
    createElement: (type, props, ...children) => ({ type, props: props ?? {}, children: children.flat(Infinity).filter(child => child !== null && child !== undefined && child !== false) }),
    useRef: initial => ({ current: initial }),
    useEffect: () => {},
    useSyncExternalStore: (subscribe, get) => { subscribe(() => {})(); return get(); },
    ...overrides,
  };
}
const expand = node => (node && typeof node === 'object' && typeof node.type === 'function' ? expand(node.type(node.props)) : node);
const all = node => { node = expand(node); return node && typeof node === 'object' ? [node, ...node.children.flatMap(all)] : []; };
const words = node => { node = expand(node); return typeof node === 'string' ? node : node && typeof node === 'object' ? node.children.map(words).join('') : ''; };
const byClass = (tree, name) => all(tree).filter(node => (node.props.className ?? '').split(' ').includes(name));

/** The input box's toolbar slot of one Session, as DSH mounts it; `effects: true` runs effects as they render. */
function mount({ answers = {}, draft = '', t = zh, sessionId = 's1', library, effects = false, stores, rpc } = {}) {
  rpc ??= rpcStub({ listStyles: LIST, templatePick: null, selectTemplate: ({ kind, id }) => (id ? { kind, id } : null), ...answers });
  stores ??= createStylePickerStores(rpc);
  const React = fakeReact(effects ? { useEffect: run => { run(); } } : {});
  const StylePicker = createStylePicker({ React, stores, library });
  const box = draftBox(draft);
  const props = () => ({ sessionId, inputActions: box.actions, useInput: select => select({ draft: box.draft, attachmentIds: [] }), useSessionStatus: select => select(new Map([[sessionId, { running: false }]])), t });
  return { rpc, stores, store: stores.get(sessionId), box, toolbar: () => StylePicker(props()) };
}

test('registration: one list item in conversation.input.left with its own locale, inside effects; no 灵感库 button any more', () => {
  const slots = [], effects = [], locales = [];
  const ctx = { effect: fn => effects.push(fn), locale: { register: (ns, messages) => { locales.push([ns, messages]); return () => {}; } }, slots: { inject: (_n, fn) => fn(), register: (options, component) => { slots.push([options, component]); return () => {}; } } };
  registerStylePicker(ctx, fakeReact(), { rpc: rpcStub() });
  assert.equal(slots.length, 0);
  effects.forEach(run => run());
  assert.deepEqual(slots.map(([options]) => options), [{ name: 'conversation.input.left', id: STYLE_PICKER_ID, order: 100, locale: STYLE_PICKER_NS }]);
  assert.equal(locales[0][0], STYLE_PICKER_NS);
  const view = mount();
  assert.equal(all(view.toolbar()).filter(node => node.type === 'button').length, 0, 'nothing without a template');
  assert.equal(byClass(view.toolbar(), 'sbs-entry').length, 0);
});

test('use: the Host records the template first, then the kind\'s prompt goes at the start of the draft, user text kept; never sent', async () => {
  const view = mount({ draft: '重庆三日游' });
  const order = [];
  view.rpc.call = (orig => async (...args) => { order.push('host'); return orig(...args); })(view.rpc.call);
  view.box.actions.setDraft = (orig => text => { order.push('draft'); orig(text); })(view.box.actions.setDraft);
  const result = await view.store.use(PPT, { actions: view.box.actions, draft: view.box.draft });
  assert.deepEqual(result, { written: true });
  assert.deepEqual(order, ['host', 'draft']);
  assert.equal(view.box.draft, `${PROMPTS.zh.ppt}重庆三日游`);
  assert.deepEqual(view.rpc.calls.at(-1), ['selectTemplate', { workspaceFileScopeId: 's1', kind: 'ppt', id: 'default' }]);
  assert.ok(!view.box.calls.some(([name]) => name === 'submit'), 'nothing is sent');
  const tag = byClass(view.toolbar(), 'sbs-tag')[0];
  assert.equal(words(tag).replace('×', ''), '蓝色品牌', 'bulb + template name, no kind text, no box');
  assert.equal(byClass(tag, 'sbs-bulb').length, 1);
  assert.equal(byClass(tag, 'sbs-tag-x')[0].props['aria-label'], '移除模板 蓝色品牌');
});

test('use another template: only the previous prompt is replaced; a prompt the user rewrote stays', async () => {
  const view = mount();
  await view.store.use(PPT, { actions: view.box.actions, draft: '' });
  view.box.draft += '主题：年终总结';
  await view.store.use(WORD, { actions: view.box.actions, draft: view.box.draft });
  assert.equal(view.box.draft, `${PROMPTS.zh.word}主题：年终总结`);
  view.box.draft = `请${view.box.draft}`;
  await view.store.use(PPT, { actions: view.box.actions, draft: view.box.draft });
  assert.equal(view.box.draft, `${PROMPTS.zh.ppt}请${PROMPTS.zh.word}主题：年终总结`);
});

test('a failed binding never looks like success: draft and prefix untouched, the failure shows with a retry that works', async () => {
  let fail = true;
  const view = mount({ draft: '已有内容', answers: { selectTemplate: ({ kind, id }) => (fail ? new Error('Host 不在线') : { kind, id }) } });
  const result = await view.store.use(PPT, { actions: view.box.actions, draft: view.box.draft });
  assert.equal(result.written, false);
  assert.equal(view.box.draft, '已有内容');
  assert.equal(view.store.state.pick, null);
  const failed = byClass(view.toolbar(), 'sbs-tag-failed')[0];
  assert.match(words(failed), /模板未能带入：Host 不在线/);
  assert.equal(byClass(view.toolbar(), 'sbs-tag').filter(node => !node.props.className.includes('failed')).length, 0);
  fail = false;
  byClass(failed, 'sbs-link')[0].props.onClick();
  await tick();
  assert.equal(view.box.draft, `${PROMPTS.zh.ppt}已有内容`);
  assert.equal(view.store.state.failed, undefined);
  assert.equal(view.store.state.pick.id, 'default');
});

test('the prefix\'s ×: the template and its untouched prompt go, the user\'s text stays, the Host is told', async () => {
  const view = mount();
  await view.store.use(PPT, { actions: view.box.actions, draft: '重庆' });
  byClass(view.toolbar(), 'sbs-tag-x')[0].props.onClick();
  await tick();
  assert.equal(view.box.draft, '重庆');
  assert.equal(view.store.state.pick, null);
  assert.deepEqual(view.rpc.calls.at(-1), ['selectTemplate', { workspaceFileScopeId: 's1', kind: 'ppt', id: null }]);
});

test('pending use from the library page lands once, only in the Session it was meant for', async () => {
  const library = new InspirationLibrary();
  const rpc = rpcStub({ listStyles: LIST, templatePick: null, selectTemplate: ({ kind, id }) => ({ kind, id }) });
  const stores = createStylePickerStores(rpc);
  library.use(PPT, { target: 's1' });
  const other = mount({ sessionId: 's2', library, effects: true, stores, rpc, draft: 'x' });
  other.toolbar();
  await tick();
  assert.equal(other.box.draft, 'x', 'another Session leaves it');
  assert.ok(library.state.pending);
  const mine = mount({ sessionId: 's1', library, effects: true, stores, rpc, draft: '补充' });
  mine.toolbar();
  await tick();
  assert.equal(mine.box.draft, `${PROMPTS.zh.ppt}补充`);
  assert.equal(library.state.pending, null, 'taken once');
  mine.toolbar();
  await tick();
  assert.equal(mine.box.draft, `${PROMPTS.zh.ppt}补充`, 'not applied twice');
});

test('a turn: sending takes the prefix away; the Host forgets the template only when that turn ends', async () => {
  const view = mount();
  await view.store.use(PPT, { actions: view.box.actions, draft: '' });
  view.store.turnStarted();
  assert.equal(view.store.state.pick, null);
  assert.equal(view.rpc.calls.filter(([name, args]) => name === 'selectTemplate' && args.id === null).length, 0);
  await view.store.turnEnded();
  assert.deepEqual(view.rpc.calls.at(-1), ['selectTemplate', { workspaceFileScopeId: 's1', kind: 'ppt', id: null }]);
  // A use while running reaches the Host only when the turn ends; the draft is written now.
  view.store.turnStarted();
  await view.store.use(WORD, { actions: view.box.actions, draft: '' });
  assert.equal(view.box.draft, PROMPTS.zh.word);
  assert.equal(view.rpc.calls.filter(([name, args]) => name === 'selectTemplate' && args.id === 'gongwen').length, 0);
  await view.store.turnEnded();
  assert.deepEqual(view.rpc.calls.at(-1), ['selectTemplate', { workspaceFileScopeId: 's1', kind: 'word', id: 'gongwen' }]);
});

test('restore: a template the Host still holds shows its prefix; its prompt can still be replaced; an unknown one is cleared', async () => {
  const view = mount({ answers: { templatePick: { kind: 'word', id: 'gongwen' } } });
  await view.store.restore('zh');
  assert.deepEqual(view.store.state.pick, WORD);
  assert.equal(view.store.state.inserted, PROMPTS.zh.word);
  const gone = mount({ answers: { templatePick: { kind: 'word', id: 'nope' } } });
  await gone.store.restore('zh');
  assert.equal(gone.store.state.pick, null);
  assert.deepEqual(gone.rpc.calls.at(-1), ['selectTemplate', { workspaceFileScopeId: 's1', kind: 'word', id: null }]);
});

test('Backspace at the very start removes the template; composing, modifiers or text before the caret never do', () => {
  assert.equal(plainBackspace({ key: 'Backspace' }), true);
  assert.equal(plainBackspace({ key: 'Backspace', isComposing: true }), false, 'IME composition');
  assert.equal(plainBackspace({ key: 'Backspace', keyCode: 229 }), false, 'IME (Safari/Chromium 229)');
  assert.equal(plainBackspace({ key: 'Backspace', metaKey: true }), false);
  assert.equal(plainBackspace({ key: 'Delete' }), false);
  const text = { nodeType: 3 };
  const root = { contains: node => node === text, ownerDocument: { createRange: () => {
    const range = { selectNodeContents() {}, setEnd(_node, offset) { range.offset = offset; }, toString: () => '重庆'.slice(0, range.offset) };
    return range;
  } } };
  const selection = offset => ({ rangeCount: 1, isCollapsed: true, getRangeAt: () => ({ startContainer: text, startOffset: offset }) });
  assert.equal(caretAtStart(root, selection(0)), true);
  assert.equal(caretAtStart(root, selection(1)), false);
  assert.equal(caretAtStart(root, { ...selection(0), isCollapsed: false }), false, 'a selection is a normal delete');
  assert.equal(caretAtStart(root, { rangeCount: 1, isCollapsed: true, getRangeAt: () => ({ startContainer: {}, startOffset: 0 }) }), false, 'caret in another field');
  assert.equal(caretAtStart(undefined, selection(0)), false);
});

test('focus: the keyboard goes to the nearest composer editor going up, root first, then Lexical', () => {
  const calls = [];
  const editor = { focus: () => calls.push('root'), __lexicalEditor: { focus: () => calls.push('lexical') } };
  const pane = { querySelector: selector => (selector.includes('lexical') ? editor : null) };
  const from = { parentElement: { querySelector: () => null, parentElement: pane } };
  assert.equal(focusComposer(from), true);
  assert.deepEqual(calls, ['root', 'lexical']);
  assert.equal(composerOf(from), editor);
  assert.equal(focusComposer({ parentElement: null }), false);
});

test('English: the prompt and the prefix labels are English', async () => {
  const view = mount({ t: en });
  await view.store.use(PPT, { actions: view.box.actions, draft: '', lang: 'en' });
  assert.equal(view.box.draft, PROMPTS.en.ppt);
  assert.equal(byClass(view.toolbar(), 'sbs-tag-x')[0].props['aria-label'], 'Remove template 蓝色品牌');
});

test('prefix styling: borderless blue text with the bulb, × only on hover; helpers kept for the recommendation strip', () => {
  assert.match(CSS, /\.sbs-tag\{[^}]*border:0[^}]*color:#245dcc/);
  assert.match(CSS, /\[data-ds-dark-theme\] \.sbs-tag\{color:#98baff\}/);
  assert.match(CSS, /\.sbs-tag:hover \.sbs-tag-x,\.sbs-tag:focus-within \.sbs-tag-x\{opacity:1;pointer-events:auto\}/);
  assert.equal(appendText('a', 'b'), 'a b');
  assert.equal(appendText('主题', '用'), '主题用');
});
