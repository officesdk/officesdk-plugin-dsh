import test from 'node:test';
import assert from 'node:assert/strict';
import { FILTER_ALL, columnsOf, createStyleCards, filterCards, galleryFilters, gridKey, moveFocus } from '../src/client/style-cards.mjs';
import { STYLE_PICKER_MESSAGES, STYLE_PICKER_NS, StylePickerStore, answerText, createStylePickerStores } from '../src/client/style-picker.mjs';
import { CSS } from '../src/client/styles.mjs';
import { STYLE_DOCK_ID, createStyleFlows, dockMode, registerStyleFlows, runningOf } from '../src/client/style-flows.mjs';

const zh = key => STYLE_PICKER_MESSAGES.zh[key];
const en = key => STYLE_PICKER_MESSAGES.en[key];
const card = (id, zhName, scenarios = [], extra = {}) => ({ id, name: { zh: zhName, en: id }, summary: { zh: `${zhName}的特点`, en: '' }, scenarios: { zh: scenarios, en: [] }, requires: { note: { zh: '', en: '' } }, thumb: `data:image/png;base64,${id}`, previews: [], ...extra });
const BLUE = card('default', '蓝色品牌', ['公司介绍', '业务汇报'], { previews: ['data:p1', 'data:p2', 'data:p3'] });
const DARK = card('second', '黑底画报', ['城市旅行', '业务汇报'], { requires: { note: { zh: '每份约 19 张照片', en: '' } } });
const INK = card('third', '水墨', ['文化']);
const LIST = { styles: [BLUE, DARK, INK], defaultStyle: 'default', remembered: null, canRemember: true };
const SUGGESTION = { id: 'g1', topic: '重庆三日游', candidates: [{ ...DARK, reason: '照片多的旅行主题' }, { ...BLUE, reason: '稳妥' }], defaultStyle: { id: 'default', name: { zh: '蓝色品牌', en: 'Blue Brand' } } };

/** DSH's inputActions over a draft; `submit` records the draft it would send. */
function draftBox(draft = '', caret = draft.length) {
  const box = { draft, rev: 0, calls: [], sent: [] };
  box.actions = {
    setDraft: text => { box.calls.push(['setDraft', text]); box.draft = text; box.rev++; },
    captureInsertion: () => ({ start: caret, end: caret, draftRev: box.rev }),
    insertText: (text, span) => { if (span.draftRev !== box.rev) return false; box.draft = box.draft.slice(0, span.start) + text + box.draft.slice(span.end); box.rev++; box.calls.push(['insertText', text]); return true; },
    submit: () => { box.sent.push(box.draft); box.calls.push(['submit']); },
  };
  return box;
}
function rpcStub(answers = {}) {
  const calls = [];
  return { calls, call: async (path, method, { args }) => {
    const name = method.split('/')[1];
    calls.push([name, args]);
    const answer = answers[name];
    const value = typeof answer === 'function' ? await answer(args) : answer;
    return value instanceof Error ? { status: 'failure', error: { code: 'officesdk/forbidden', message: value.message } } : { status: 'success', value };
  } };
}
function fakeReact() {
  return {
    Fragment: 'fragment',
    useState: initial => [typeof initial === 'function' ? initial() : initial, () => {}],
    createElement: (type, props, ...children) => ({ type, props: props ?? {}, children: children.flat(Infinity).filter(child => child !== null && child !== undefined && child !== false) }),
    useRef: initial => ({ current: initial }),
    useEffect: () => {},
    useSyncExternalStore: (subscribe, get) => { subscribe(() => {})(); return get(); },
  };
}
const expand = node => (node && typeof node === 'object' && typeof node.type === 'function' ? expand(node.type(node.props)) : node);
const all = node => { node = expand(node); return node && typeof node === 'object' ? [node, ...node.children.flatMap(all)] : []; };
const words = node => { node = expand(node); return typeof node === 'string' ? node : node && typeof node === 'object' ? node.children.map(words).join('') : ''; };
const byClass = (tree, name) => all(tree).filter(node => (node.props.className ?? '').split(' ').includes(name));
const tick = async (n = 3) => { for (let i = 0; i < n; i++) await new Promise(resolve => setImmediate(resolve)); };

test('cards: scenario filters, most common first; a filter keeps its cards, an unknown one keeps all', () => {
  assert.deepEqual(galleryFilters(LIST.styles, 'zh'), ['业务汇报', '公司介绍', '城市旅行', '文化']);
  assert.deepEqual(galleryFilters([card('x', 'x', [...'abcdefghij'])]).length, 8, 'at most eight');
  assert.deepEqual(galleryFilters([{ id: 'x', scenarios: { zh: [], en: ['Travel'] } }], 'zh'), ['Travel'], 'the other language when one is empty');
  assert.deepEqual(filterCards(LIST.styles, '业务汇报').map(item => item.id), ['default', 'second']);
  assert.deepEqual(filterCards(LIST.styles, FILTER_ALL).length, 3);
  assert.deepEqual(filterCards(LIST.styles, '没有这个').length, 3);
});

test('cards: keyboard moves across the measured columns', () => {
  assert.equal(gridKey('ArrowDown', 0, 5, 3), 3);
  assert.equal(gridKey('ArrowRight', 4, 5), 4);
  assert.equal(gridKey('Tab', 0, 5), undefined);
  assert.equal(columnsOf([{ offsetTop: 0 }, { offsetTop: 0 }, { offsetTop: 0 }, { offsetTop: 90 }]), 3);
  assert.equal(columnsOf([{}]), 2, 'nothing laid out: two');
  const focused = [];
  const cards = [0, 1, 2, 3].map(i => ({ offsetTop: i < 3 ? 0 : 90, focus: () => focused.push(i) }));
  const prevented = [];
  assert.equal(moveFocus({ key: 'ArrowDown', preventDefault: () => prevented.push(1) }, { querySelectorAll: () => cards }, { activeElement: cards[0] }), true);
  assert.deepEqual([focused, prevented.length], [[3], 1]);
  assert.equal(moveFocus({ key: 'a' }, { querySelectorAll: () => cards }, {}), false);
});

test('cards rendered: the full card has thumbnail, page previews, name, summary, scenarios and prerequisite; compact shows the reason', () => {
  const { StyleCard, StyleGallery } = createStyleCards(fakeReact());
  const full = expand({ type: StyleCard, props: { card: BLUE, lang: 'zh', t: zh } });
  assert.deepEqual(byClass(full, 'sbs-preview').map(node => node.props.src), ['data:p1', 'data:p2', 'data:p3']);
  assert.match(words(full), /蓝色品牌default蓝色品牌的特点适用：公司介绍、业务汇报/);
  assert.match(words(expand({ type: StyleCard, props: { card: DARK, lang: 'zh', t: zh } })), /每份约 19 张照片$/);
  assert.equal(byClass(expand({ type: StyleCard, props: { card: DARK, lang: 'zh', t: zh } }), 'sbs-previews').length, 0, 'no previews row without previews');
  const compact = expand({ type: StyleCard, props: { card: SUGGESTION.candidates[0], lang: 'zh', t: zh, variant: 'compact' } });
  assert.equal(words(compact), '黑底画报照片多的旅行主题');
  assert.equal(byClass(compact, 'sbs-card-compact').length, 1);
  // The gallery: chips only with two cards or more; a chip filters through the surface's handler.
  const filters = [];
  const gallery = expand({ type: StyleGallery, props: { cards: LIST.styles, lang: 'zh', t: zh, filter: '城市旅行', onFilter: value => filters.push(value) } });
  assert.deepEqual(byClass(gallery, 'sbs-chip').map(words), ['全部', '业务汇报', '公司介绍', '城市旅行', '文化']);
  assert.deepEqual(byClass(gallery, 'sbs-card').map(node => node.props['data-style']), ['second']);
  byClass(gallery, 'sbs-chip')[0].props.onClick();
  assert.deepEqual(filters, [FILTER_ALL]);
  assert.equal(byClass(expand({ type: StyleGallery, props: { cards: [BLUE], lang: 'zh', t: zh, onFilter: () => {} } }), 'sbs-chip').length, 0);
});

test('words: answers name the pack the way SKILL.md reads it', () => {
  assert.equal(answerText(DARK, 'zh'), '用 second（黑底画报）样式');
  assert.equal(answerText(BLUE, 'zh', { fallback: true }), '都不要，用默认样式 default（蓝色品牌）');
  assert.equal(answerText(BLUE, 'en', { fallback: true }), 'None of these; use the default style default (default)');
});

test('store: an idle fetch shows the recommendation; dismiss ends it here and on the Host; a failed fetch keeps what is shown', async () => {
  let answer = SUGGESTION;
  const rpc = rpcStub({ styleSuggestion: () => answer, dismissStyleSuggestion: ({ id }) => ({ dismissed: id === 'g1' }) });
  const store = new StylePickerStore({ rpc, sessionId: 's1' });
  assert.equal(await store.fetchSuggestion(), SUGGESTION);
  assert.equal(store.state.suggestion.id, 'g1');
  await store.setSuggestView('all');
  assert.equal(store.state.suggestView, 'all');
  answer = new Error('offline');
  await store.fetchSuggestion();
  assert.equal(store.state.suggestion.id, 'g1', 'kept');
  assert.equal(store.state.suggestView, 'all', 'the same recommendation keeps its view');
  assert.equal(await store.dismissSuggestion(), true);
  assert.equal(store.state.suggestion, undefined);
  assert.deepEqual(rpc.calls.filter(([name]) => name === 'dismissStyleSuggestion'), [['dismissStyleSuggestion', { workspaceFileScopeId: 's1', id: 'g1' }]]);
  answer = null;
  assert.equal(await store.fetchSuggestion(), undefined);
  assert.equal(await store.dismissSuggestion(), false, 'nothing to dismiss');
});

test('store: a pick on an empty draft is sent as the answer; text or an attachment in the draft, or a running turn, only writes it', async () => {
  const cases = [
    [{ draft: '', attachmentIds: [] }, false, 'sent', '用 second（黑底画报）样式', ['用 second（黑底画报）样式']],
    [{ draft: '  ', attachmentIds: ['a1'] }, false, 'set', '用 second（黑底画报）样式', []],
    [{ draft: '', attachmentIds: [] }, true, 'set', '用 second（黑底画报）样式', []],
    [{ draft: '照片我自己给', attachmentIds: [] }, false, 'inserted', '照片我自己给用 second（黑底画报）样式', []],
  ];
  for (const [input, running, how, draft, sent] of cases) {
    const rpc = rpcStub({ styleSuggestion: SUGGESTION, dismissStyleSuggestion: { dismissed: true } });
    const store = new StylePickerStore({ rpc, sessionId: 's1' });
    await store.fetchSuggestion();
    const box = draftBox(input.draft);
    const deferred = [];
    assert.equal(store.answer(DARK, { actions: box.actions, input, lang: 'zh', running, defer: run => deferred.push(run) }), how);
    deferred.forEach(run => run());
    assert.equal(box.draft, draft);
    assert.deepEqual(box.sent, sent);
    assert.equal(store.state.suggestion, undefined, 'ended either way');
    await tick();
    assert.ok(rpc.calls.some(([name]) => name === 'dismissStyleSuggestion'));
  }
  // The send waits for the editor to take the text: setDraft first, submit later.
  const store = new StylePickerStore({ rpc: rpcStub({ styleSuggestion: SUGGESTION }), sessionId: 's1' });
  await store.fetchSuggestion();
  const box = draftBox('');
  const deferred = [];
  store.answer(BLUE, { actions: box.actions, input: { draft: '' }, lang: 'zh', fallback: true, defer: run => deferred.push(run) });
  assert.deepEqual(box.calls, [['setDraft', '都不要，用默认样式 default（蓝色品牌）']]);
  deferred[0]();
  assert.deepEqual(box.sent, ['都不要，用默认样式 default（蓝色品牌）']);
});

test('dock: only the recommendation (the library is the sidebar page now), shown only while idle; running is read from DSH status, else the Session', () => {
  assert.equal(dockMode({ suggestion: SUGGESTION, gallery: 'switch', suggestView: 'cards' }, false), 'suggest', 'no library in the dock any more');
  assert.equal(dockMode({ suggestion: SUGGESTION, gallery: false, suggestView: 'cards' }, false), 'suggest');
  assert.equal(dockMode({ suggestion: SUGGESTION, suggestView: 'all' }, undefined), 'suggest-all');
  assert.equal(dockMode({ suggestion: SUGGESTION, gallery: false }, true), 'none', 'a running turn hides the recommendation');
  assert.equal(dockMode({ gallery: false }, false), 'none');
  const statuses = new Map([['s1', { running: true }]]);
  assert.equal(runningOf({ sessionId: 's1', useSessionStatus: select => select(statuses) }), true);
  assert.equal(runningOf({ sessionId: 's2', useSessionStatus: select => select(statuses), session: { running: false } }), false);
  assert.equal(runningOf({ sessionId: 's2' }), undefined);
});

function mount({ answers = {}, input = { draft: '', attachmentIds: [] }, running = false, t = zh } = {}) {
  const rpc = rpcStub({ listStyles: LIST, styleSuggestion: null, dismissStyleSuggestion: { dismissed: true }, ...answers });
  const stores = createStylePickerStores(rpc);
  const deferred = [];
  const { StyleDock } = createStyleFlows({ React: fakeReact(), stores, document: {}, defer: run => deferred.push(run) });
  const box = draftBox(input.draft);
  const props = () => ({ sessionId: 's1', t, inputActions: box.actions, useInput: select => select({ ...input, draft: box.draft }), useSessionStatus: select => select(new Map([['s1', { running }]])) });
  return { rpc, store: stores.get('s1'), box, deferred, dock: () => expand({ type: StyleDock, props: props() }) };
}

test('flow 2: the strip shows the recommended cards with reasons, the default and see-all; a pick is sent and ends it', async () => {
  const view = mount({ answers: { styleSuggestion: SUGGESTION } });
  await view.store.fetchSuggestion();
  const dock = view.dock();
  assert.equal(dock.props['data-shimo-style-dock'], 'suggest');
  assert.match(words(byClass(dock, 'sbs-head')[0]), /按你的内容推荐的版式「重庆三日游」/);
  const cards = byClass(dock, 'sbs-card');
  assert.deepEqual(cards.map(node => [node.props['data-style'], words(node)]), [['second', '黑底画报照片多的旅行主题'], ['default', '蓝色品牌稳妥']]);
  assert.ok(byClass(cards[0], 'sbs-thumb').length, 'with its thumbnail');
  assert.deepEqual(byClass(dock, 'sbs-action').map(words), ['都不要，用默认（蓝色品牌）', '看全部样式']);
  cards[0].props.onClick();
  view.deferred.forEach(run => run());
  assert.deepEqual(view.box.sent, ['用 second（黑底画报）样式']);
  assert.equal(view.dock(), null, 'gone at once');
  await tick();
  assert.deepEqual(view.rpc.calls.filter(([name]) => name === 'dismissStyleSuggestion').map(([, args]) => args.id), ['g1']);
});

test('flow 2: "none of these" sends the default; "see all" swaps in the full gallery and a pick there answers too; × only closes', async () => {
  const fallback = mount({ answers: { styleSuggestion: SUGGESTION } });
  await fallback.store.fetchSuggestion();
  byClass(fallback.dock(), 'sbs-action')[0].props.onClick();
  fallback.deferred.forEach(run => run());
  assert.deepEqual(fallback.box.sent, ['都不要，用默认样式 default（蓝色品牌）']);

  const browse = mount({ answers: { styleSuggestion: SUGGESTION } });
  await browse.store.fetchSuggestion();
  byClass(browse.dock(), 'sbs-action')[1].props.onClick();
  await tick();
  const full = browse.dock();
  assert.equal(full.props['data-shimo-style-dock'], 'suggest');
  assert.deepEqual(byClass(full, 'sbs-card').map(node => node.props['data-style']), ['default', 'second', 'third'], 'every style, same cards');
  assert.deepEqual(byClass(full, 'sbs-action').map(words), ['都不要，用默认（蓝色品牌）', '返回推荐']);
  byClass(full, 'sbs-card')[2].props.onClick();
  browse.deferred.forEach(run => run());
  assert.deepEqual(browse.box.sent, ['用 third（水墨）样式']);

  const closing = mount({ answers: { styleSuggestion: SUGGESTION } });
  await closing.store.fetchSuggestion();
  byClass(closing.dock(), 'sbs-x')[0].props.onClick();
  assert.equal(closing.dock(), null);
  assert.deepEqual(closing.box.sent, [], 'closing answers nothing');
});

test('flow 2: while a turn runs the strip is hidden; English labels', async () => {
  const running = mount({ answers: { styleSuggestion: SUGGESTION }, running: true });
  await running.store.fetchSuggestion();
  assert.equal(running.dock(), null);
  const english = mount({ answers: { styleSuggestion: SUGGESTION }, t: en });
  await english.store.fetchSuggestion();
  assert.deepEqual(byClass(english.dock(), 'sbs-action').map(words), ['None of these, use the default (Blue Brand)', 'See all styles']);
});

test('registration: only the dock in conversation.input.dock (no first-screen chip any more), with the picker\'s dictionary, inside effects', () => {
  const slots = [], effects = [];
  const ctx = { effect: fn => effects.push(fn), slots: { inject: (_name, fn) => fn(), register: (options, component) => { slots.push([options, component]); return () => {}; } } };
  registerStyleFlows(ctx, fakeReact(), { stores: createStylePickerStores(rpcStub()) });
  assert.equal(slots.length, 0);
  effects.forEach(run => run());
  assert.deepEqual(slots.map(([options]) => options), [
    { name: 'conversation.input.dock', id: STYLE_DOCK_ID, order: 100, locale: STYLE_PICKER_NS },
  ]);
});

test('a dozen styles and more: the recommendation\'s "all styles" view keeps at most eight occasion filters plus All', () => {
  const OCCASIONS = ['业务汇报', '城市旅行', '年度报告', '公司介绍', '品牌发布', '摄影作品', '教学课件', '环保公益', '科技产品', '文化'];
  // Fourteen packs: six occasions on two packs each, four on one each — the two seen last are left out.
  const cards = Array.from({ length: 14 }, (_, index) => card(`pack-${index}`, `样式${index}`, [OCCASIONS[index % 8], ...(index >= 12 ? [OCCASIONS[index - 4]] : [])]));
  const filters = galleryFilters(cards);
  assert.equal(filters.length, 8, 'eight at most');
  assert.ok(!filters.includes('文化') && !filters.includes('科技产品'), 'the rarest are left out');
  assert.deepEqual(filterCards(cards, '业务汇报').map(item => item.id), ['pack-0', 'pack-8']);
  assert.match(CSS, /\.sbs-dock \.sbs-grid\{[^}]*overflow:auto/, 'the strip\'s cards scroll under a fixed head');
});
