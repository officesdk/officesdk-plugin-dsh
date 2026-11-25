import test from 'node:test';
import assert from 'node:assert/strict';
import { INSPIRATION_PANEL_ID, INSPIRATION_PANEL_ORDER, InspirationLibrary, PENDING_MS, createInspirationLibrary, registerInspirationLibrary } from '../src/client/inspiration-library.mjs';
import { STYLE_PICKER_MESSAGES, STYLE_PICKER_NS, createStylePickerStores } from '../src/client/style-picker.mjs';
import { FILE_ICONS, OUTLINE_ICONS } from '../src/client/inspiration-icons.mjs';
import { CSS } from '../src/client/styles.mjs';

const zh = key => STYLE_PICKER_MESSAGES.zh[key];
const pack = (id, name, scenarios, previews = []) => ({ id, name: { zh: name, en: id }, summary: { zh: `${name}的摘要`, en: '' }, scenarios: { zh: scenarios, en: [] }, thumb: `data:image/png;base64,${id}`, previews });
const LIST = { styles: [pack('default', '蓝色品牌', ['业务汇报'], ['data:p2', 'data:p3']), pack('dark', '黑底画报', ['城市旅行', '业务汇报'])], defaultStyle: 'default', remembered: null, canRemember: true };

function rpcStub(answers = {}) {
  const calls = [];
  return { calls, call: async (_p, method, { args }) => {
    const name = method.split('/')[1];
    calls.push([name, args]);
    const value = typeof answers[name] === 'function' ? await answers[name](args) : answers[name];
    return value instanceof Error ? { status: 'failure', error: { message: value.message } } : { status: 'success', value: value ?? null };
  } };
}
const tick = async (n = 4) => { for (let i = 0; i < n; i++) await new Promise(resolve => setImmediate(resolve)); };
function fakeReact() {
  return {
    Fragment: 'fragment',
    useState: initial => [typeof initial === 'function' ? initial() : initial, () => {}],
    createElement: (type, props, ...children) => ({ type, props: props ?? {}, children: children.flat(Infinity).filter(child => child !== null && child !== undefined && child !== false) }),
    useRef: initial => ({ current: initial }),
    useEffect: () => {},
    useLayoutEffect: () => {},
    useSyncExternalStore: (subscribe, get) => { subscribe(() => {})(); return get(); },
  };
}
const expand = node => (node && typeof node === 'object' && typeof node.type === 'function' ? expand(node.type(node.props)) : node);
const all = node => { node = expand(node); return node && typeof node === 'object' ? [node, ...node.children.flatMap(all)] : []; };
const words = node => { node = expand(node); return typeof node === 'string' ? node : node && typeof node === 'object' ? node.children.map(words).join('') : ''; };
const byClass = (tree, name) => all(tree).filter(node => (node.props.className ?? '').split(' ').includes(name));

async function page({ sessionId = 's1', answers = {} } = {}) {
  const rpc = rpcStub({ listStyles: LIST, ...answers });
  const stores = createStylePickerStores(rpc);
  const library = new InspirationLibrary();
  const calls = [];
  const layout = { selectPanel: id => calls.push(['selectPanel', id]) };
  const selection = { getSnapshot: () => (sessionId ? { sessionId } : {}), subscribe: () => () => {} };
  const workspace = { selection, startSession: () => calls.push(['startSession']) };
  const parts = createInspirationLibrary({ React: fakeReact(), stores, library, layout: () => layout, workspace: () => workspace, document: {} });
  if (sessionId) await stores.get(sessionId).load();
  const render = () => expand({ type: parts.InspirationPage, props: { t: zh } });
  return { rpc, stores, library, calls, parts, render };
}
const cards = tree => byClass(tree, 'sbi-card').map(node => node.props['data-template']);

test('registration: the sidebar entry between 自动化任务 (10) and 更多, and the main page under the same key, inside effects', () => {
  const slots = [], effects = [];
  const ctx = { effect: fn => effects.push(fn), get: () => undefined, locale: { bind: () => key => zh(key) }, slots: { inject: (_n, fn) => fn(), register: (options, component) => { slots.push([options, component]); return () => {}; } } };
  registerInspirationLibrary(ctx, fakeReact(), { stores: createStylePickerStores(rpcStub()), library: new InspirationLibrary() });
  assert.equal(slots.length, 0);
  effects.forEach(run => run());
  assert.deepEqual(slots.map(([options]) => options.name), ['main', 'sidebar.panellist']);
  assert.deepEqual(slots[0][0], { name: 'main', key: INSPIRATION_PANEL_ID, locale: STYLE_PICKER_NS });
  const entry = slots[1][0];
  assert.equal(entry.id, INSPIRATION_PANEL_ID);
  assert.equal(entry.order, INSPIRATION_PANEL_ORDER);
  assert.ok(INSPIRATION_PANEL_ORDER > 10, 'after 自动化任务');
  assert.equal(entry.label(), '灵感库');
  const glyph = expand({ type: slots[1][1], props: { size: 18, active: false } });
  assert.equal(byClass(glyph, 'sbi-bulb')[0].props.width, 18);
  assert.equal(all(glyph).filter(node => node.type === 'path').length, 2, 'the bulb and its thin bolt');
});

test('window store: the mode switch never touches the filter; use follows the template kind, closes the detail and waits for its Session', () => {
  let now = 1000;
  const library = new InspirationLibrary({ now: () => now });
  library.setScope('业务汇报');
  library.setMode('word');
  assert.deepEqual([library.state.mode, library.state.scope], ['word', '业务汇报']);
  library.setMode('pdf');
  assert.equal(library.state.mode, 'word');
  const template = { kind: 'excel', id: 'blank', title: '空白表格' };
  library.openDetail(template);
  library.use(template, { target: 's1' });
  assert.deepEqual([library.state.mode, library.state.detail], ['excel', null]);
  assert.equal(library.takePending('s2'), undefined, 'not for another Session');
  assert.equal(library.takePending('s1'), template);
  assert.equal(library.takePending('s1'), undefined, 'once');
  library.use(template);
  now += PENDING_MS + 1;
  assert.equal(library.takePending('s9'), undefined, 'an old one expires');
  assert.equal(library.state.pending, null);
  library.use(template);
  assert.equal(library.takePending('new'), template, 'no target: the next Session that mounts takes it');
});

test('page: centred title without subtitle, the glass switch, 我的模板｜精选｜scenes; featured by default, every kind mixed', async () => {
  const view = await page();
  const tree = view.render();
  assert.equal(words(byClass(tree, 'sbi-title')[0]), '今天，你想从哪开始？');
  assert.equal(byClass(tree, 'sbi-back').length, 1);
  const radios = all(byClass(tree, 'sbi-switch')[0]).filter(node => node.props.role === 'radio');
  assert.deepEqual(radios.map(node => [node.props['data-kind'], node.props['aria-checked']]), [['word', false], ['excel', false], ['ppt', true]]);
  const ppt = radios[2];
  assert.equal(all(ppt).find(node => node.type === 'img').props.src, FILE_ICONS.ppt, 'selected: the colour original');
  assert.match(byClass(radios[0], 'sbi-type-outline')[0].props.style['--sbi-outline'], new RegExp(OUTLINE_ICONS.word.slice(0, 40).replace(/[+/]/g, '\\$&')), 'unselected: the outline');
  // Scenes come from every kind's categories, most used first (ties: first use).
  assert.deepEqual(byClass(tree, 'sbi-chip').map(words), ['我的模板', '精选', '业务汇报', '公文通知', '城市旅行']);
  assert.equal(byClass(tree, 'sbi-chip').find(node => node.props['aria-pressed']).props['data-scope'], 'featured');
  assert.deepEqual(cards(tree), ['word:gongwen', 'ppt:default', 'ppt:dark']);
  // The switch only sets the mode: the cards and the filter stay.
  view.library.setScope('业务汇报');
  view.library.setMode('word');
  assert.deepEqual(cards(view.render()), ['ppt:default', 'ppt:dark']);
  assert.equal(view.library.state.scope, '业务汇报');
});

test('cards: kind frame, the type mark the switch uses, 0.5px divider, the name; no tick, dot or selected frame; quick use has no arrow', async () => {
  const view = await page();
  const card = byClass(view.render(), 'sbi-card').find(node => node.props['data-template'] === 'ppt:default');
  assert.equal(byClass(card, 'sbi-cover')[0].props.style.aspectRatio, '16 / 9');
  assert.equal(all(byClass(card, 'sbi-cover')[0]).filter(node => node.type === 'img').length, 3, 'the first three pages, only the first shown');
  assert.equal(byClass(card, 'sbi-current').length, 1);
  assert.equal(byClass(card, 'sbi-type').length, 1);
  assert.equal(words(byClass(card, 'sbi-card-name')[0]), '蓝色品牌');
  assert.equal(words(byClass(card, 'sbi-quick-use')[0]), '添加并使用');
  for (const gone of ['sbs-tpl-check', 'sbs-tpl-badge']) assert.equal(byClass(card, gone).length, 0);
  const word = byClass(view.render(), 'sbi-card').find(node => node.props['data-template'] === 'word:gongwen');
  assert.equal(byClass(word, 'sbi-cover')[0].props.style.aspectRatio, '210 / 297');
  assert.match(CSS, /\.sbi-card-divider\{flex:0 0 \.5px;width:\.5px;height:\.9em;margin-inline:8px;background:currentColor;opacity:\.3\}/);
  assert.match(CSS, /\.sbi-card-title\{[^}]*font:400 12px\/14px/);
  assert.match(CSS, /\.sbi-card-name\{[^}]*text-overflow:ellipsis/);
  assert.match(CSS, /\.sbi-cover>img\{[^}]*transition:opacity 650ms/);
  assert.match(CSS, /\.sbi-switch\{[^}]*width:288px;height:40px;padding:4px;[^}]*border-radius:22px/);
  assert.match(CSS, /--sbi-card-width,280px/);
});

test('a card\'s body opens the detail with no side effect; its quick use and the detail\'s use go back to the Session with a pending use', async () => {
  const view = await page();
  const tree = view.render();
  const card = byClass(tree, 'sbi-card').find(node => node.props['data-template'] === 'ppt:dark');
  const before = view.rpc.calls.length;
  byClass(card, 'sbi-card-open')[0].props.onClick();
  assert.equal(view.library.state.detail.id, 'dark');
  assert.equal(view.rpc.calls.length, before, 'no Host call, no draft');
  assert.deepEqual(view.calls, []);
  const detail = byClass(view.render(), 'sbi-detail')[0];
  byClass(detail, 'sbi-detail-use')[0].props.onClick();
  assert.equal(view.library.state.detail, null);
  assert.deepEqual(view.library.state.pending.target, 's1');
  assert.equal(view.library.state.mode, 'ppt');
  assert.deepEqual(view.calls, [['selectPanel', null]]);
  view.library.state.pending = null;
  byClass(byClass(view.render(), 'sbi-card').find(node => node.props['data-template'] === 'word:gongwen'), 'sbi-quick-use')[0].props.onClick();
  assert.equal(view.library.state.pending.template.id, 'gongwen');
  assert.equal(view.library.state.mode, 'word');
});

test('no Session on screen: the PPT list says why, a use opens a new conversation and waits for it', async () => {
  const view = await page({ sessionId: null });
  const tree = view.render();
  assert.match(words(byClass(tree, 'sbi-state')[0]), /打开一个工作区/);
  assert.deepEqual(cards(tree), ['word:gongwen']);
  byClass(byClass(tree, 'sbi-card')[0], 'sbi-quick-use')[0].props.onClick();
  assert.deepEqual(view.calls, [['startSession'], ['selectPanel', null]]);
  assert.equal(view.library.state.pending.target, null);
});

test('detail: category, title, summary, carousel (arrows only with pages, counter), example page 2, numbered thumbnails, one footer action', async () => {
  const view = await page();
  const template = { kind: 'ppt', id: 'default', title: '蓝色品牌', summary: '摘要', categories: ['业务汇报'], pages: ['data:1', 'data:2', 'data:3'] };
  const detail = expand({ type: view.parts.TemplateDetail, props: { template, onClose: () => {}, onUse: () => {}, t: zh } });
  const dialog = all(detail).find(node => node.type === 'dialog');
  assert.equal(words(byClass(dialog, 'sbi-detail-category')[0]), '业务汇报');
  assert.equal(words(all(dialog).find(node => node.type === 'h2')), '蓝色品牌');
  assert.equal(byClass(dialog, 'sbi-carousel-arrow').length, 2);
  assert.equal(byClass(dialog, 'sbi-carousel-arrow').find(node => node.props.className.includes('sbi-prev')).props.disabled, true, 'first page: back disabled');
  assert.equal(words(byClass(dialog, 'sbi-carousel-counter')[0]), '1 / 3');
  assert.equal(all(byClass(dialog, 'sbi-feature')[0]).find(node => node.type === 'img').props.src, 'data:2');
  assert.deepEqual(byClass(dialog, 'sbi-thumb-badge').map(words), ['01', '02', '03']);
  assert.match(words(byClass(dialog, 'sbi-pages-heading')[0]), /全部幻灯片3 张幻灯片/);
  const buttons = all(byClass(dialog, 'sbi-detail-foot')[0]).filter(node => node.type === 'button');
  assert.deepEqual(buttons.map(words), ['添加并使用']);
  assert.ok(!/下载|复制/.test(words(dialog)), 'no download or name copy');
  const single = expand({ type: view.parts.TemplateDetail, props: { template: { ...template, pages: ['data:1'] }, onClose: () => {}, onUse: () => {}, t: zh } });
  assert.equal(byClass(single, 'sbi-carousel-arrow').length, 0, 'one page: no arrows');
  assert.equal(all(byClass(single, 'sbi-feature')[0]).find(node => node.type === 'img').props.src, 'data:1');
  // Escape inside the detail never reaches the page; cancel and the backdrop close it; content clicks do not.
  let closed = 0;
  const live = all(expand({ type: view.parts.TemplateDetail, props: { template, onClose: () => { closed++; }, onUse: () => {}, t: zh } })).find(node => node.type === 'dialog');
  let stopped = false;
  live.props.onKeyDown({ key: 'Escape', stopPropagation: () => { stopped = true; } });
  assert.equal(stopped, true);
  live.props.onCancel({ preventDefault() {} });
  const self = {};
  live.props.onClick({ target: self, currentTarget: self });
  live.props.onClick({ target: {}, currentTarget: self });
  assert.equal(closed, 2);
});

test('lightbox: arrows, ←/→/Home/End within bounds, the counter; keys stop there; Esc closes only it', async () => {
  const view = await page();
  const template = { kind: 'word', id: 'basic', title: '普通 A4', pages: ['a', 'b', 'c'] };
  const moves = [];
  let closed = 0;
  const box = all(expand({ type: view.parts.PageLightbox, props: { template, index: 1, onChange: next => moves.push(next), onClose: () => { closed++; }, t: zh } })).find(node => node.type === 'dialog');
  const key = name => { let stopped = false; box.props.onKeyDown({ key: name, preventDefault() {}, stopPropagation: () => { stopped = true; } }); return stopped; };
  assert.equal(key('ArrowLeft'), true);
  key('ArrowRight'); key('Home'); key('End');
  assert.deepEqual(moves, [0, 2, 0, 2]);
  assert.equal(words(byClass(box, 'sbi-lightbox-counter')[0]), '2 / 3');
  box.props.onCancel({ preventDefault() {} });
  assert.equal(closed, 1);
  const single = expand({ type: view.parts.PageLightbox, props: { template: { ...template, pages: ['a'] }, index: 0, onChange() {}, onClose() {}, t: zh } });
  assert.ok(byClass(single, 'sbi-lightbox-arrow').every(node => node.props.hidden), 'one page: no arrows');
});

test('my templates: the create entry says from the start the service is not connected; a good file ends at "not connected"', async () => {
  const view = await page();
  view.library.setScope('mine');
  const tree = view.render();
  assert.match(words(byClass(tree, 'sbi-create')[0]), /创建我的模板.*上传与制作服务暂未接入/);
  assert.deepEqual(view.library.receiveFiles('ppt', [{ name: 'a.pptx', size: 3 }]), { state: 'unavailable', fileName: 'a.pptx' });
  assert.deepEqual(view.library.receiveFiles('ppt', [{ name: 'a.xlsx', size: 3 }]), { state: 'error', error: 'type' });
});

test('footer: the default box works only for a PPT the current Session carries; Word and Excel say theirs is not connected', async () => {
  const view = await page({ answers: { selectTemplate: ({ kind, id }) => ({ kind, id }), rememberStyle: ({ style }) => ({ remembered: style }) } });
  const box = () => all(byClass(view.render(), 'sbi-foot')[0]).find(node => node.type === 'input');
  assert.equal(box().props.disabled, true);
  const store = view.stores.get('s1');
  await store.use({ kind: 'ppt', id: 'default', title: '蓝色品牌' }, { actions: { setDraft() {} }, draft: '' });
  assert.equal(box().props.disabled, false);
  box().props.onChange({ target: { checked: true } });
  await tick();
  assert.equal(view.stores.get('s1').state.list.remembered, 'default');
  await store.use({ kind: 'word', id: 'gongwen', title: '公文版式' }, { actions: { setDraft() {} }, draft: '' });
  assert.match(words(byClass(view.render(), 'sbi-foot')[0]), /Word、Excel 的工作区默认模板暂未接入/);
});

test('theme: the page follows DSH\'s dark mark and the system only without DSH; reduced motion and transparency have their rules', () => {
  assert.match(CSS, /\[data-ds-dark-theme\] \.sbi-page,\[data-ds-dark-theme\] \.sbi-detail,\[data-ds-dark-theme\] \.sbi-lightbox\{--sbi-bg:/);
  assert.match(CSS, /:root:not\(\[data-ds-theme-source\]\) body:not\(\[data-ds-dark-theme\]\) :is\(\.sbi-page/);
  assert.match(CSS, /@media \(prefers-reduced-motion:reduce\)\{\.sbi-switch-slider/);
  assert.match(CSS, /@media \(prefers-reduced-transparency:reduce\)/);
  assert.match(CSS, /\.sbi-detail\{[^}]*width:min\(770px,calc\(100vw - 48px\)\)[^}]*height:min\(720px,calc\(100dvh - 64px\)\)/);
  assert.match(CSS, /\.sbi-thumbs\{display:grid;grid-template-columns:repeat\(4,minmax\(0,1fr\)\)/);
  assert.match(CSS, /\.sbi-summary::before\{[^}]*width:2px/);
});

test('a use with no target lands only in the Session the main view shows; pressing the open entry again (click or keyboard) goes back', async () => {
  let current = 's2';
  const library = new InspirationLibrary({ currentSession: () => current });
  const template = { kind: 'ppt', id: 'default', title: '蓝色品牌' };
  library.use(template);
  assert.equal(library.takePending('s1'), undefined, 'another pane leaves it');
  assert.equal(library.takePending('s2'), template);
  library.use(template);
  current = undefined;
  assert.equal(library.takePending('s1'), template, 'unknown main Session: the first that mounts');
  const calls = [];
  const listeners = {};
  const button = { addEventListener: (type, fn) => { listeners[type] = fn; }, removeEventListener() {} };
  const React = { ...fakeReact(), useRef: initial => ({ current: initial === null ? { closest: () => button } : initial }), useEffect: run => { run(); } };
  const { PanelGlyph } = createInspirationLibrary({ React, stores: createStylePickerStores(rpcStub()), library, layout: () => ({ selectPanel: id => calls.push(id) }), workspace: () => undefined, document: {} });
  expand({ type: PanelGlyph, props: { size: 18, active: true } });
  assert.equal(listeners.pointerdown, undefined, 'no pointer-only step');
  listeners.click();
  await new Promise(resolve => setTimeout(resolve, 5));
  assert.deepEqual(calls, [null]);
  expand({ type: PanelGlyph, props: { size: 18, active: false } });
  listeners.click();
  await new Promise(resolve => setTimeout(resolve, 5));
  assert.deepEqual(calls, [null], 'opening it does not close it');
});

test('page: a Host that lists its templates (the public build) shows Word gongwen only, no Excel tab, card or create entry (261113)', async () => {
  const view = await page({ answers: { listStyles: { ...LIST, templates: { word: ['gongwen'], excel: [] } } } });
  const tree = view.render();
  const switcher = byClass(tree, 'sbi-switch')[0];
  assert.deepEqual(all(switcher).filter(node => node.props.role === 'radio').map(node => node.props['data-kind']), ['word', 'ppt']);
  assert.equal(switcher.props.style['--sbi-count'], 2);
  assert.deepEqual(cards(tree), ['word:gongwen', 'ppt:default', 'ppt:dark']);
  assert.ok(!byClass(tree, 'sbi-chip').map(words).includes('通用表格'), 'no Excel scene');
  // A mode the copy cannot offer (Excel, remembered from before) falls back to PPT, also for "my templates".
  view.library.setMode('excel');
  view.library.setScope('mine');
  const mine = view.render();
  assert.equal(all(byClass(mine, 'sbi-switch')[0]).find(node => node.props['aria-checked'] === true).props['data-kind'], 'ppt');
  assert.ok(all(mine).some(node => node.props?.accept?.includes('.pptx')), 'the create entry takes PPT files');
  // An older Host (no list): every kind, as before.
  const old = (await page()).render();
  assert.equal(byClass(old, 'sbi-switch')[0].props.style['--sbi-count'], 3);
});
