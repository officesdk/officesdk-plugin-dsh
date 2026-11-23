import test from 'node:test';
import assert from 'node:assert/strict';
import { NEW_SLIDES_MESSAGE, PPTX_ALL_OPS, PPTX_HELD_OPS, PPTX_LIMITS, PPTX_OPS, countNewSlides, describePptxSteps, summarizeDeck, validatePptxOps as check } from '../src/index.mjs';

// Held operations are implemented and tested here; the model is not offered them (see the last test).
const validatePptxOps = (ops, deck) => check(ops, deck, { enabled: PPTX_ALL_OPS });
import { LABELS, NORMAL, SORTER, inspectDeck, playsTogether, runSteps } from '../src/browser/presentation-agent.mjs';
import { createKitPresentationAgent } from '../src/browser/kit-agent.mjs';

/**
 * A recording stand-in for the editor's PowerPoint JS API: enough of `PowerPoint.run` for the generation
 * steps, with every mutating call written to `calls` in order.
 */
function fakePowerPoint({ attention = true, inherited = false, failText, layouts = [{ id: 'L1', name: 'Title Slide', placeholders: [['Placeholder', 'CenterTitle'], ['Placeholder', 'Subtitle']] }, { id: 'L2', name: 'Blank', placeholders: [] }] } = {}) {
  const calls = [];
  let next = 100;
  const slides = [{ id: 's1', shapes: [] }];
  const slideById = id => { const slide = slides.find(item => item.id === id); if (!slide) throw new Error(`no slide ${id}`); return slide; };
  let selected = [];
  const textRange = shape => ({
    get text() { return shape.text; }, set text(value) { if (failText && value === failText) throw Object.assign(new Error('the editor refused that text'), { code: 'editor-failed' }); shape.text = value; calls.push(['text', shape.id, value]); },
    font: new Proxy({}, { set(_, key, value) { calls.push(['font', shape.id, key, value]); return true; } }),
    paragraphFormat: { set horizontalAlignment(value) { calls.push(['align', shape.id, value]); } },
    load() {},
  });
  const shapeApi = (slide, shape) => ({
    get id() { return shape.id; }, get type() { return shape.type; },
    get left() { return shape.left; }, get top() { return shape.top; }, get width() { return shape.width; }, get height() { return shape.height; },
    load() {},
    textFrame: { textRange: textRange(shape) },
    fill: { clear: () => calls.push(['fill.clear', shape.id]), setSolidColor: color => calls.push(['fill', shape.id, color]) },
    lineFormat: { set visible(value) { calls.push(['line.visible', shape.id, value]); }, set color(value) { calls.push(['line.color', shape.id, value]); } },
    placeholderFormat: { load() {}, get type() { return shape.placeholder; } },
    delete: () => { slide.shapes.splice(slide.shapes.indexOf(shape), 1); calls.push(['delete', shape.id]); },
  });
  const add = (slide, record, call) => { const shape = { id: `n${next++}`, text: '', ...record }; slide.shapes.push(shape); calls.push([call, slide.id, shape.id]); return shapeApi(slide, shape); };
  const shapesApi = slide => ({
    load() {}, get items() { return slide.shapes.map(shape => shapeApi(slide, shape)); },
    getItem: id => { const shape = slide.shapes.find(item => item.id === id); if (!shape) throw new Error(`no shape ${id}`); return shapeApi(slide, shape); },
    addTextBox: (text, box) => add(slide, { type: 'TextBox', text, ...box }, 'addTextBox'),
    addGeometricShape: (kind, options) => { calls.push(['addGeometricShape.args', kind, options]); return add(slide, { type: 'GeometricShape', ...options }, 'addGeometricShape'); },
    addImage: (data, box) => { calls.push(['addImage.args', data, box]); return add(slide, { type: 'Image', ...box }, 'addImage'); },
    addTable: (rows, cols, options) => { calls.push(['addTable.args', rows, cols, options]); return add(slide, { type: 'Table', ...options }, 'addTable'); },
    addChart: options => { calls.push(['addChart.args', options]); return add(slide, { type: 'Chart', ...options }, 'addChart'); },
  });
  const layoutShapes = layout => ({ load() {}, items: (layout?.placeholders ?? []).map(([type, placeholder, geometry = {}]) => ({ type, ...geometry, placeholderFormat: { load() {}, type: placeholder } })) });
  const slideApi = slide => ({
    get id() { return slide.id; },
    get layout() { return { shapes: layoutShapes(slide.layout) }; },
    shapes: shapesApi(slide),
    background: { fill: { setSolidFill: options => calls.push(['background', slide.id, options.color]) } },
    moveTo: index => { slides.splice(slides.indexOf(slide), 1); slides.splice(index, 0, slide); calls.push(['moveTo', slide.id, index]); },
    delete: () => { slides.splice(slides.indexOf(slide), 1); calls.push(['deleteSlide', slide.id]); },
  });
  const presentation = {
    slides: {
      load() {}, get items() { return slides.map(slideApi); },
      getItem: id => slideApi(slideById(id)),
      add: options => {
        // Like the editor, a slide added without a layout takes the first one.
        const layout = options?.layoutId ? layouts.find(item => item.id === options.layoutId) : layouts[0];
        // `inherited`: like the editor, a new page's placeholders read 0,0,0,0 (their place is the layout's).
        const box = inherited ? { left: 0, top: 0, width: 0, height: 0 } : { left: 10, top: 10, width: 100, height: 20 };
        const slide = { id: `s${next++}`, layout, shapes: (layout?.placeholders ?? []).map(([type, placeholder]) => ({ id: `p${next++}`, type, placeholder, text: '', ...box })) };
        slides.push(slide);
        calls.push(['addSlide', slide.id, options?.layoutId ?? null]);
      },
    },
    slideMasters: { load() {}, items: [{ id: 'M1', name: 'Master', layouts: { load() {}, items: layouts.map(({ id, name, placeholders }) => ({ id, name,
      shapes: { load() {}, items: placeholders.map(([type, placeholder]) => ({ type, placeholderFormat: { load() {}, type: placeholder } })) } })) } }] },
    pageSetup: { load() {}, slideWidth: 960, slideHeight: 540 },
    setSelectedSlides: ids => { selected = [...ids]; calls.push(['select', ids[0]]); },
    getSelectedSlides: () => ({ load() {}, items: selected.map(id => ({ id })) }),
    getSelectedShapes: () => ({ load() {}, items: [] }),
    ...(attention && { focusAttention: rect => calls.push(['attention', rect && { ...rect }]) }),
  };
  // Each run yields once, as the editor's does, so runs that were not queued would overlap; `overlap` counts that.
  let inside = 0;
  const state = { overlap: 0 };
  const api = { PowerPoint: { run: async body => {
    if (++inside > 1) state.overlap++;
    try { const result = await body({ presentation, sync: async () => {} }); await new Promise(resolve => setImmediate(resolve)); return result; }
    finally { inside--; }
  } } };
  return { api, calls, slides, state, select: id => { selected = [id]; } };
}

const blank = { slides: [{ id: 's1', shapes: [] }], layouts: [{ id: 'L1', name: 'Title Slide' }, { id: 'L2', name: 'Blank' }] };
const overlay = () => { const log = []; return { log, frame: (rect, options) => log.push(['frame', rect ? 'on' : 'off', options?.tone]), write: () => {}, say: text => log.push(['say', text]), busy: () => {} }; };
const quick = { animate: false, wait: async () => {} };

test('a batch builds a page and fills it through refs; the result maps each ref to the id the editor gave it', async () => {
  const ops = [
    { op: 'addSlide', ref: 'p2', layout: 'Blank' },
    { op: 'setBackground', slideRef: 'p2', color: '#FFFFFF' },
    { op: 'addShape', slideRef: 'p2', ref: 'bar', shape: 'Rectangle', left: 0, top: 0, width: 960, height: 8, fill: '#2826C9', line: 'none' },
    { op: 'addTextBox', slideRef: 'p2', ref: 'title', text: '三个结论', left: 40, top: 40, width: 600, height: 60, size: 32, bold: true, color: '#2826C9' },
    { op: 'setText', slideRef: 'p2', shapeRef: 'title', text: '三个关键结论' },
    { op: 'addTable', slideRef: 'p2', ref: 'grid', values: [['项目', '数值'], ['收入', 12]], left: 40, top: 140, width: 400, height: 120 },
    { op: 'addChart', slideRef: 'p2', chartType: 'column', categories: ['Q1', 'Q2'], series: [{ name: '收入', values: [1, 2] }], title: '季度收入', legend: 'none', left: 480, top: 140, width: 400, height: 300 },
  ];
  const { steps, danger } = validatePptxOps(ops, blank);
  assert.equal(danger, false);
  assert.equal(steps[0].layoutId, 'L2');
  assert.deepEqual(steps[5].values, [['项目', '数值'], ['收入', '12']], 'numbers in a table become text');
  assert.deepEqual(steps[6].data, [['', 'Q1', 'Q2'], ['收入', 1, 2]], 'the chart gets a header row and one row per series');
  const { api, calls } = fakePowerPoint();
  const refs = await runSteps(api, steps, { ...quick, overlay: overlay(), document: {} });
  const slideId = refs.p2;
  assert.ok(slideId && refs.bar && refs.title && refs.grid, 'every ref resolved');
  assert.deepEqual(calls.find(call => call[0] === 'addSlide'), ['addSlide', slideId, 'L2']);
  assert.deepEqual(calls.find(call => call[0] === 'background'), ['background', slideId, '#FFFFFF']);
  assert.deepEqual(calls.find(call => call[0] === 'addGeometricShape.args'), ['addGeometricShape.args', 'Rectangle', { left: 0, top: 0, width: 960, height: 8, fillColor: '#2826C9', lineVisible: false }]);
  assert.ok(calls.some(call => call[0] === 'font' && call[1] === refs.title && call[2] === 'size' && call[3] === 32), 'inline font reaches the new text box');
  assert.equal(calls.filter(call => call[0] === 'text' && call[1] === refs.title).at(-1)[2], '三个关键结论', 'a later step edits the box through its ref');
  assert.deepEqual(calls.find(call => call[0] === 'addTable.args').slice(1, 3), [2, 2]);
  const chart = calls.find(call => call[0] === 'addChart.args')[1];
  assert.equal(chart.chartType, 'column');
  assert.equal(chart.legendVisible, false);
  assert.equal(chart.title, '季度收入');
});

test('placeholders of a layout are filled by type; a missing one fails the edit', async () => {
  const { steps } = validatePptxOps([{ op: 'addSlide', ref: 'cover', layout: 'L1', placeholders: { title: '年度汇报', subtitle: '2026' } }], blank);
  const { api, calls } = fakePowerPoint();
  const refs = await runSteps(api, steps, { ...quick, overlay: overlay(), document: {} });
  const texts = calls.filter(call => call[0] === 'text').map(call => call[2]);
  assert.deepEqual(texts, ['年度汇报', '2026']);
  assert.ok(refs.cover);
  // A placeholder the step leaves empty is removed from the new page.
  const partial = fakePowerPoint();
  const only = await runSteps(partial.api, validatePptxOps([{ op: 'addSlide', ref: 'c', layout: 'L1', placeholders: { title: '只有标题' } }], blank).steps, { ...quick, document: {} });
  assert.equal(partial.calls.filter(call => call[0] === 'delete').length, 1);
  assert.deepEqual(partial.slides.find(slide => slide.id === only.c).shapes.map(shape => shape.text), ['只有标题']);
  // Empty text counts as not filled; a slide added without a layout keeps its placeholders.
  const empty = fakePowerPoint();
  const blankTitle = await runSteps(empty.api, validatePptxOps([{ op: 'addSlide', ref: 'e', layout: 'L1', placeholders: { title: '', subtitle: '副标题' } }], blank).steps, { ...quick, document: {} });
  assert.deepEqual(empty.slides.find(slide => slide.id === blankTitle.e).shapes.map(shape => shape.text), ['副标题']);
  const plain = fakePowerPoint();
  await runSteps(plain.api, validatePptxOps([{ op: 'addSlide' }], blank).steps, { ...quick, document: {} });
  assert.equal(plain.calls.filter(call => call[0] === 'delete').length, 0);
  assert.equal(plain.slides.at(-1).shapes.length, 2, 'its two placeholders stay');
  const missing = validatePptxOps([{ op: 'addSlide', layout: 'Blank', placeholders: { title: 'x' } }], blank).steps;
  await assert.rejects(runSteps(fakePowerPoint().api, missing, { ...quick, document: {} }), /no title placeholder/);
});

test('a new page\'s placeholders are outlined where the layout puts them; a box without size is not outlined', async () => {
  const layouts = [{ id: 'L1', name: 'Title Slide', placeholders: [['Placeholder', 'CenterTitle', { left: 43, top: 86, width: 634, height: 158 }], ['Placeholder', 'Subtitle']] }];
  const { api } = fakePowerPoint({ inherited: true, layouts, attention: false });
  const frames = [], writes = [];
  const tracked = { ...overlay(), frame: target => frames.push(target && structuredClone(target)), write: target => writes.push(target && structuredClone(target)) };
  const { steps } = validatePptxOps([{ op: 'addSlide', ref: 'c', layout: 'L1', placeholders: { title: '封面', subtitle: '副标题' } }], blank);
  await runSteps(api, steps, { animate: true, wait: async () => {}, overlay: tracked, document: {}, typing: 'overlay' });
  const shown = frames.filter(Boolean);
  assert.deepEqual(shown.map(target => target.geometry), [{ left: 43, top: 86, width: 634, height: 158 }], 'the title takes its layout box');
  assert.equal(frames.filter(target => target === undefined).length >= 2, true, 'the subtitle (no size anywhere) shows no outline');
  assert.ok(writes.filter(Boolean).every(target => target.geometry.width > 0 && target.geometry.height > 0), 'typing is never drawn in a 0-size box');
  // A layout with two placeholders of one kind (two body columns): the kind does not say which box, so none is used.
  const columns = [{ id: 'L1', name: 'Two Content', placeholders: [['Placeholder', 'Body', { left: 36, top: 140, width: 312, height: 230 }], ['Placeholder', 'Body', { left: 372, top: 140, width: 312, height: 230 }]] }];
  const two = fakePowerPoint({ inherited: true, layouts: columns, attention: false });
  const twoFrames = [];
  await runSteps(two.api, validatePptxOps([{ op: 'addSlide', ref: 'c', layout: 'L1', placeholders: { body: '左栏' } }], blank).steps,
    { animate: true, wait: async () => {}, overlay: { ...overlay(), frame: target => twoFrames.push(target) }, document: {}, typing: 'overlay' });
  assert.deepEqual(twoFrames.filter(Boolean), [], 'no outline over a box that may be the other column');
});

test('animation outlines the work on the overlay by slide and geometry; the editor\'s focusAttention only on request', async () => {
  const { steps } = validatePptxOps([{ op: 'addTextBox', slideId: 's1', text: '你好', left: 1, top: 2, width: 3, height: 4 }], blank);
  const targets = [];
  const tracked = { ...overlay(), frame: (target, options) => targets.push([target && structuredClone(target), options?.tone]) };
  const plain = fakePowerPoint();
  await runSteps(plain.api, structuredClone(steps), { animate: true, wait: async () => {}, overlay: tracked, document: {}, typing: 'overlay' });
  // office-kit K8: the editor's outline leaves the shape once the slide repaints, so it is not used by default.
  assert.equal(plain.calls.filter(call => call[0] === 'attention').length, 0, 'no focusAttention by default');
  assert.deepEqual(targets[0], [{ slideId: 's1', geometry: { left: 1, top: 2, width: 3, height: 4 } }, 'edit'], 'the overlay is told which slide and where, not a screen position');
  assert.deepEqual(targets.at(-1), [undefined, undefined], 'and cleared at the end');
  const withAttention = fakePowerPoint();
  const shown = overlay();
  await runSteps(withAttention.api, structuredClone(steps), { animate: true, wait: async () => {}, overlay: shown, document: {}, typing: 'overlay', outline: 'editor' });
  const outlines = withAttention.calls.filter(call => call[0] === 'attention').map(call => call[1]);
  assert.deepEqual(outlines, [{ left: 1, top: 2, width: 3, height: 4 }, null], 'outlined where the box lands, then cleared');
  assert.ok(!shown.log.some(entry => entry[0] === 'frame' && entry[1] === 'on'), 'no overlay frame while the editor outlines');
  const without = fakePowerPoint({ attention: false });
  const fallback = overlay();
  await runSteps(without.api, structuredClone(steps), { animate: true, wait: async () => {}, overlay: fallback, document: {}, typing: 'overlay', outline: 'editor' });
  assert.ok(fallback.log.some(entry => entry[0] === 'frame' && entry[2] === 'edit'), 'the overlay frames instead');
  // A deletion keeps the overlay frame (its cursor says it deletes) even when the editor can outline.
  const deck = { slides: [{ id: 's1', shapes: [{ id: 'x', left: 0, top: 0, width: 5, height: 5 }] }] };
  const del = fakePowerPoint();
  del.slides[0].shapes.push({ id: 'x', type: 'TextBox', text: '' });
  const red = overlay();
  await runSteps(del.api, validatePptxOps([{ op: 'deleteShape', slideId: 's1', shapeId: 'x' }], deck).steps, { animate: true, wait: async () => {}, overlay: red, document: {}, outline: 'editor', geometryOf: () => deck.slides[0].shapes[0] });
  assert.ok(red.log.some(entry => entry[0] === 'frame' && entry[2] === 'delete'));
});

test('refs, new shapes, images, tables and charts are checked before anything runs', () => {
  const reject = (ops, pattern) => assert.throws(() => validatePptxOps(ops, blank), error => error.code === 'invalid-edit' && pattern.test(error.message), pattern.source);
  reject([{ op: 'addTextBox', slideRef: 'later', text: 'x', left: 0, top: 0, width: 1, height: 1 }], /slideRef "later" is not a slide added earlier/);
  reject([{ op: 'addSlide', ref: 'a' }, { op: 'addSlide', ref: 'a' }], /already used/);
  reject([{ op: 'addSlide', ref: '1st' }], /ref must start with a letter/);
  reject([{ op: 'addSlide', ref: 'a' }, { op: 'setText', slideRef: 'a', shapeRef: 'b', text: 'x' }], /shapeRef "b" is not a shape added earlier/);
  reject([{ op: 'addTextBox', slideId: 's1', slideRef: 'a', text: 'x', left: 0, top: 0, width: 1, height: 1 }], /not both/);
  reject([{ op: 'addTextBox', slideId: 's1', text: 'x', left: 0, top: 0, width: 1 }], /height/);
  reject([{ op: 'addSlide', layout: 'Fancy' }], /layout "Fancy" is not in this presentation/);
  reject([{ op: 'addSlide', placeholders: { title: 'x' } }], /need a layout/);
  reject([{ op: 'addSlide', layout: 'L1', placeholders: { footer: 'x' } }], /placeholders takes/);
  reject([{ op: 'addShape', slideId: 's1', shape: 'Heart', left: 0, top: 0, width: 1, height: 1 }], /shape must be one of/);
  reject([{ op: 'addShape', slideId: 's1', shape: 'Rectangle', left: 0, top: 0, width: 1, height: 1, bold: true }], /font settings need text/);
  reject([{ op: 'addImage', slideId: 's1', path: 'a.png', left: 0, top: 0, width: 1, height: 1 }], /image was not loaded/);
  reject([{ op: 'addImage', slideId: 's1', imageType: 'image/gif', imageBase64: 'R0lG', left: 0, top: 0, width: 1, height: 1 }], /image was not loaded/);
  const big = 'A'.repeat(Math.ceil((PPTX_LIMITS.maxImageBytes + 3) / 3) * 4);
  reject([{ op: 'addImage', slideId: 's1', imageType: 'image/png', imageBase64: big, left: 0, top: 0, width: 1, height: 1 }], /larger than 10 MB/);
  const nine = 'A'.repeat(Math.floor(9 * 1024 * 1024 / 3) * 4);
  reject([0, 1, 2].map(() => ({ op: 'addImage', slideId: 's1', imageType: 'image/png', imageBase64: nine, left: 0, top: 0, width: 1, height: 1 })), /add up to more than 20 MB/);
  reject([{ op: 'addTable', slideId: 's1', values: [['a', 'b'], ['c']], left: 0, top: 0, width: 1, height: 1 }], /row 2 must have 2 cells/);
  reject([{ op: 'addTable', slideId: 's1', values: Array.from({ length: PPTX_LIMITS.maxTableRows + 1 }, () => ['x']), left: 0, top: 0, width: 1, height: 1 }], /rows/);
  reject([{ op: 'addChart', slideId: 's1', chartType: 'sankey', categories: ['a'], series: [{ name: 's', values: [1] }], left: 0, top: 0, width: 1, height: 1 }], /chartType/);
  reject([{ op: 'addChart', slideId: 's1', categories: ['a'], series: Array.from({ length: 7 }, (_, i) => ({ name: `s${i}`, values: [1] })), left: 0, top: 0, width: 1, height: 1 }], /1–6 series/);
  reject([{ op: 'addChart', slideId: 's1', categories: ['a', 'b'], series: [{ name: 's', values: [1] }], left: 0, top: 0, width: 1, height: 1 }], /needs 2 numbers/);
  reject([{ op: 'addChart', slideId: 's1', chartType: 'pie', categories: ['a'], series: [{ name: 's', values: [1] }, { name: 't', values: [2] }], left: 0, top: 0, width: 1, height: 1 }], /exactly one series/);
  reject([{ op: 'setBackground', slideId: 's1', color: 'white' }], /#RRGGBB/);
  // A deleted ref is gone for the rest of the batch.
  reject([{ op: 'addSlide', ref: 'a' }, { op: 'addTextBox', slideRef: 'a', ref: 't', text: 'x', left: 0, top: 0, width: 1, height: 1 }, { op: 'deleteShape', slideRef: 'a', shapeRef: 't' }, { op: 'setText', slideRef: 'a', shapeRef: 't', text: 'y' }], /Operation 4: shapeRef "t"/);
});

test('the confirmation card and the model\'s read show new objects, refs and layouts', async () => {
  const { steps } = validatePptxOps([
    { op: 'addSlide', ref: 'p2', layout: 'Title Slide', placeholders: { title: '封面' } },
    { op: 'addImage', slideRef: 'p2', path: 'logo.png', imageType: 'image/png', imageBase64: 'iVBORw0KGgo=', left: 0, top: 0, width: 10, height: 10 },
    { op: 'addTable', slideId: 's1', values: [['a', 'b']], left: 0, top: 0, width: 1, height: 1 },
  ], blank);
  assert.deepEqual(describePptxSteps(steps, blank), [
    { op: 'addSlide', where: '#2 @p2', after: 'Title Slide: 封面' },
    { op: 'addImage', where: '@p2', after: 'logo.png' },
    { op: 'addTable', where: '#1', after: '1×2: a | b' },
  ]);
  const { api } = fakePowerPoint();
  const deck = await inspectDeck(api);
  assert.deepEqual(deck.layouts, [{ id: 'L1', name: 'Title Slide', placeholders: ['title', 'subtitle'] }, { id: 'L2', name: 'Blank', placeholders: [] }]);
  assert.deepEqual(deck.slideSize, { width: 960, height: 540 });
  const summary = summarizeDeck(deck);
  assert.deepEqual(summary.layouts, deck.layouts);
  assert.deepEqual(summary.slideSize, { width: 960, height: 540 });
});

test('tables and charts stay held back from the model until Kit can save them (K6)', () => {
  assert.deepEqual([...PPTX_HELD_OPS], ['addTable', 'addChart']);
  assert.ok(PPTX_HELD_OPS.every(op => !PPTX_OPS.includes(op) && PPTX_ALL_OPS.includes(op)));
  for (const op of [{ op: 'addTable', slideId: 's1', values: [['a']], left: 0, top: 0, width: 1, height: 1 },
    { op: 'addChart', slideId: 's1', categories: ['a'], series: [{ name: 's', values: [1] }], left: 0, top: 0, width: 1, height: 1 }]) {
    assert.throws(() => check([op], blank), error => error.code === 'invalid-edit' && /not available yet/.test(error.message));
  }
  assert.throws(() => check([{ op: 'eval' }], blank), /unknown operation/);
});


// Several slides at once (spec §11): a batch of new slides plays together.
const deckWith = layouts => ({ slides: [{ id: 's1', shapes: [] }], layouts });
const TITLED = [{ id: 'L1', name: 'Title and Content', placeholders: ['title', 'body'] }, { id: 'L2', name: 'Blank', placeholders: [] }];
const titledLayouts = [{ id: 'L1', name: 'Title and Content', placeholders: [['Placeholder', 'Title'], ['Placeholder', 'Body']] }, { id: 'L2', name: 'Blank', placeholders: [] }];
const yielding = { animate: true, typing: 'overlay', wait: () => new Promise(resolve => setImmediate(resolve)) };
const keyed = () => {
  const log = [];
  return { log, frame: (target, options) => log.push(['frame', options?.key, target?.slideId]), write: (target, value, key) => log.push(['write', key, target?.slideId]),
    say: text => log.push(['say', text]), mark: ids => log.push(['mark', [...ids]]), busy: () => {} };
};
const threePages = [
  { op: 'addSlide', ref: 'a', layout: 'L1', placeholders: { title: '第一页', body: '要点一' } },
  { op: 'addSlide', ref: 'b', layout: 'L1', placeholders: { title: '第二页', body: '要点二' } },
  { op: 'addSlide', ref: 'c', layout: 'L2' },
  { op: 'addTextBox', slideRef: 'c', ref: 'note', text: '第三页的字', left: 40, top: 40, width: 400, height: 60 },
  { op: 'addTextBox', slideRef: 'a', text: '第一页补充', left: 40, top: 300, width: 400, height: 60 },
];

test('one edit adds at most four slides; the Host and the editor check the same count', () => {
  const add = n => Array.from({ length: n }, () => ({ op: 'addSlide' }));
  assert.equal(PPTX_LIMITS.maxNewSlides, 4);
  assert.equal(countNewSlides(add(5)), 5);
  assert.equal(countNewSlides('nope'), 0);
  assert.doesNotThrow(() => validatePptxOps(add(4), blank));
  assert.throws(() => validatePptxOps(add(5), blank), error => error.code === 'invalid-edit' && error.message === NEW_SLIDES_MESSAGE && /several presentation_edit calls/.test(error.message));
  // Changing existing slides is not limited by it.
  const many = { slides: Array.from({ length: 8 }, (_, index) => ({ id: `s${index}`, shapes: [{ id: 't', type: 'TextBox', text: 'x' }] })) };
  assert.doesNotThrow(() => validatePptxOps(many.slides.map(slide => ({ op: 'setText', slideId: slide.id, shapeId: 't', text: 'y' })), many));
});

test('a placeholder the layout does not have fails before anything runs; empty text for it is fine', () => {
  assert.throws(() => validatePptxOps([{ op: 'addSlide', layout: 'L2', placeholders: { title: '标题' } }], deckWith(TITLED)), error => error.code === 'invalid-edit' && /has no title placeholder \(it has: none\)/.test(error.message));
  assert.doesNotThrow(() => validatePptxOps([{ op: 'addSlide', layout: 'L2', placeholders: { title: '' } }], deckWith(TITLED)));
  assert.doesNotThrow(() => validatePptxOps([{ op: 'addSlide', layout: 'L1', placeholders: { title: '标题', body: '正文' } }], deckWith(TITLED)));
  // Without placeholder lists (an older editor) the check is left to the run.
  assert.doesNotThrow(() => validatePptxOps([{ op: 'addSlide', layout: 'L2', placeholders: { title: '标题' } }], blank));
});

test('which batches play together: new slides, 2–4 slides, overlay typing and outline, no slide deleted or moved', () => {
  const steps = ops => validatePptxOps(ops, deckWith(TITLED)).steps;
  const options = { animate: true, typing: 'overlay', outline: 'overlay' };
  assert.equal(playsTogether(steps(threePages), options), true);
  assert.equal(playsTogether(steps(threePages.slice(0, 1)), options), false, 'one slide');
  assert.equal(playsTogether(steps([{ op: 'addSlide' }, { op: 'setBackground', slideId: 's1', color: '#FFFFFF' }]), options), true, 'a new slide and an existing one');
  assert.equal(playsTogether(steps(threePages), { ...options, animate: false }), false, 'no animation');
  assert.equal(playsTogether(steps(threePages), { ...options, typing: 'document' }), false, 'typing into the document');
  assert.equal(playsTogether(steps(threePages), { ...options, outline: 'editor' }), false, 'the editor outline reaches the shown slide only');
  assert.equal(playsTogether(steps([...threePages, { op: 'moveSlide', slideId: 's1', index: 0 }]), options), false, 'order matters');
  const five = { slides: Array.from({ length: 4 }, (_, index) => ({ id: `e${index}`, shapes: [] })), layouts: TITLED };
  const wide = validatePptxOps([{ op: 'addSlide' }, ...five.slides.map(slide => ({ op: 'setBackground', slideId: slide.id, color: '#FFFFFF' }))], five).steps;
  assert.equal(playsTogether(wide, options), false, 'more than four slides');
  const existingOnly = validatePptxOps(five.slides.slice(0, 2).map(slide => ({ op: 'setBackground', slideId: slide.id, color: '#FFFFFF' })), five).steps;
  assert.equal(playsTogether(existingOnly, options), false, 'changing existing slides only still goes slide by slide');
});

test('a batch of new slides builds them in order, selects the first once and fills them at the same time', async () => {
  const { steps } = validatePptxOps([{ op: 'addSlide', ref: 'z', layout: 'L2', index: 0 }, ...threePages], deckWith(TITLED));
  assert.equal(steps.length, 6);
  const { api, calls, state } = fakePowerPoint({ layouts: titledLayouts });
  const shown = keyed();
  const refs = await runSteps(api, steps, { ...yielding, overlay: shown, document: {}, labels: LABELS.zh });
  assert.equal(state.overlap, 0, 'editor calls never overlap');
  assert.ok(refs.a && refs.b && refs.c && refs.z && refs.note, 'every ref resolved');
  const firstText = calls.findIndex(call => call[0] === 'text');
  assert.ok(calls.findLastIndex(call => call[0] === 'addSlide') < firstText, 'all slides exist before any is filled');
  assert.deepEqual(calls.find(call => call[0] === 'moveTo'), ['moveTo', refs.z, 0], 'a slide asked for at the front is moved there before the others are added');
  assert.deepEqual(calls.filter(call => call[0] === 'select'), [['select', refs.z]], 'one selection: the first slide in the deck');
  // The tracks interleave: slide c gets its text box while slide a is still being filled.
  const writes = calls.filter(call => call[0] === 'text').map(call => call[2]);
  assert.deepEqual(writes.toSorted(), ['第一页', '第二页', '要点一', '要点二'].toSorted());
  const boxOn = slideId => calls.findIndex(call => call[0] === 'addTextBox' && call[1] === slideId);
  assert.ok(boxOn(refs.c) < calls.findIndex(call => call[0] === 'text' && call[2] === '要点一'), 'slides fill together');
  assert.ok(boxOn(refs.c) < boxOn(refs.a));
  // One outline track per slide, keyed by its id; the thumbnails of all of them are ringed, then cleared.
  const keys = new Set(shown.log.filter(entry => entry[0] === 'frame' && entry[2]).map(entry => entry[1]));
  assert.deepEqual([...keys].toSorted(), [refs.a, refs.b, refs.c].toSorted());
  assert.ok(shown.log.filter(entry => entry[0] === 'frame' && entry[2]).every(entry => entry[1] === entry[2]), 'each track outlines its own slide');
  const marks = shown.log.filter(entry => entry[0] === 'mark').map(entry => entry[1]);
  assert.deepEqual(marks[0], [refs.z, refs.a, refs.b, refs.c]);
  assert.deepEqual(marks.at(-1), []);
  assert.ok(shown.log.some(entry => entry[0] === 'say' && entry[1] === 'AI 正在生成第 1、3、4、5 页'), 'the status names the slides');
  assert.equal(LABELS.zh.slides([4, 2, 3]), 'AI 正在生成第 2–4 页');
});

test('a slide that fails stops the others; the error comes after every track has stopped', async () => {
  const { steps } = validatePptxOps(threePages, deckWith(TITLED));
  const { api, calls } = fakePowerPoint({ layouts: titledLayouts, failText: '要点二' });
  await assert.rejects(runSteps(api, steps, { ...yielding, overlay: keyed(), document: {} }), error => error.code === 'editor-failed');
  const settled = calls.length;
  await new Promise(resolve => setTimeout(resolve, 20));
  assert.equal(calls.length, settled, 'nothing is written after the error');
  assert.equal(calls.filter(call => call[0] === 'addTextBox').length < 2, true, 'the other tracks stopped early');
  // Cancelled from outside: the same, with the abort as the error.
  const controller = new AbortController();
  const cancelled = fakePowerPoint({ layouts: titledLayouts });
  let waits = 0;
  const running = runSteps(cancelled.api, validatePptxOps(threePages, deckWith(TITLED)).steps, { ...yielding, overlay: keyed(), document: {}, signal: controller.signal,
    wait: () => { if (++waits === 6) controller.abort(); return new Promise(resolve => setImmediate(resolve)); } });
  await assert.rejects(running, error => error.name === 'AbortError');
  const stopped = cancelled.calls.length;
  await new Promise(resolve => setTimeout(resolve, 20));
  assert.equal(cancelled.calls.length, stopped);
});

test('step by step as before when a batch does not play together; its current slide is ringed', async () => {
  const { steps } = validatePptxOps([threePages[0], { op: 'addTextBox', slideRef: 'a', text: '补充', left: 1, top: 2, width: 3, height: 4 }], deckWith(TITLED));
  const { api, calls } = fakePowerPoint({ layouts: titledLayouts });
  const shown = keyed();
  const refs = await runSteps(api, steps, { ...yielding, overlay: shown, document: {} });
  assert.ok(calls.filter(call => call[0] === 'select').every(call => call[1] === refs.a));
  assert.ok(shown.log.every(entry => entry[0] !== 'frame' || entry[1] === undefined), 'the main track only');
  assert.deepEqual(shown.log.filter(entry => entry[0] === 'mark').map(entry => entry[1]), [[refs.a], [refs.a], []]);
  // A new slide on its own: added, its unused placeholders removed, then shown and filled (no "click to add" flash).
  const single = fakePowerPoint({ layouts: titledLayouts });
  const one = await runSteps(single.api, validatePptxOps([{ op: 'addSlide', ref: 'p', layout: 'L1', placeholders: { title: '只有标题', body: '' } }], deckWith(TITLED)).steps, { ...yielding, overlay: keyed(), document: {} });
  assert.deepEqual(single.calls.filter(call => ['addSlide', 'delete', 'select', 'text'].includes(call[0])).map(call => call[0]), ['addSlide', 'delete', 'select', 'text']);
  assert.deepEqual(single.calls.find(call => call[0] === 'text'), ['text', single.slides.find(slide => slide.id === one.p).shapes[0].id, '只有标题']);
});

test('Kit agent (v1.1): new slides are filled page by page, each selected in turn, and a pick inside the batch stays; a failed batch is undone whole', async () => {
  const { api, calls, select } = fakePowerPoint({ layouts: titledLayouts });
  const node = () => ({ style: {}, dataset: {}, append() {}, remove() {}, setAttribute() {} });
  const page = { body: { append() {} }, createElement: node, createElementNS: node, addEventListener() {}, removeEventListener() {}, querySelectorAll: () => [] };
  const editor = { api: { presentation: { getJsApi: () => api, editorInternals: () => ({ getDocument: () => ({ store: { encodeBinary: () => new Uint8Array([0]), canUndo: false } }), undo() {} }) } } };
  // The user clicks the second new slide's thumbnail while the batch runs.
  const picking = () => { const ids = calls.filter(call => call[0] === 'addSlide').map(call => call[1]); if (ids.length === 2 && calls.some(call => call[0] === 'select')) select(ids[1]); return new Promise(resolve => setImmediate(resolve)); };
  const agent = createKitPresentationAgent({ editor, document: page, wait: picking });
  const ops = [threePages[0], threePages[1]];
  const result = await agent.apply({ ops, editId: 'e1' });
  const added = calls.filter(call => call[0] === 'addSlide').map(call => call[1]);
  assert.deepEqual(result.slideIds.toSorted(), added.toSorted());
  assert.deepEqual(calls.filter(call => call[0] === 'select'), [['select', added[0]], ['select', added[1]]], 'page by page (spec §7.6 4), not pulled back at the end');
  // The user went to a slide outside the batch: the first of the batch is shown at the end.
  const quiet = fakePowerPoint({ layouts: titledLayouts });
  const editor2 = { api: { presentation: { getJsApi: () => quiet.api, editorInternals: () => ({ getDocument: () => ({ store: { encodeBinary: () => new Uint8Array([0]), canUndo: false } }), undo() {} }) } } };
  const away = () => { if (quiet.calls.some(call => call[0] === 'select')) quiet.select('s1'); return new Promise(resolve => setImmediate(resolve)); };
  const agent2 = createKitPresentationAgent({ editor: editor2, document: page, wait: away });
  await agent2.apply({ ops, editId: 'e2' });
  const added2 = quiet.calls.filter(call => call[0] === 'addSlide').map(call => call[1]);
  assert.deepEqual(quiet.calls.filter(call => call[0] === 'select'), [['select', added2[0]], ['select', added2[1]], ['select', added2[0]]]);
  // A failing slide: the editor history is walked back to the bytes before the batch (here: one byte per editor call).
  const failing = fakePowerPoint({ layouts: titledLayouts, failText: '要点二' });
  let bytes = 0;
  const editor3 = { api: { presentation: { getJsApi: () => failing.api, editorInternals: () => ({ getDocument: () => ({ store: { encodeBinary: () => new Uint8Array([bytes]), get canUndo() { return bytes > 0; } } }), undo() { bytes--; } }) } } };
  const counting = { ...failing.api, PowerPoint: { run: async body => { bytes++; return failing.api.PowerPoint.run(body); } } };
  editor3.api.presentation.getJsApi = () => counting;
  const agent3 = createKitPresentationAgent({ editor: editor3, document: page, wait: () => new Promise(resolve => setImmediate(resolve)) });
  await assert.rejects(agent3.apply({ ops, editId: 'e3' }), error => error.code === 'editor-failed');
  assert.equal(bytes, 0, 'back to the bytes before the batch');
  assert.ok(failing.calls.some(call => call[0] === 'addSlide'), 'it had written before it failed');
});

test('the agent cursor says what it does; a chart is outlined on its edge, a background change and a deleted slide outline the whole slide', async () => {
  const deck = { slides: [
    { id: 's1', shapes: [{ id: 'x', type: 'TextBox', text: '旧', left: 10, top: 10, width: 100, height: 20 }, { id: 'c', type: 'Chart', left: 10, top: 60, width: 200, height: 120 }] },
    { id: 's2', shapes: [] },
  ] };
  const pp = fakePowerPoint();
  pp.slides[0].shapes.push({ id: 'x', type: 'TextBox', text: '旧' }, { id: 'c', type: 'Chart', text: '' });
  pp.slides.push({ id: 's2', shapes: [] });
  const framed = [];
  const tracked = { ...overlay(), frame: (target, options) => framed.push([target && structuredClone(target), options?.text, options?.kind, pp.calls.length]) };
  const { steps } = validatePptxOps([
    { op: 'setText', slideId: 's1', shapeId: 'x', text: '新' },
    { op: 'setFill', slideId: 's1', shapeId: 'c', color: '#FF0000' },
    { op: 'addTextBox', slideId: 's1', text: '加', left: 1, top: 2, width: 3, height: 4 },
    { op: 'deleteShape', slideId: 's1', shapeId: 'x' },
    { op: 'setBackground', slideId: 's1', color: '#FFFFFF' },
    { op: 'deleteSlide', slideId: 's2' },
  ], deck);
  const shapes = new Map(deck.slides.flatMap(slide => slide.shapes.map(shape => [`${slide.id}/${shape.id}`, shape])));
  await runSteps(pp.api, steps, { animate: true, wait: async () => {}, overlay: tracked, document: {}, typing: 'overlay', labels: LABELS.zh, geometryOf: (slideId, shapeId) => shapes.get(`${slideId}/${shapeId}`) });
  const shown = framed.filter(([target]) => target).map(([target, text, kind]) => [target.whole ? `${target.slideId} whole` : `${target.slideId} ${target.geometry.left}`, text, kind]);
  assert.deepEqual(shown, [
    ['s1 10', 'Agent · 修改中', 'text'],
    ['s1 10', 'Agent · 修改中', 'edge'],
    ['s1 1', 'Agent · 新增中', 'text'],
    ['s1 10', 'Agent · 删除中', 'text'],
    ['s1 whole', 'Agent · 修改中', 'edge'],
    ['s2 whole', 'Agent · 删除中', 'edge'],
  ]);
  // The slide's outline goes as soon as the slide does, not left over the slide the editor shows next.
  const deleted = pp.calls.findIndex(call => call[0] === 'deleteSlide');
  const whole = framed.findIndex(([target]) => target?.whole && target.slideId === 's2');
  assert.ok(framed[whole][3] <= deleted, 'outlined before the deletion');
  assert.deepEqual(framed[whole + 1].slice(0, 1), [undefined], 'cleared right after it');
  assert.ok(framed[whole + 1][3] > deleted);
  // A new chart (held for now, office-kit K6) is outlined on its edge as well.
  const chart = fakePowerPoint();
  const charts = [];
  const added = validatePptxOps([{ op: 'addChart', slideId: 's1', chartType: 'column', categories: ['Q1'], series: [{ name: '收入', values: [1] }], left: 0, top: 0, width: 100, height: 80 }], blank);
  await runSteps(chart.api, added.steps, { animate: false, overlay: { ...overlay(), frame: (target, options) => charts.push(options?.kind) }, document: {}, labels: LABELS.zh });
  assert.equal(charts[0], 'edge');
});

test('with the editor\'s outline (K8) a whole slide is still outlined on the overlay, and the shape\'s editor outline goes first', async () => {
  const deck = { slides: [{ id: 's1', shapes: [{ id: 'x', type: 'TextBox', text: '旧', left: 10, top: 10, width: 100, height: 20 }] }, { id: 's2', shapes: [] }] };
  const pp = fakePowerPoint();
  pp.slides[0].shapes.push({ id: 'x', type: 'TextBox', text: '旧' });
  pp.slides.push({ id: 's2', shapes: [] });
  const framed = [];
  const tracked = { ...overlay(), frame: (target, options) => framed.push([target && structuredClone(target), options?.text, pp.calls.length]) };
  const { steps } = validatePptxOps([
    { op: 'setText', slideId: 's1', shapeId: 'x', text: '新' },
    { op: 'setBackground', slideId: 's1', color: '#FFFFFF' },
    { op: 'deleteSlide', slideId: 's2' },
  ], deck);
  await runSteps(pp.api, steps, { animate: true, wait: async () => {}, overlay: tracked, document: {}, labels: LABELS.zh, outline: 'editor', geometryOf: (slideId, shapeId) => deck.slides.find(slide => slide.id === slideId).shapes.find(shape => shape.id === shapeId) });
  const outlines = pp.calls.map((call, index) => [call, index]).filter(([call]) => call[0] === 'attention');
  assert.deepEqual(outlines.map(([call]) => call[1] && 'shape'), ['shape', null], 'the editor outlined the text box only, then cleared it once');
  const wholes = framed.filter(([target]) => target?.whole);
  assert.deepEqual(wholes.map(([target, text]) => [target.slideId, text]), [['s1', 'Agent · 修改中'], ['s2', 'Agent · 删除中']]);
  assert.ok(wholes[0][2] > outlines[1][1], 'the editor outline was cleared before the slide was outlined');
});

test('v1.1: each target is selected before it is written, text grows in a few real writes, statuses follow the operation', async () => {
  const { api, calls } = fakePowerPoint();
  const deck = await inspectDeck(api);
  const long = '这是一段较长的正文内容，用来检查分段写入是否真的发生在形状里，'.repeat(2);
  const { steps } = validatePptxOps([{ op: 'addTextBox', slideId: 's1', text: '大标题', size: 40, left: 10, top: 10, width: 400, height: 60 }, { op: 'addTextBox', slideId: 's1', text: long, left: 10, top: 100, width: 400, height: 200 }], deck);
  const events = [];
  const overlay = {
    frame: (target, options) => { events.push(['frame', options?.text, options?.select === true, options?.key]); return Promise.resolve(); },
    write() {}, say() {}, mark() {}, busy() {},
  };
  const original = calls.push.bind(calls);
  calls.push = (...items) => { for (const item of items) if (item[0] === 'text' || item[0] === 'addTextBox') events.push([item[0], item[2]]); return original(...items); };
  await runSteps(api, steps, { overlay, document: {}, animate: true, typing: 'chunks', locale: 'zh', wait: async () => {} });
  const firstText = events.findIndex(event => event[0] === 'addTextBox');
  assert.deepEqual(events[0], ['frame', 'Agent · 写入标题', true, 'k0'], 'the title is selected (with its status) before anything is written');
  assert.ok(firstText > 0);
  const secondFrame = events.findIndex((event, index) => index > 0 && event[0] === 'frame' && event[1]);
  assert.deepEqual(events[secondFrame].slice(0, 4), ['frame', 'Agent · 完善结论', true, 'k1'], 'the lowest body text is the conclusion, on the next track');
  const bodyWrites = events.slice(secondFrame).filter(event => event[0] === 'text').map(event => event[1]);
  assert.ok(bodyWrites.length >= 3 && bodyWrites.length <= 5, `a few real writes, got ${bodyWrites.length}`);
  assert.equal(bodyWrites.at(-1), long);
  assert.ok(bodyWrites.every((text, index) => index === 0 || text.length >= bodyWrites[index - 1].length), 'the text only grows');
});

test('v1.1: a new slide\'s placeholders are labelled by their kind, not as page creation', async () => {
  const { api } = fakePowerPoint({ layouts: titledLayouts });
  const deck = await inspectDeck(api);
  const { steps } = validatePptxOps([threePages[0]], deck);
  const labels = [];
  const overlay = { frame: (target, options) => { if (options?.text) labels.push(options.text); return Promise.resolve(); }, write() {}, say() {}, mark() {}, busy() {} };
  await runSteps(api, steps, { overlay, document: {}, animate: true, typing: 'chunks', locale: 'zh', wait: async () => {} });
  assert.ok(labels.includes('Agent · 写入标题'), labels.join(' / '));
  assert.ok(!labels.slice(1).includes('Agent · 创建页面'), 'the fills are not labelled as page creation');
});

// Slide sorter (spec 261105-sorter-view-live B2–B5): PowerPoint.Public.setViewMode, faked; the view is read through `viewOf`.
const twoNew = [
  { op: 'addSlide', ref: 'a', layout: 'L1', placeholders: { title: '第一页', body: '要点一' } },
  { op: 'addSlide', ref: 'b', layout: 'L1', placeholders: { title: '第二页', body: '要点二' } },
  { op: 'addTextBox', slideRef: 'b', text: '第二页补充', left: 40, top: 300, width: 400, height: 60 },
];
/** A fake editor with (or without: `setViewMode: false`) the view switch; `views` lists each call, `current` is the view shown. */
function sorterEditor({ setViewMode = true, behave = () => {} } = {}) {
  const fake = fakePowerPoint({ layouts: titledLayouts });
  const views = [];
  const state = { current: NORMAL };
  if (setViewMode) fake.api.PowerPoint.Public = { print() {}, setViewMode(mode) { views.push(mode); fake.calls.push(['view', mode]); state.current = mode === SORTER ? SORTER : NORMAL; return behave(mode, state); } };
  return { ...fake, views, state, viewOf: () => state.current };
}
const recorder = () => {
  const log = [];
  return { log, frame: (target, options) => { log.push(['frame', options?.key, target?.slideId, target?.whole === true, options?.text, options?.select === true]); return Promise.resolve(); },
    write() {}, say: text => log.push(['say', text]), mark() {}, busy() {}, view: mode => log.push(['view', mode]) };
};
const v11 = { animate: true, typing: 'chunks', locale: 'zh', wait: async () => {} };
const sorterRun = async (editor, ops = twoNew, options = {}) => {
  const { steps } = validatePptxOps(ops, await inspectDeck(editor.api));
  const overlay = recorder(), report = {};
  await runSteps(editor.api, steps, { overlay, document: {}, viewOf: editor.viewOf, report, ...v11, ...options });
  return { overlay, report, steps };
};
const withoutViews = calls => calls.filter(call => call[0] !== 'view');

test('sorter: without setViewMode nothing changes — the same editor calls and overlay calls as 0.0.40', async () => {
  const before = sorterEditor({ setViewMode: false });
  const plain = await sorterRun(before);
  const other = sorterEditor({ setViewMode: false });
  other.api.PowerPoint.Public = { print() {} };
  const withPublic = await sorterRun(other);
  const throwing = sorterEditor({ setViewMode: false });
  Object.defineProperty(throwing.api.PowerPoint, 'Public', { get() { throw new Error('InvalidContext'); } });
  const unreadable = await sorterRun(throwing);
  assert.equal(plain.report.view, 'unsupported');
  assert.deepEqual(other.calls, before.calls);
  assert.deepEqual(throwing.calls, before.calls);
  assert.deepEqual(withPublic.overlay.log, plain.overlay.log);
  assert.deepEqual(unreadable.overlay.log, plain.overlay.log);
  assert.ok(plain.report.together);
});

test('sorter: two new slides are created blank in the slide sorter, one card at a time, then filled page by page in the normal view', async () => {
  const editor = sorterEditor();
  const { overlay, report, steps } = await sorterRun(editor);
  assert.deepEqual(editor.views, [SORTER, NORMAL], 'the exact modes, once each');
  const at = (kind, value) => editor.calls.findIndex(call => call[0] === kind && (value === undefined || call[1] === value));
  const lastIndex = kind => editor.calls.map(call => call[0]).lastIndexOf(kind);
  assert.ok(at('view', SORTER) < at('addSlide'), 'into the sorter before the first slide is created');
  assert.ok(lastIndex('addSlide') < at('view', NORMAL) && lastIndex('delete') < at('view', NORMAL), 'every blank slide is ready (unused placeholders gone) before going back');
  assert.ok(at('view', NORMAL) < at('text'), 'nothing is written before the normal view is back');
  const [a, b] = steps.filter(step => step.op === 'addSlide').map(step => step.createdId);
  const cards = overlay.log.filter(entry => entry[0] === 'frame' && entry[1] === 'sorter' && entry[2]);
  assert.deepEqual(cards, [['frame', 'sorter', a, true, 'Agent · 创建页面', true], ['frame', 'sorter', b, true, 'Agent · 创建页面', true]], 'each card outlined with the agent cursor, selected first');
  const viewsOnOverlay = overlay.log.filter(entry => entry[0] === 'view').map(entry => entry[1]);
  assert.deepEqual(viewsOnOverlay.slice(0, 1), [SORTER]);
  assert.equal(viewsOnOverlay.at(-1), NORMAL);
  const says = overlay.log.filter(entry => entry[0] === 'say').map(entry => entry[1]);
  assert.ok(says.indexOf(LABELS.zh.sorter) >= 0 && says.indexOf(LABELS.zh.toNormal(2)) > says.indexOf(LABELS.zh.sorter), says.join(' / '));
  assert.equal(report.view, 'sorter');
  assert.equal(report.together, true, 'report.together unchanged');
  assert.ok(editor.calls.findIndex(call => call[0] === 'select' && call[1] === a) > at('view', NORMAL), 'the first new slide is shown in the normal view');
  // Same document in the end as without the switch.
  const plain = sorterEditor({ setViewMode: false });
  await sorterRun(plain);
  assert.deepEqual(withoutViews(editor.calls), plain.calls);
});

test('sorter: not used for one new slide (even with an existing one), without motion, outside v1.1, from the sorter, or when the view is too narrow', async () => {
  const mixed = sorterEditor();
  const deck = await inspectDeck(mixed.api);
  const ops = [{ op: 'addSlide', ref: 'a', layout: 'L1', placeholders: { title: '新页' } }, { op: 'addTextBox', slideId: 's1', text: '旧页补充', left: 10, top: 10, width: 100, height: 40 }];
  assert.ok(playsTogether(validatePptxOps(ops, deck).steps, { animate: true, typing: 'chunks' }), 'the batch plays together');
  assert.equal((await sorterRun(mixed, ops)).report.view, 'skipped');
  assert.deepEqual(mixed.views, []);
  const still = sorterEditor();
  await sorterRun(still, twoNew, { animate: false, typing: 'overlay' });
  assert.deepEqual(still.views, [], 'reduced motion');
  const old = sorterEditor();
  await sorterRun(old, twoNew, { locale: undefined, typing: 'overlay' });
  assert.deepEqual(old.views, [], 'the flow before v1.1');
  const sorting = sorterEditor();
  sorting.state.current = SORTER;
  assert.equal((await sorterRun(sorting)).report.view, 'skipped');
  assert.deepEqual(sorting.views, [], 'the user is in the sorter already: the view stays theirs');
  // A main view 400 px wide: two readable cards do not fit.
  const area = { clientWidth: 400, getBoundingClientRect: () => ({ left: 0, top: 0, right: 400, bottom: 300, width: 400, height: 300 }) };
  const owner = { defaultView: { getComputedStyle: node => ({ overflow: node === area ? 'auto' : 'visible', overflowX: 'visible', overflowY: 'visible' }) } };
  const surface = { parentElement: area, ownerDocument: owner, viewBox: { baseVal: { width: 9144000 } }, getAttribute: name => name === 'data-mop-svg-definition-namespace' ? 'mop-main-x' : 'x', getBoundingClientRect: () => ({ left: 0, top: 0, width: 380, height: 214 }) };
  const narrow = sorterEditor();
  await sorterRun(narrow, twoNew, { document: { querySelectorAll: () => [surface] } });
  assert.deepEqual(narrow.views, [], 'too narrow');
  area.clientWidth = 900;
  const wide = sorterEditor();
  await sorterRun(wide, twoNew, { document: { querySelectorAll: () => [surface] } });
  assert.deepEqual(wide.views, [SORTER, NORMAL], 'wide enough');
});

test('sorter: a switch that throws, rejects or does not change the view builds in the normal view as before, and still ends with normal', async () => {
  const plain = sorterEditor({ setViewMode: false });
  await sorterRun(plain);
  for (const [name, behave] of [
    ['throws', mode => { if (mode === SORTER) throw new Error('InvalidContext'); }],
    ['rejects', mode => mode === SORTER ? Promise.reject(new Error('InvalidContext')) : undefined],
    ['no change', (mode, state) => { state.current = NORMAL; }],
  ]) {
    const editor = sorterEditor({ behave });
    const { report, overlay } = await sorterRun(editor);
    assert.equal(report.view, 'failed', name);
    assert.deepEqual(editor.views, [SORTER, NORMAL], `${name}: asked for the sorter, so back to normal`);
    assert.deepEqual(withoutViews(editor.calls), plain.calls, `${name}: the same document as without the switch`);
    assert.ok(!overlay.log.some(entry => entry[0] === 'frame' && entry[1] === 'sorter' && entry[2]), `${name}: no card frames`);
  }
});

test('sorter: stopped or failed while creating blank slides, it goes back to normal after the editor call in flight, and the error stays', async () => {
  const editor = sorterEditor();
  const controller = new AbortController();
  const original = editor.api.PowerPoint.run;
  editor.api.PowerPoint.run = async body => { try { return await original(body); } finally { editor.calls.push(['run-end']); } };
  const push = editor.calls.push.bind(editor.calls);
  let added = 0;
  editor.calls.push = (...items) => { for (const item of items) if (item[0] === 'addSlide' && ++added === 2) controller.abort(); return push(...items); };
  await assert.rejects(sorterRun(editor, twoNew, { signal: controller.signal }), error => error.name === 'AbortError');
  assert.deepEqual(editor.views, [SORTER, NORMAL]);
  const normalAt = editor.calls.findIndex(call => call[0] === 'view' && call[1] === NORMAL);
  assert.ok(editor.calls.map(call => call[0]).lastIndexOf('run-end') < normalAt, 'normal only after the call in flight finished');

  // The switch back fails too: the error that ended the batch is the one thrown.
  const failing = sorterEditor({ behave: mode => { if (mode === NORMAL) throw new Error('switch failed'); } });
  const boom = recorder();
  boom.frame = (target, options) => options?.key === 'sorter' && target ? Promise.reject(new Error('boom')) : Promise.resolve();
  const { steps } = validatePptxOps(twoNew, await inspectDeck(failing.api));
  const report = {};
  await assert.rejects(runSteps(failing.api, steps, { overlay: boom, document: {}, viewOf: failing.viewOf, report, ...v11 }), /boom/);
  assert.deepEqual(failing.views, [SORTER, NORMAL, NORMAL], 'tried twice');
  assert.equal(report.view, 'stuck');
});

test('sorter: left by someone else meanwhile, the rest is built without card frames; stuck in the sorter, normal is asked once more', async () => {
  const editor = sorterEditor();
  const push = editor.calls.push.bind(editor.calls);
  editor.calls.push = (...items) => { for (const item of items) if (item[0] === 'addSlide') editor.state.current = NORMAL; return push(...items); };
  const left = await sorterRun(editor);
  assert.equal(left.report.view, 'left');
  assert.ok(!left.overlay.log.some(entry => entry[0] === 'frame' && entry[1] === 'sorter' && entry[2]), 'no card frame once the sorter is gone');
  assert.deepEqual(editor.views, [SORTER, NORMAL], 'it still ends with normal');

  const stuck = sorterEditor({ behave: (mode, state) => { state.current = SORTER; } });
  const result = await sorterRun(stuck);
  assert.deepEqual(stuck.views, [SORTER, NORMAL, NORMAL]);
  assert.equal(result.report.view, 'stuck');
  assert.ok(stuck.calls.some(call => call[0] === 'text'), 'the slides are still filled');
});

test('sorter: an editor call failing while blank slides are created still ends in the normal view with that error; a stopped batch never switches', async () => {
  const editor = sorterEditor();
  const original = editor.api.PowerPoint.run;
  let fail = false;
  editor.api.PowerPoint.run = body => { if (fail) { fail = false; return Promise.reject(new Error('editor down')); } return original(body); };
  const push = editor.calls.push.bind(editor.calls);
  editor.calls.push = (...items) => { for (const item of items) if (item[0] === 'addSlide') fail = true; return push(...items); };
  await assert.rejects(sorterRun(editor), /editor down/);
  assert.deepEqual(editor.views, [SORTER, NORMAL]);

  const stopped = sorterEditor();
  const controller = new AbortController();
  controller.abort();
  await assert.rejects(sorterRun(stopped, twoNew, { signal: controller.signal }), error => error.name === 'AbortError');
  assert.deepEqual(stopped.views, [], 'no flash of the sorter');
});
// The `halted` path (another track failed) belongs to the parallel flow before v1.1; the sorter is only used with a locale.
