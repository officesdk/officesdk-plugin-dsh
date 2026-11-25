import test from 'node:test';
import assert from 'node:assert/strict';
import { DOCX_ALL_OPS, DOCX_HELD_OPS, DOCX_LIMITS, DOCX_OPS, checkDocxToken, describeDocxSteps, docxNeedsToken, docxProgress, docxReadToken, summarizeDocument, validateDocxOps as check } from '../src/index.mjs';
import { boxOf, createKitDocumentAgent as createAgent, findParagraphElements, inspectDocument, runDocxSteps } from '../src/browser/document-agent.mjs';

// Held operations (tables, office-kit K14) are implemented and tested here; the model is not offered them (see below).
const validateDocxOps = (ops, doc) => check(ops, doc, { enabled: DOCX_ALL_OPS });
const createKitDocumentAgent = options => createAgent({ enabled: DOCX_ALL_OPS, ...options });

test('tables are held back from the model until Kit can save them', () => {
  assert.deepEqual([...DOCX_HELD_OPS], ['insertTable', 'setCell']);
  assert.ok(!DOCX_OPS.includes('insertTable') && DOCX_OPS.includes('insertParagraph'));
  assert.throws(() => check([{ op: 'insertTable', values: [['a']] }], { paragraphs: [], tables: [] }), error => error.code === 'invalid-edit' && /not available yet/.test(error.message));
});

/**
 * A stand-in for the editor's Word JS API: paragraphs and tables in document order, every mutating call recorded in
 * `calls`. Style ids the document lacks fail at sync like the real editor; old paragraph proxies are not modelled
 * (the agent takes fresh ones every step).
 */
function fakeWord({ paragraphs = [''], styles = ['Normal', 'Title', 'Heading1', 'Heading2', 'Heading3', 'Subtitle', 'Quote'], marked = true, refuseDelete = false, failText } = {}) {
  const calls = [];
  let lists = 0;
  // A paragraph is its text, or { text, style } for one in another style (the redhead template's layout paragraphs).
  const doc = paragraphs.map(item => ({ text: item?.text ?? item, style: item?.style ?? 'Normal', list: null, table: 0, fmt: {} }));
  const tables = [];
  let pending;
  const record = (text, extra = {}) => ({ text, style: 'Normal', list: null, table: 0, fmt: {}, ...extra });
  const insertAt = (at, items) => doc.splice(at, 0, ...items);
  const newTable = (at, rows, columns, values) => {
    const cells = values.flatMap(row => row.map(value => record(String(value), { table: 1 })));
    insertAt(at, cells);
    const table = { cells, rows, columns };
    // Tables in document order, by where their first cell sits.
    tables.push(table);
    tables.sort((a, b) => doc.indexOf(a.cells[0]) - doc.indexOf(b.cells[0]));
    calls.push(['insertTable', at, rows, columns]);
    return {
      getCell: (row, column) => ({ body: { paragraphs: { getFirst: () => proxy(cells[row * columns + column]) } } }),
    };
  };
  // Like office-kit's Writer proxies: assignments are refused, `set()` works (and a nested font in a paragraph's set is not).
  const refuse = { set(_, key) { throw new TypeError(`'set' on proxy: trap returned falsish for property '${String(key)}'`); } };
  const font = item => new Proxy({ set(values) { for (const [key, value] of Object.entries(values)) { item.fmt[key] = value; calls.push(['font', item.text, key, value]); } } }, refuse);
  const proxy = item => new Proxy({
    get text() { return item.text; }, get style() { return item.style; }, get isListItem() { return Boolean(item.list); }, get tableNestingLevel() { return item.table; },
    set(values) {
      if ('font' in values) { pending = Object.assign(new Error('paragraph.font is not mapped to a Writer document action yet'), { code: 'UnsupportedCapability' }); return; }
      for (const [key, value] of Object.entries(values)) {
        if (key === 'styleBuiltIn' || key === 'style') {
          if (!styles.includes(value)) { pending = Object.assign(new Error(`style.apply style "${value}" does not exist.`), { code: 'NotAllowed' }); return; }
          item.style = value; calls.push([key === 'style' ? 'styleId' : 'style', item.text, value]);
        } else { item.fmt[key === 'outlineLevel' ? 'outlineLevel' : key] = value; if (key === 'alignment') calls.push(['alignment', item.text, value]); }
      }
    },
    font: font(item),
    insertText(text, where) {
      if (failText && text === failText) { pending = Object.assign(new Error('the editor refused that text'), { code: 'GeneralException' }); return; }
      if (where === 'End') { item.text += text; calls.push(['append', text]); return; }
      assert.equal(where, 'Replace'); item.text = text; calls.push(['replace', text]);
    },
    insertParagraph(text, where) { const at = doc.indexOf(item) + (where === 'After' ? 1 : 0); const created = record(text, { list: item.list && where === 'After' ? item.list : null }); insertAt(at, [created]); calls.push(['insertParagraph', text, where]); return proxy(created); },
    insertTable(rows, columns, where, values) { return newTable(doc.indexOf(item) + (where === 'After' ? 1 : 0), rows, columns, values); },
    insertBreak(kind, where) { insertAt(doc.indexOf(item) + (where === 'After' ? 1 : 0), [record('\f')]); calls.push(['break', kind, where]); },
    delete() { if (refuseDelete) return; doc.splice(doc.indexOf(item), 1); calls.push(['delete', item.text]); },
    getRange: () => ({ select: () => calls.push(['select', item.text]) }),
    get list() { return { id: item.list?.id, levelTypes: [item.list?.kind === 'bullet' ? 'Bullet' : 'Number'], load() {} }; },
    startNewList() { item.list = { id: ++lists, kind: 'number' }; calls.push(['startNewList', item.text]); return { setLevelBullet: (level, kind) => { item.list.kind = 'bullet'; calls.push(['bullet', item.text, level, kind]); } }; },
    attachToList(id) { item.list = { id, kind: doc.find(other => other.list?.id === id)?.list.kind }; calls.push(['attachToList', item.text, id]); },
    detachFromList() { item.list = null; calls.push(['detachFromList', item.text]); },
  }, refuse);
  const body = {
    get paragraphs() { return { load() {}, get items() { return doc.map(proxy); }, getLast: () => proxy(doc.at(-1)) }; },
    get tables() { return { load() {}, get items() { return tables.map(table => ({ rowCount: table.rows, get values() { const values = []; for (let r = 0; r < table.rows; r++) values.push(table.cells.slice(r * table.columns, (r + 1) * table.columns).map(cell => cell.text)); return values; }, getCell: (row, column) => ({ set value(text) { calls.push(['cell.value ignored', row, column, text]); }, body: { paragraphs: { getFirst: () => proxy(table.cells[row * table.columns + column]) } } }) })); } }; },
    insertParagraph(text, where) { const created = record(text); insertAt(where === 'Start' ? 0 : doc.length, [created]); calls.push(['insertParagraph', text, where]); return proxy(created); },
    insertTable(rows, columns, where, values) { return newTable(where === 'Start' ? 0 : doc.length, rows, columns, values); },
  };
  const Word = {
    async run(callback) {
      const ctx = { document: { body, getStyles: () => ({ load() {}, items: [...styles, ...(marked ? ['Shimo Live'] : [])].map(nameLocal => ({ nameLocal })) }) }, async sync() { if (pending) { const error = pending; pending = undefined; throw error; } } };
      return callback(ctx);
    },
  };
  return { Word, doc, tables, calls };
}

const readOf = async fake => inspectDocument(fake.Word);

test('the read is plain data that postMessage can clone, even when the editor hands out proxies', async () => {
  const fake = fakeWord({ paragraphs: ['一'] });
  const proxied = { async run(callback) { return fake.Word.run(ctx => callback({ ...ctx, document: { body: { get paragraphs() { const real = ctx.document.body.paragraphs; return { load() {}, get items() { return new Proxy(real.items, {}); } }; }, get tables() { return { load() {}, items: new Proxy([{ rowCount: 1, values: new Proxy([new Proxy(['x'], {})], {}) }], {}) }; } } } })); } };
  const doc = await inspectDocument(proxied);
  assert.deepEqual(structuredClone(doc), { paragraphs: [{ text: '一', style: 'Normal', list: false, table: false }], tables: [{ rows: 1, columns: 1, values: [['x']] }] });
});

test('readToken follows the text and table count; edits naming paragraphs must bring the current one', async () => {
  const fake = fakeWord({ paragraphs: ['一', '二'] });
  const doc = await readOf(fake);
  const token = docxReadToken(doc);
  assert.match(token, /^d[0-9a-f]{16}$/);
  assert.equal(docxReadToken(doc), token);
  assert.notEqual(docxReadToken({ ...doc, paragraphs: [{ text: '一' }, { text: '二!' }] }), token);
  assert.equal(docxNeedsToken([{ op: 'insertParagraph', text: 'x' }]), false);
  assert.equal(docxNeedsToken([{ op: 'setText', paragraph: 0, text: 'x' }]), true);
  assert.equal(docxNeedsToken([{ op: 'insertParagraph', after: 1, text: 'x' }]), true);
  checkDocxToken([{ op: 'insertParagraph', text: 'x' }], doc, undefined);
  assert.throws(() => checkDocxToken([{ op: 'setText', paragraph: 0, text: 'x' }], doc, undefined), error => error.code === 'invalid-edit' && /readToken/.test(error.message));
  assert.throws(() => checkDocxToken([{ op: 'setText', paragraph: 0, text: 'x' }], doc, 'd0000000000000000'), error => error.code === 'source-changed');
  checkDocxToken([{ op: 'setText', paragraph: 0, text: 'x' }], doc, token);
});

test('validation turns away bad batches before anything is written', async () => {
  const doc = { paragraphs: [{ text: '甲' }, { text: '乙', table: true }], tables: [{ rows: 1, columns: 1, values: [['乙']] }] };
  const bad = (ops, pattern) => assert.throws(() => validateDocxOps(ops, doc), error => error.code === 'invalid-edit' && pattern.test(error.message));
  bad([], /at least one/);
  bad(Array.from({ length: DOCX_LIMITS.maxOps + 1 }, () => ({ op: 'insertParagraph', text: 'x' })), /At most/);
  bad([{ op: 'draw' }], /unknown operation/);
  bad([{ op: 'insertParagraph', text: 'a\nb' }], /no line breaks/);
  bad([{ op: 'insertParagraph', text: 'x', after: 0, at: 'end' }], /only one of/);
  bad([{ op: 'insertParagraph', text: 'x', after: 1 }], /inside a table/);
  bad([{ op: 'setText', paragraph: 5, text: 'x' }], /not in the document/);
  bad([{ op: 'deleteParagraph', paragraph: 0 }, { op: 'setText', paragraph: 0, text: 'x' }], /deleted earlier/);
  bad([{ op: 'deleteParagraph', paragraph: 1 }], /inside a table/);
  assert.throws(() => validateDocxOps([{ op: 'deleteParagraph', paragraph: 0 }], { paragraphs: [{ text: '唯一' }], tables: [] }), /at least one paragraph/);
  bad([{ op: 'setText', paragraphRef: 'nope', text: 'x' }], /not a paragraph added earlier/);
  bad([{ op: 'insertParagraph', text: 'x', ref: 'a' }, { op: 'insertParagraph', text: 'y', ref: 'a' }], /already used/);
  bad([{ op: 'replaceText', paragraph: 0, find: '丙', replace: 'x' }], /does not occur/);
  bad([{ op: 'formatParagraph', paragraph: 0 }], /at least one/);
  bad([{ op: 'formatParagraph', paragraph: 0, align: 'middle' }], /align/);
  bad([{ op: 'formatParagraph', paragraph: 0, color: 'red' }], /#RRGGBB/);
  bad([{ op: 'insertParagraph', text: 'x', style: 'heading9' }], /style must be/);
  bad([{ op: 'insertTable', values: [['a', 'b'], ['c']] }], /row 2/);
  bad([{ op: 'insertTable', values: Array.from({ length: 31 }, () => ['a']) }], /1–30 rows/);
  bad([{ op: 'setCell', table: 0, row: 1, column: 0, text: 'x' }], /row must be/);
  bad([{ op: 'insertParagraph', text: 'x'.repeat(DOCX_LIMITS.maxParagraphText + 1) }], /longer than/);
  // Tables must not touch another table; nothing goes in front of a table that opens the document.
  const tabled = { paragraphs: [{ text: 'c', table: true }, { text: '后' }, { text: '再后' }], tables: [{ rows: 1, columns: 1, values: [['c']] }] };
  assert.throws(() => validateDocxOps([{ op: 'insertTable', values: [['x']], before: 1 }], tabled), /would touch the table/);
  assert.throws(() => validateDocxOps([{ op: 'insertPageBreak', at: 'start' }], tabled), /starts with a table/);
  assert.equal(validateDocxOps([{ op: 'insertTable', values: [['x']], after: 1 }], tabled).steps.length, 1);
});

test('validation resolves targets, refs and replaceText, and flags deletions for the card', () => {
  const doc = { paragraphs: [{ text: '第一段文字' }, { text: '第二段' }], tables: [] };
  const { steps, danger, reasons } = validateDocxOps([
    { op: 'insertParagraph', text: '标题', style: 'title', at: 'start', ref: 'title' },
    { op: 'insertParagraph', text: '要点', list: 'bullet', afterRef: 'title', ref: 'p1' },
    { op: 'replaceText', paragraph: 0, find: '一段', replace: '1 段' },
    { op: 'formatParagraph', paragraphRef: 'p1', bold: true, align: 'center' },
    { op: 'deleteParagraph', paragraph: 1 },
    { op: 'insertTable', values: [['名称', 3]], header: true, ref: 't' },
    { op: 'setCell', tableRef: 't', row: 0, column: 1, text: '4' },
    { op: 'insertPageBreak' },
  ], doc);
  assert.deepEqual(steps[0], { op: 'insertParagraph', where: { at: 'start' }, text: '标题', ref: 'title', style: 'title' });
  assert.deepEqual(steps[1].where, { after: { ref: 'title' } });
  assert.deepEqual(steps[2], { op: 'setText', target: { paragraph: 0 }, before: '第一段文字', text: '第1 段文字', find: '一段', replace: '1 段' });
  assert.deepEqual(steps[5].values, [['名称', '3']]);
  assert.equal(danger, true);
  assert.deepEqual(reasons, ['deletes 1 paragraph']);
  const rows = describeDocxSteps(steps);
  assert.deepEqual(rows.map(row => [row.op, row.where]), [['insertParagraph', 'start @title'], ['insertParagraph', 'after @title @p1'], ['replaceText', '¶1'], ['formatParagraph', '@p1'], ['deleteParagraph', '¶2'], ['insertTable', 'end'], ['setCell', '@t R1C2'], ['insertPageBreak', 'end']]);
  assert.equal(rows[2].before, '一段');
  assert.equal(rows[2].after, '1 段');
});

test('summarizeDocument numbers paragraphs, keeps tables and cuts long text', () => {
  const summary = summarizeDocument({ paragraphs: [{ text: 'a'.repeat(30), style: 'Heading1' }, { text: 'b', list: true }, { text: 'c', table: true, style: 'Normal' }], tables: [{ rows: 1, columns: 1, values: [['c']] }] }, { maxParagraph: 10 });
  assert.deepEqual(summary.paragraphs, [{ index: 0, text: 'a'.repeat(10), style: 'Heading1' }, { index: 1, text: 'b', list: true }, { index: 2, text: 'c', table: true }]);
  assert.deepEqual(summary.tables, [{ index: 0, rows: 1, columns: 1, values: [['c']] }]);
  assert.equal(summary.truncated, true);
  assert.equal(summary.paragraphCount, 3);
  assert.match(summary.readToken, /^d/);
});

test('docxProgress says how far a batch got and that nothing was undone', () => {
  assert.equal(docxProgress(0, 3), 'None of the 3 steps was applied.');
  assert.match(docxProgress(1, 3, {}, 1), /Steps 1–1 of 3 were applied .* Step 2 was partly applied/);
  assert.equal(docxProgress(2, 5, { a: { paragraph: 4 } }), 'Steps 1–2 of 5 were applied and stay in the editor: nothing was undone and nothing was saved. refs: {"a":{"paragraph":4}}');
});

test('a new document fills its empty paragraph first, then appends with styles, lists and formats', async () => {
  const fake = fakeWord({ paragraphs: [''] });
  const doc = await readOf(fake);
  const { steps } = validateDocxOps([
    { op: 'insertParagraph', text: '关于开展检查的通知', style: 'title', ref: 'title' },
    { op: 'insertParagraph', text: '一、总体要求', style: 'heading1' },
    { op: 'insertParagraph', text: '第一项', list: 'bullet' },
    { op: 'insertParagraph', text: '第二项', list: 'bullet' },
    { op: 'insertParagraph', text: '正文', firstLineIndent: 32, align: 'justify' },
  ], doc);
  const { refs, styleFallback } = await runDocxSteps(fake.Word, steps, doc);
  assert.deepEqual(fake.doc.map(item => item.text), ['关于开展检查的通知', '一、总体要求', '第一项', '第二项', '正文']);
  assert.deepEqual(fake.doc.map(item => item.style), ['Title', 'Heading1', 'Normal', 'Normal', 'Normal']);
  assert.equal(fake.doc[2].list.kind, 'bullet');
  assert.equal(fake.doc[3].list.id, fake.doc[2].list.id, 'the second item joins the list before it');
  assert.equal(fake.doc[4].list, null);
  assert.deepEqual(fake.doc[4].fmt, { firstLineIndent: 32, alignment: 'Justified' });
  assert.deepEqual(refs, { title: { paragraph: 0 } });
  assert.equal(styleFallback, 0);
  assert.ok(fake.calls.some(call => call[0] === 'select'), 'the caret follows the writing');
});

test('only the first insert fills a new document\'s empty paragraph; a list of another kind starts its own list', async () => {
  const fake = fakeWord({ paragraphs: [''] });
  const doc = await readOf(fake);
  const { steps } = validateDocxOps([{ op: 'setText', paragraph: 0, text: '标题' }, { op: 'insertParagraph', text: '正文' }, { op: 'insertParagraph', text: '甲', list: 'bullet' }, { op: 'insertParagraph', text: '一', list: 'number' }], doc);
  await runDocxSteps(fake.Word, steps, doc);
  assert.deepEqual(fake.doc.map(item => item.text), ['标题', '正文', '甲', '一']);
  assert.notEqual(fake.doc[3].list.id, fake.doc[2].list.id);
  assert.equal(fake.doc[3].list.kind, 'number');
});

test('261109 redhead: the two layout paragraphs are rewritten in place and the text goes on below the red rule, without an empty line', async () => {
  const fake = fakeWord({ paragraphs: [{ text: '【发文机关】文件', style: 'ShimoIssuer' }, { text: '【发文字号】', style: 'ShimoDocNumber' }], styles: ['Normal', 'Title', 'Heading1', 'ShimoIssuer', 'ShimoDocNumber'] });
  const doc = await readOf(fake);
  const { steps } = validateDocxOps([
    { op: 'setText', paragraph: 0, text: '××市教育局文件' },
    { op: 'setText', paragraph: 1, text: '×教发〔2026〕12号' },
    { op: 'insertParagraph', text: '关于开展校园安全检查的通知', style: 'title' },
    { op: 'insertParagraph', text: '各区县教育局：', firstLineIndent: 0 },
    { op: 'insertParagraph', text: '一、检查范围', style: 'heading1' },
  ], doc);
  await runDocxSteps(fake.Word, steps, doc);
  assert.deepEqual(fake.doc.map(item => [item.text, item.style]), [
    ['××市教育局文件', 'ShimoIssuer'], ['×教发〔2026〕12号', 'ShimoDocNumber'],
    ['关于开展校园安全检查的通知', 'Title'], ['各区县教育局：', 'Normal'], ['一、检查范围', 'Heading1'],
  ]);
});

test('261109 redhead: a paragraph whose built-in style is missing never takes the red layout style as its body style', async () => {
  // No Normal in this document: the body paragraph falls back to direct formatting from the most used body style,
  // which must not be the template's red 42 pt issuer line.
  const fake = fakeWord({ paragraphs: [{ text: '××市教育局文件', style: 'ShimoIssuer' }, { text: '×教发〔2026〕12号', style: 'ShimoDocNumber' }], styles: ['ShimoIssuer', 'ShimoDocNumber'] });
  const doc = await readOf(fake);
  const { steps } = validateDocxOps([{ op: 'insertParagraph', text: '正文' }], doc);
  const { styleFallback } = await runDocxSteps(fake.Word, steps, doc);
  assert.equal(styleFallback, 1);
  assert.equal(fake.doc[2].text, '正文');
  assert.ok(!fake.calls.some(call => call[0] === 'styleId' && /^Shimo/.test(call[2])), 'no layout style is applied to body text');
  assert.equal(fake.doc[2].style, 'Normal');
});

test('a step that fails after its insert reports itself as partly applied, with its ref', async () => {
  const fake = fakeWord({ paragraphs: ['A'], styles: [] });
  const doc = await readOf(fake);
  const editor = { api: { writer: { Word: fake.Word } }, prepare: async () => {} };
  // No style id and the direct formatting refused too: the paragraph is in, its look is not.
  const original = fake.Word.run;
  let runs = 0;
  fake.Word.run = async callback => { runs += 1; if (runs === 6) throw Object.assign(new Error('format refused'), { code: 'GeneralException' }); return original(callback); };
  const agent = createKitDocumentAgent({ editor, document: fakePage(), animate: () => false, wait: async () => {} });
  await assert.rejects(agent.apply({ ops: [{ op: 'insertParagraph', text: '新', ref: 'n', style: 'heading1' }] }),
    error => error.code === 'editor-failed' && /None of the 1 steps was applied\. Step 1 was partly applied.*refs: \{"n":\{"paragraph":1\}\}/.test(error.message));
  assert.deepEqual(fake.doc.map(item => item.text), ['A', '新']);
});

test('titles and headings outside a live template are formatted directly (their outline level would break saving)', async () => {
  const fake = fakeWord({ paragraphs: [''], marked: false });
  const doc = await readOf(fake);
  const { steps } = validateDocxOps([{ op: 'insertParagraph', text: '标题', style: 'title' }, { op: 'insertParagraph', text: '一、', style: 'heading1' }, { op: 'insertParagraph', text: '引文', style: 'quote' }], doc);
  const { styleFallback } = await runDocxSteps(fake.Word, steps, doc);
  assert.equal(styleFallback, 2);
  assert.deepEqual(fake.doc.map(item => item.style), ['Normal', 'Normal', 'Quote']);
  assert.deepEqual(fake.doc[0].fmt, { size: 22, bold: true, alignment: 'Centered' });
});

test('body text in a non-template document takes the look of the document\'s own body paragraph, found after earlier inserts', async () => {
  const fake = fakeWord({ paragraphs: ['已有正文'], marked: false });
  const original = fake.Word.run;
  // The reference paragraph reports its look when loaded; anything else reports none.
  fake.Word.run = callback => original(ctx => callback({ ...ctx, document: { ...ctx.document, body: new Proxy(ctx.document.body, { get(target, key) {
    if (key !== 'paragraphs') return target[key];
    const real = target.paragraphs;
    return { ...real, load() {}, get items() { return real.items.map(item => item.text === '已有正文' ? new Proxy(item, { get(t, k) { return k === 'alignment' ? 'Justified' : k === 'font' ? Object.assign(Object.create(null), { size: 15, set: t.font.set }) : k === 'load' ? () => {} : t[k]; } }) : item); } };
  } }) } }));
  const doc = await readOf(fake);
  const { steps } = validateDocxOps([{ op: 'insertParagraph', text: '大标题', style: 'title', at: 'start' }, { op: 'insertParagraph', text: '新正文' }], doc);
  await runDocxSteps(fake.Word, steps, doc);
  assert.deepEqual(fake.doc.map(item => item.text), ['大标题', '已有正文', '新正文']);
  assert.deepEqual(fake.doc[2].fmt, { bold: false, italic: false, alignment: 'Justified', size: 15 });
});

test('a document without the style id gets its body style and direct formatting instead', async () => {
  const fake = fakeWord({ paragraphs: ['已有', '也有'], styles: ['a', 'a3'] });
  fake.doc.forEach(item => { item.style = 'a'; });
  fake.doc[1].style = 'a3';
  fake.doc.push({ text: '第三', style: 'a', list: null, table: 0, fmt: {} });
  const doc = await readOf(fake);
  const { steps } = validateDocxOps([{ op: 'insertParagraph', text: '标题', style: 'heading1', after: 1 }], doc);
  fake.doc[1].style = 'a3';
  const { styleFallback } = await runDocxSteps(fake.Word, steps, doc);
  assert.equal(styleFallback, 1);
  assert.deepEqual(fake.doc[2], { text: '标题', style: 'a', list: null, table: 0, fmt: { size: 16, bold: true } }, 'the most common body style id, then the heading look (no outline level: K15)');
});

test('numbers from the read keep naming the same paragraphs after inserts, tables and deletions', async () => {
  const fake = fakeWord({ paragraphs: ['A', 'B', 'C', 'D'] });
  const doc = await readOf(fake);
  const { steps } = validateDocxOps([
    { op: 'insertTable', values: [['x', 'y'], ['z', 'w']], after: 0, ref: 't' },
    { op: 'insertParagraph', text: 'B之前', before: 1 },
    { op: 'deleteParagraph', paragraph: 2 },
    { op: 'setText', paragraph: 3, text: 'D改' },
    { op: 'setCell', tableRef: 't', row: 1, column: 1, text: 'W' },
    { op: 'insertParagraph', text: 'A之后', after: 0 },
  ], doc);
  const { refs } = await runDocxSteps(fake.Word, steps, doc);
  assert.deepEqual(fake.doc.map(item => item.text), ['A', 'A之后', 'x', 'y', 'z', 'W', 'B之前', 'B', 'D改']);
  assert.deepEqual(refs, { t: { table: 0 } });
});

test('a step whose paragraph count does not change as planned fails instead of drifting', async () => {
  const fake = fakeWord({ paragraphs: ['A', 'B'], refuseDelete: true });
  const doc = await readOf(fake);
  const { steps } = validateDocxOps([{ op: 'deleteParagraph', paragraph: 0 }, { op: 'setText', paragraph: 1, text: 'x' }], doc);
  await assert.rejects(runDocxSteps(fake.Word, steps, doc), error => error.code === 'editor-failed' && /count changed by 0/.test(error.message));
});

/** A page with no Writer DOM: the agent still runs, the frame simply has nothing to outline. */
function fakePage() {
  const listeners = new Map();
  const element = () => ({ style: {}, dataset: {}, append() {}, remove() {}, setAttribute() {}, getBoundingClientRect: () => ({ width: 0 }), set textContent(value) { this._text = value; }, get textContent() { return this._text; } });
  return {
    defaultView: { requestAnimationFrame: () => 1, cancelAnimationFrame() {}, performance: { now: () => 0 }, addEventListener: (type, fn) => listeners.set(type, fn), removeEventListener() {} },
    body: { append() {} },
    createElement: element, createElementNS: element,
    addEventListener: (type, fn) => listeners.set(type, fn), removeEventListener() {},
    querySelectorAll: () => [], querySelector: () => null,
    listeners,
  };
}

const agentFor = (fake, extra = {}) => {
  const order = [];
  const editor = { api: { writer: { Word: fake.Word } }, prepare: async () => { order.push('prepare'); } };
  const agent = createKitDocumentAgent({ editor, document: fakePage(), animate: () => false, wait: async () => {}, ...extra });
  return { agent, order };
};

test('the agent applies a batch, returns refs and a fresh readToken, and is the AI until prepare has run', async () => {
  const fake = fakeWord({ paragraphs: ['旧'] });
  const { agent, order } = agentFor(fake);
  const seen = [];
  const doc = await agent.inspect();
  const pending = agent.apply({ ops: [{ op: 'insertParagraph', text: '新', ref: 'n' }, { op: 'setText', paragraph: 0, text: '旧改' }], readToken: docxReadToken(doc) });
  seen.push(agent.applying);
  const result = await pending;
  assert.deepEqual(seen, [true]);
  assert.equal(agent.applying, false);
  assert.deepEqual(order, ['prepare']);
  assert.equal(result.applied, 2);
  assert.deepEqual(result.refs, { n: { paragraph: 1 } });
  assert.equal(result.readToken, docxReadToken(await agent.inspect()));
});

test('concurrent results share a read and commit in ready order without rereading the other paragraph', async () => {
  const fake = fakeWord({ paragraphs: ['Alpha old', 'Beta old'] });
  const { agent } = agentFor(fake);
  const readToken = docxReadToken(await agent.inspect());
  const edit = (paragraph, text) => ({ ops: [{ op: 'setText', paragraph, text }], readToken });
  const second = agent.apply(edit(1, 'Beta ready'));
  const first = agent.apply(edit(0, 'Alpha ready'));
  await Promise.all([second, first]);
  assert.deepEqual(fake.doc.map(p => p.text), ['Alpha ready', 'Beta ready']);
  await assert.rejects(agent.inspect({ checkEdit: edit(0, 'stale') }), { code: 'source-changed' });
  await assert.rejects(agent.apply(edit(0, 'stale')), { code: 'source-changed' });
});

test('a late target declaration and replacement preserve a disjoint edit within a paragraph', async () => {
  const fake = fakeWord({ paragraphs: ['morning / evening', 'untouched'] });
  const { agent } = agentFor(fake);
  const readToken = docxReadToken(await agent.inspect());
  const replace = (find, replace) => ({ ops: [{ op: 'replaceText', paragraph: 0, find, replace }], readToken });
  await agent.apply(replace('morning', 'early sunrise'));
  const doc = await agent.inspect({ checkEdit: replace('evening', 'sunset') });
  const ref = agent.modelEditRange(doc, { paragraph: 0, find: 'evening', readToken });
  assert.equal(ref.anchor.start, 'early sunrise / '.length);
  await agent.apply(replace('evening', 'sunset'));
  assert.equal(fake.doc[0].text, 'early sunrise / sunset');
  await assert.rejects(agent.apply(replace('morning', 'stale')), { code: 'source-changed' });
  await assert.rejects(agent.apply({ ops: [{ op: 'deleteParagraph', paragraph: 1 }], readToken }), { code: 'source-changed' });
});

test('a stale readToken is refused on the page before any step', async () => {
  const fake = fakeWord({ paragraphs: ['旧'] });
  const { agent } = agentFor(fake);
  await assert.rejects(agent.apply({ ops: [{ op: 'setText', paragraph: 0, text: 'x' }], readToken: 'd0000000000000000' }), error => error.code === 'source-changed' && /None of the 1 steps was applied/.test(error.message));
  assert.deepEqual(fake.doc.map(item => item.text), ['旧']);
});

test('a failing step keeps the steps before it and says so; nothing is rolled back', async () => {
  const fake = fakeWord({ paragraphs: ['A', 'B'], failText: '坏' });
  const { agent } = agentFor(fake);
  const doc = await agent.inspect();
  await assert.rejects(agent.apply({ ops: [{ op: 'insertParagraph', text: '一', ref: 'one' }, { op: 'setText', paragraph: 1, text: '坏' }, { op: 'insertParagraph', text: '三' }], readToken: docxReadToken(doc) }),
    error => error.code === 'editor-failed' && /Steps 1–1 of 3 were applied and stay in the editor: nothing was undone and nothing was saved\. refs: \{"one":\{"paragraph":2\}\}/.test(error.message));
  assert.deepEqual(fake.doc.map(item => item.text), ['A', 'B', '一']);
  assert.equal(agent.applying, false);
});

test('stopping keeps what was written and reports cancelled with the step count', async () => {
  const fake = fakeWord({ paragraphs: [''] });
  const abort = new AbortController();
  let steps = 0;
  const { agent } = agentFor(fake, { wait: async () => { if (++steps === 2) abort.abort(); } });
  await assert.rejects(agent.apply({ ops: [{ op: 'insertParagraph', text: '一' }, { op: 'insertParagraph', text: '二' }, { op: 'insertParagraph', text: '三' }] }, abort.signal),
    error => error.code === 'cancelled' && /Stopped by the user\. Steps 1–2 of 3 were applied/.test(error.message));
  assert.deepEqual(fake.doc.map(item => item.text), ['一', '二']);
});

test('AI undo is not offered for Word', async () => {
  const { agent } = agentFor(fakeWord());
  await assert.rejects(agent.undo({ editId: 'e' }), error => error.code === 'unsupported-operation');
});

test('the frame finds a paragraph by its drawn lines, nearest block first, clipped to the scroll view', () => {
  const line = (key, text, rect) => ({ isConnected: true, textContent: text, getAttribute: name => name === 'data-paint-key' ? key : null, getBoundingClientRect: () => rect });
  const lines = [
    line('fragment:0/3/0/0:0-', '重复', { left: 10, top: 10, width: 50, height: 10, right: 60, bottom: 20 }),
    line('fragment:0/9/0/0:0-', '重复', { left: 10, top: 100, width: 50, height: 10, right: 60, bottom: 110 }),
    line('fragment:0/12/0/0:0-', '一段很长的', { left: 10, top: 200, width: 100, height: 10, right: 110, bottom: 210 }),
    line('fragment:0/12/1/0:0-', ' 文字', { left: 10, top: 212, width: 40, height: 10, right: 50, bottom: 222 }),
    line('fragment:1/12/0/0:0-', '一段很长的文字', { left: 0, top: 0, width: 1, height: 1, right: 1, bottom: 1 }),
  ];
  const page = { querySelectorAll: () => lines, querySelector: () => ({ getBoundingClientRect: () => ({ left: 0, top: 50, right: 500, bottom: 400, width: 500, height: 350 }) }) };
  assert.equal(findParagraphElements(page, '重复', 8)[0], lines[1]);
  assert.equal(findParagraphElements(page, '重复', 2)[0], lines[0]);
  const long = findParagraphElements(page, '一段很长的文字', 12);
  assert.deepEqual(long, [lines[2], lines[3]]);
  assert.deepEqual(boxOf(page, long), { left: 10, top: 200, width: 100, height: 22, clip: { left: 0, top: 50, right: 500, bottom: 400 } });
  assert.equal(findParagraphElements(page, '', 0), undefined);
  assert.equal(findParagraphElements(page, '没有', 0), undefined);
});

test('v1.1: each step selects before writing; new paragraphs grow in parts, up to three of them round-robin', async () => {
  const fake = fakeWord({ paragraphs: ['已有'] });
  const doc = await readOf(fake);
  const { steps } = validateDocxOps([
    { op: 'insertParagraph', text: 'A1A2', style: 'heading1' },
    { op: 'insertParagraph', text: 'B1B2' },
    { op: 'insertParagraph', text: 'C1C2' },
    { op: 'insertParagraph', text: 'D1D2' },
    { op: 'setText', paragraph: 0, text: '改' },
  ], doc);
  const log = [];
  const halves = text => text.length > 2 ? [text.slice(0, 2), text] : [text];
  await runDocxSteps(fake.Word, steps, doc, {
    chunk: halves,
    onSelect: async (number, target) => { log.push(`select ${number}${target?.insert ? ` ${target.insert} ${target.text}` : ''}`); },
    onWrite: (number, { text }) => log.push(`write ${number} ${text}`),
    onFinish: number => log.push(`done ${number}`),
  });
  assert.deepEqual(log, [
    'select 0 after 已有', 'write 0 A1', 'select 1 after A1', 'write 1 B1', 'select 2 after B1', 'write 2 C1',
    'write 0 A1A2', 'done 0', 'write 1 B1B2', 'done 1', 'write 2 C1C2', 'done 2',
    'select 3 after C1C2', 'write 3 D1', 'write 3 D1D2', 'done 3',
    'select 4', 'write 4 改', 'done 4',
  ], 'a fourth paragraph waits for the first three; another kind of step finishes them first');
  assert.deepEqual(fake.doc.map(item => item.text), ['改', 'A1A2', 'B1B2', 'C1C2', 'D1D2']);
});

test('a substring replacement reports only the old and new character spans, in code points', async () => {
  const fake = fakeWord({ paragraphs: ['🙂开始这里结束'] });
  const doc = await readOf(fake);
  const { steps } = validateDocxOps([{ op: 'replaceText', paragraph: 0, find: '这里', replace: '了' }], doc);
  const ranges = [];
  const capture = (_, target) => ranges.push([target.block, target.text, target.start, target.end]);
  await runDocxSteps(fake.Word, steps, doc, { onSelect: capture, onWrite: capture });
  assert.deepEqual(ranges, [[0, '🙂开始这里结束', 3, 5], [0, '🙂开始了结束', 3, 4]]);
});

test('v1.1: stopping while paragraphs grow counts only the finished ones and names the one left half written', async () => {
  const fake = fakeWord({ paragraphs: [''] });
  const abort = new AbortController();
  const { agent } = agentFor(fake);
  let writes = 0;
  const original = fake.Word.run;
  fake.Word.run = callback => original(async ctx => { const result = await callback(ctx); if (++writes === 9) abort.abort(); return result; });
  const long = '内容'.repeat(30);
  await assert.rejects(createKitDocumentAgent({ editor: { api: { writer: { Word: fake.Word } }, prepare: async () => {} }, document: fakePage(), animate: () => true, wait: async () => {} })
    .apply({ ops: [{ op: 'insertParagraph', text: long }, { op: 'insertParagraph', text: long }] }, abort.signal),
    error => error.code === 'cancelled' && /None of the 2 steps was applied\. Steps 1, 2 were partly applied/.test(error.message));
  assert.ok(fake.doc.some(item => item.text && item.text !== long), 'a paragraph stays half written: no rollback');
  void agent;
});

test('v1.1: a stop during the selection inserts nothing', async () => {
  const fake = fakeWord({ paragraphs: ['已有'] });
  const abort = new AbortController();
  const agent = createKitDocumentAgent({ editor: { api: { writer: { Word: fake.Word } }, prepare: async () => {} }, document: fakePage(), animate: () => true, wait: async () => abort.abort() });
  await assert.rejects(agent.apply({ ops: [{ op: 'insertParagraph', text: '新段落' }] }, abort.signal), error => error.code === 'cancelled' && /None of the 1 steps was applied\.$/.test(error.message));
  assert.deepEqual(fake.doc.map(item => item.text), ['已有']);
});

test('261109: a whole rewrite (setText) grows in parts; replaceText and reduced motion write once', async () => {
  const long = '改写后的段落内容'.repeat(5);
  const fake = fakeWord({ paragraphs: ['原来的第一段文字', '第二段里有一个词'] });
  const page = { editor: { api: { writer: { Word: fake.Word } }, prepare: async () => {} }, document: fakePage(), wait: async () => {} };
  const doc = await readOf(fake);
  await createKitDocumentAgent({ ...page, animate: () => true }).apply({ ops: [{ op: 'setText', paragraph: 0, text: long }, { op: 'replaceText', paragraph: 1, find: '一个词', replace: '两个词' }], readToken: docxReadToken(doc) });
  const writes = fake.calls.filter(call => call[0] === 'replace' || call[0] === 'append');
  const rewrite = writes.slice(0, -1);
  assert.equal(rewrite.length, 3, 'a 40-character rewrite is written in three parts: two states before the whole text');
  assert.equal(rewrite[0][0], 'replace');
  assert.ok(rewrite.slice(1).every(call => call[0] === 'append'), 'later parts are added at the end');
  assert.deepEqual(writes.at(-1), ['replace', '第二段里有两个词'], 'replaceText is one write');
  assert.deepEqual(fake.doc.map(item => item.text), [long, '第二段里有两个词']);
  const still = fakeWord({ paragraphs: ['原来的第一段文字'] });
  await createKitDocumentAgent({ ...page, editor: { api: { writer: { Word: still.Word } }, prepare: async () => {} }, animate: () => false }).apply({ ops: [{ op: 'setText', paragraph: 0, text: long }], readToken: docxReadToken(await readOf(still)) });
  assert.deepEqual(still.calls.filter(call => call[0] === 'replace' || call[0] === 'append'), [['replace', long]]);
});

test('261109: the frame follows the rewrite as it grows, and a stop mid-rewrite names that step as half written', async () => {
  const long = '改写后的段落内容'.repeat(5);
  const fake = fakeWord({ paragraphs: ['原来的第一段文字'] });
  const doc = await readOf(fake);
  const { steps } = validateDocxOps([{ op: 'setText', paragraph: 0, text: long }], doc);
  const written = [];
  await runDocxSteps(fake.Word, steps, doc, { chunk: text => [text.slice(0, 12), text.slice(0, 24), text], onWrite: (number, { text }) => written.push(text.length) });
  assert.deepEqual(written, [12, 24, 40]);
  const stopping = fakeWord({ paragraphs: ['原来的第一段文字'] });
  const abort = new AbortController();
  let pauses = 0;
  const agent = createKitDocumentAgent({ editor: { api: { writer: { Word: stopping.Word } }, prepare: async () => {} }, document: fakePage(), animate: () => true, wait: async () => { if (++pauses === 2) abort.abort(); } });
  await assert.rejects(agent.apply({ ops: [{ op: 'setText', paragraph: 0, text: long }], readToken: docxReadToken(await readOf(stopping)) }, abort.signal),
    error => error.code === 'cancelled' && /None of the 1 steps was applied\. Step 1 was partly applied/.test(error.message));
  assert.ok(stopping.doc[0].text.length > 0 && stopping.doc[0].text !== long && long.startsWith(stopping.doc[0].text), 'the half rewrite stays: no rollback');
});

test('261109: a rewrite of 24 characters shows one state before the whole text, one of 36 shows two', async () => {
  for (const [length, writes] of [[24, 2], [36, 3]]) {
    const fake = fakeWord({ paragraphs: ['原来'] });
    const text = '字'.repeat(length);
    await createKitDocumentAgent({ editor: { api: { writer: { Word: fake.Word } }, prepare: async () => {} }, document: fakePage(), animate: () => true, wait: async () => {} })
      .apply({ ops: [{ op: 'setText', paragraph: 0, text }], readToken: docxReadToken(await readOf(fake)) });
    assert.equal(fake.calls.filter(call => call[0] === 'replace' || call[0] === 'append').length, writes, `${length} characters`);
    assert.equal(fake.doc[0].text, text);
  }
});
