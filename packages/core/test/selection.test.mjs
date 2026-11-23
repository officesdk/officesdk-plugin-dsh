import test from 'node:test';
import assert from 'node:assert/strict';
import { DOCX_ALL_OPS, SELECTION_LIMITS, SuiteError, cellKey, cpLength, cpSlice, planDocRevert, planDocRewrite, planSheetRevert, planSheetRewrite, selectedText, selectionMessage, sheetAddress, validateDocxOps, validateSelectionRequest, valueKey } from '../src/index.mjs';

const doc = texts => texts.map(text => ({ text }));
const code = name => error => error instanceof SuiteError && error.code === name;

test('offsets count code points: an emoji is one, as the Writer API counts them', () => {
  assert.equal(cpLength('A😀B中文'), 5);
  assert.equal(cpSlice('A😀B中文', 1, 3), '😀B');
  assert.equal(selectedText(['A😀B中文加粗'], { paragraphs: [0, 0], start: 2, end: 4 }), 'B中');
  assert.equal(selectedText(['123', '456', '789'], { paragraphs: [0, 2], start: 1, end: 2 }), '23\n456\n78');
});

test('a selection inside one paragraph is replaced in place; the rest of the paragraph is left alone', () => {
  const paragraphs = doc(['标题', 'A😀B中文加粗尾巴', '下一段']);
  const plan = planDocRewrite(paragraphs, { paragraphs: [1, 1], start: 3, end: 5 }, '中文', '英语句子😀');
  assert.deepEqual(plan, { mode: 'inline', paragraph: 1, start: 3, length: 2, text: '英语句子😀', after: { paragraphs: [1, 1], start: 3, end: 8 }, afterText: '英语句子😀' });
  // Revert is the same plan backwards, from where the new text is.
  const after = doc(['标题', 'A😀B英语句子😀加粗尾巴', '下一段']);
  assert.deepEqual(planDocRevert(after, plan.after, plan.afterText, '中文'), { mode: 'inline', paragraph: 1, start: 3, length: 5, text: '中文', after: { paragraphs: [1, 1], start: 3, end: 5 }, afterText: '中文' });
});

test('new text with line breaks becomes paragraphs: prefix and suffix stay, only the selected paragraphs change', () => {
  const paragraphs = doc(['前文', '123', '456', '789', '后文']);
  const anchor = { paragraphs: [1, 3], start: 1, end: 2 };
  // Three paragraphs become two.
  const fewer = planDocRewrite(paragraphs, anchor, '23\n456\n78', 'AB\nCD');
  assert.equal(fewer.mode, 'paragraphs');
  assert.deepEqual(fewer.ops, [{ op: 'setText', paragraph: 1, text: '1AB' }, { op: 'setText', paragraph: 2, text: 'CD9' }, { op: 'deleteParagraph', paragraph: 3 }]);
  assert.deepEqual(fewer.after, { paragraphs: [1, 2], start: 1, end: 2 });
  // The plan is a valid document_edit batch on that document.
  assert.doesNotThrow(() => validateDocxOps(fewer.ops, paragraphs.map(p => ({ ...p })) && { paragraphs, tables: [] }, { enabled: DOCX_ALL_OPS }));
  // Three become four: inserted after the last selected paragraph, chained by refs; numbers outside never appear.
  const more = planDocRewrite(paragraphs, anchor, '23\n456\n78', 'a\nb\nc\nd');
  assert.deepEqual(more.ops, [
    { op: 'setText', paragraph: 1, text: '1a' }, { op: 'setText', paragraph: 2, text: 'b' }, { op: 'setText', paragraph: 3, text: 'c' },
    { op: 'insertParagraph', after: 3, ref: 'sel3', text: 'd9' }]);
  assert.ok(more.ops.every(op => op.paragraph === undefined || (op.paragraph >= 1 && op.paragraph <= 3)));
  // One paragraph, but the model answered with two: split, not a line break inside the paragraph.
  const split = planDocRewrite(doc(['甲乙丙丁']), { paragraphs: [0, 0], start: 1, end: 3 }, '乙丙', 'X\nY');
  assert.deepEqual(split.ops, [{ op: 'setText', paragraph: 0, text: '甲X' }, { op: 'insertParagraph', after: 0, ref: 'sel1', text: 'Y丁' }]);
  assert.deepEqual(split.after, { paragraphs: [0, 1], start: 1, end: 1 });
  // \r\n from the model counts as one break.
  assert.equal(planDocRewrite(doc(['甲乙丙丁']), { paragraphs: [0, 0], start: 1, end: 3 }, '乙丙', 'X\r\nY').ops.length, 2);
  // A rewrite into many paragraphs can still be taken back (its anchor spans more than a selection may).
  const many = Array.from({ length: 30 }, (_, i) => `p${i}`).join('\n');
  const wide = planDocRewrite(doc(['甲乙丙丁']), { paragraphs: [0, 0], start: 1, end: 3 }, '乙丙', many);
  const widened = doc(['甲p0', ...Array.from({ length: 28 }, (_, i) => `p${i + 1}`), 'p29丁']);
  assert.deepEqual(planDocRevert(widened, wide.after, wide.afterText, '乙丙').ops.at(0), { op: 'setText', paragraph: 0, text: '甲乙丙丁' });
  // Same text: nothing to write.
  assert.equal(planDocRewrite(paragraphs, anchor, '23\n456\n78', '23\n456\n78').mode, 'unchanged');
});

test('the selection must still hold what the user selected; table text is refused', () => {
  const paragraphs = doc(['前文', '123']);
  assert.throws(() => planDocRewrite(paragraphs, { paragraphs: [1, 1], start: 0, end: 2 }, '13', 'x'), code('source-changed'));
  assert.throws(() => planDocRewrite(paragraphs, { paragraphs: [1, 2], start: 0, end: 1 }, '1', 'x'), code('source-changed'));
  assert.throws(() => planDocRewrite(paragraphs, { paragraphs: [1, 1], start: 5, end: 6 }, '', 'x'), code('source-changed'));
  assert.throws(() => planDocRewrite([{ text: '格子', table: true }], { paragraphs: [0, 0], start: 0, end: 2 }, '格子', 'x'), code('invalid-edit'));
  assert.throws(() => planDocRewrite(paragraphs, { paragraphs: [1, 1], start: 0, end: 2 }, '12', 'x'.repeat(SELECTION_LIMITS.maxRewrite + 1)), code('invalid-edit'));
  assert.throws(() => planDocRewrite(paragraphs, { paragraphs: [1, 1], start: 0, end: 2 }, '12', 7), code('invalid-edit'));
});

test('selection requests are checked and normalized', () => {
  const long = '长'.repeat(1500);
  const ok = validateSelectionRequest({ kind: 'docx', anchor: { paragraphs: [2, 2], start: 0, end: 1500 }, original: long, instruction: '  改得更正式 ' });
  assert.equal(ok.instruction, '改得更正式');
  assert.throws(() => validateSelectionRequest({ kind: 'docx', anchor: { paragraphs: [0, 0], start: 0, end: 1 }, original: 'a', instruction: ' ' }), code('invalid-edit'));
  assert.throws(() => validateSelectionRequest({ kind: 'docx', anchor: { paragraphs: [0, 0], start: 0, end: 1 }, original: '   ', instruction: 'x' }), code('invalid-edit'));
  assert.throws(() => validateSelectionRequest({ kind: 'docx', anchor: { paragraphs: [0, 1], start: 0, end: 1 }, original: 'one line', instruction: 'x' }), code('invalid-edit'));
  assert.throws(() => validateSelectionRequest({ kind: 'docx', anchor: { paragraphs: [0, 0], start: 0, end: 1 }, original: 'x'.repeat(SELECTION_LIMITS.maxText + 1), instruction: 'x' }), code('invalid-edit'));
  assert.throws(() => validateSelectionRequest({ kind: 'docx', anchor: { paragraphs: [0, 30], start: 0, end: 1 }, original: 'a\n'.repeat(30) + 'b', instruction: 'x' }), code('invalid-edit'));
  assert.throws(() => validateSelectionRequest({ kind: 'pptx', instruction: 'x' }), code('invalid-edit'));
  const sheet = validateSelectionRequest({ kind: 'xlsx', anchor: { sheet: 'Sheet1', row: 1, column: 1, rowCount: 1, columnCount: 2 }, cells: [[{ text: '1', raw: 1 }, { text: '3', raw: 3, formula: '=A1+2' }]], instruction: 'x' });
  assert.deepEqual(sheet.cells, [[{ text: '1', raw: 1 }, { text: '3', raw: 3, formula: 'A1+2' }]]);
  assert.equal(sheet.original, '1\t=A1+2');
  assert.throws(() => validateSelectionRequest({ kind: 'xlsx', anchor: { sheet: 'S', row: 0, column: 0, rowCount: 20, columnCount: 11 }, cells: [], instruction: 'x' }), code('invalid-edit'));
  assert.throws(() => validateSelectionRequest({ kind: 'xlsx', anchor: { sheet: 'S', row: 0, column: 0, rowCount: 1, columnCount: 2 }, cells: [[{ text: 'a' }]], instruction: 'x' }), code('invalid-edit'));
});

test('the chat message carries the whole selection, the id and the request', () => {
  const original = '第一句很长。'.repeat(120);
  const text = selectionMessage({ id: 'sel-0a1b2c3d', name: '报告.docx', request: { kind: 'docx', anchor: { paragraphs: [2, 3], start: 0, end: 3 }, original: `${original}\n第二段`, instruction: '改得更正式' } });
  assert.ok(text.includes('sel-0a1b2c3d') && text.includes('第 3–4 段') && text.includes(`> ${original}\n> 第二段`) && text.includes('要求：改得更正式') && text.includes('selection_rewrite'));
  assert.ok(cpLength(text) > 700, 'nothing is cut short');
  const sheet = selectionMessage({ id: 'sel-0a1b2c3d', name: '销售.xlsx', lang: 'en', request: { kind: 'xlsx', anchor: { sheet: 'Sheet1', row: 1, column: 1, rowCount: 2, columnCount: 2 }, cells: [[{ text: 'a|b', raw: 'a|b' }, { text: '', raw: null }], [{ text: '3', raw: 3, formula: 'A1+2' }, { text: '4', raw: 4 }]], instruction: 'Translate' } });
  assert.ok(sheet.includes('Sheet1!B2:C3') && sheet.includes('| B | C |') && sheet.includes('| a\\|b |  |') && sheet.includes('| =A1+2 | 4 |'));
  assert.equal(sheetAddress({ row: 0, column: 0, rowCount: 1, columnCount: 1 }), 'A1');
});

test('spreadsheet rewrite: same size only, cells must be unchanged, revert restores values and formulas', () => {
  const anchor = { sheet: 'Sheet1', row: 1, column: 1, rowCount: 1, columnCount: 2 };
  const cells = [[{ text: 'hello', raw: 'hello' }, { text: '3', raw: 3, formula: 'A1+2' }]];
  const { step } = planSheetRewrite([[{ text: 'hello', raw: 'hello' }, { text: '3', raw: 3, formula: '=a1 + 2' }]], anchor, cells, [['你好', 5]]);
  assert.deepEqual(step, { op: 'setValues', sheet: 'Sheet1', range: { row: 1, column: 1, rowCount: 1, columnCount: 2 }, values: [['你好', 5]] });
  assert.throws(() => planSheetRewrite([[{ text: 'hello', raw: 'hello' }, { text: '3', raw: 3, formula: 'A1+2' }]], anchor, cells, [['x']]), code('invalid-edit'));
  assert.throws(() => planSheetRewrite([[{ text: 'hi', raw: 'hi' }, { text: '3', raw: 3, formula: 'A1+2' }]], anchor, cells, [['x', 'y']]), code('source-changed'));
  // Revert: the written values must still be there; the recorded cells go back, formulas with "=".
  const back = planSheetRevert([[{ text: '你好', raw: '你好' }, { text: '5', raw: '5' }]], anchor, [['你好', 5]], cells);
  assert.deepEqual(back.step.values, [['hello', '=A1+2']]);
  assert.throws(() => planSheetRevert([[{ text: 'edited', raw: 'edited' }, { text: '5', raw: 5 }]], anchor, [['你好', 5]], cells), code('source-changed'));
  // Dates and text starting with "=" go back with their own type, not through setValues.
  const special = [[{ text: '2026/11/9', raw: 46335, date: true }, { text: '=not a formula', raw: '=not a formula' }]];
  const typed = planSheetRevert([[{ text: 'a', raw: 'a' }, { text: 'b', raw: 'b' }]], anchor, [['a', 'b']], special);
  assert.deepEqual(typed.step.values, [[null, null]]);
  assert.deepEqual(typed.specials, [{ r: 0, c: 0, value: { type: 'date', value: 46335 } }, { r: 0, c: 1, value: { type: 'primitive', value: '=not a formula' } }]);
  // What spreadsheet_edit would refuse is refused before anything is written.
  assert.throws(() => planSheetRewrite([[{ text: 'hello', raw: 'hello' }, { text: '3', raw: 3, formula: 'A1+2' }]], anchor, cells, [['=', 'x']]), code('invalid-edit'));
  assert.throws(() => planSheetRewrite([[{ text: 'hello', raw: 'hello' }, { text: '3', raw: 3, formula: 'A1+2' }]], anchor, cells, [['==x', 'x']]), code('invalid-edit'));
  assert.equal(cellKey({ text: '', raw: null }), valueKey(null));
  assert.equal(cellKey({ text: 'x', formula: 'SUM(A1:A2)' }), valueKey('=sum(A1:A2)'));
});
