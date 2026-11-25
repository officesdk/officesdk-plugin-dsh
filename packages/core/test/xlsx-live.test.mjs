import test from 'node:test';
import assert from 'node:assert/strict';
import { columnIndex, columnLetters, describeXlsxSteps, parseRange, rangeText, summarizeWorkbook, validateXlsxOps, xlsxProgress } from '../src/index.mjs';
import { chartPosition, createKitSheetAgent, inspectWorkbook, runXlsxSteps } from '../src/browser/sheet-agent.mjs';

/** A stand-in for the patched `editor.api.sheet`: worksheets of cells `{ value?, formula?, ...format }`, calls recorded. */
function fakeSheet({ sheets = [{ name: 'Sheet1', rows: 50, columns: 10 }], failOn, bounding = true } = {}) {
  const calls = [];
  let nextId = 1;
  const book = sheets.map((sheet, index) => ({ id: `s${nextId++}`, name: sheet.name, active: index === 0, rowCount: sheet.rows, columnCount: sheet.columns, cells: new Map() }));
  const key = (row, column) => `${row},${column}`;
  const range = (sheet, { row, column, rowCount, columnCount }) => ({
    row, column, rowCount, columnCount,
    setValue(values) { if (failOn === 'setValue') throw new Error('refused'); values.forEach((line, r) => line.forEach((value, c) => { const at = key(row + r, column + c); if (value === null) sheet.cells.delete(at); else sheet.cells.set(at, { value }); })); calls.push(['setValue', sheet.name, rangeText({ row, column, rowCount, columnCount })]); },
    getData() { return Array.from({ length: rowCount }, (_, r) => Array.from({ length: columnCount }, (_, c) => { const cell = sheet.cells.get(key(row + r, column + c)); if (!cell) return { value: null, formula: null, text: '' }; const { value, formula, ...style } = cell; return { ...style, value: value === undefined ? null : typeof value === 'object' ? value : { type: 'primitive', value }, formula: formula?.slice(1) ?? null, text: formula ? '=' : String(value?.value ?? value ?? '') }; })); },
    setData(data) { data.forEach((line, r) => line.forEach((cell, c) => { const at = key(row + r, column + c); sheet.cells.set(at, { ...(sheet.cells.get(at) ?? {}), ...cell, ...(cell.formula && { formula: cell.formula }) }); })); calls.push(['setData', sheet.name, rangeText({ row, column, rowCount, columnCount })]); },
    select() { calls.push(['select', sheet.name, rangeText({ row, column, rowCount, columnCount })]); },
    // The page tests have no layout: without a box the agent's frame is simply not drawn.
    getBounding: () => bounding ? { left: 10, top: 20, width: 30, height: 40 } : null,
  });
  const worksheet = sheet => ({
    get id() { return sheet.id; }, get name() { return sheet.name; }, get isActive() { return sheet.active; }, get rowCount() { return sheet.rowCount; }, get columnCount() { return sheet.columnCount; },
    getRange: spec => range(sheet, spec),
    getUsedRange() { if (!sheet.cells.size) return null; let rows = 0, columns = 0; for (const at of sheet.cells.keys()) { const [r, c] = at.split(',').map(Number); rows = Math.max(rows, r + 1); columns = Math.max(columns, c + 1); } return range(sheet, { row: 0, column: 0, rowCount: rows, columnCount: columns }); },
    getCell: (row, column) => ({ setCellFormula(formula) { sheet.cells.set(key(row, column), { formula }); calls.push(['formula', sheet.name, rangeText({ row, column }), formula]); } }),
    rename(name) { calls.push(['rename', sheet.name, name]); sheet.name = name; },
    addRows(row, count) { sheet.rowCount += count; calls.push(['addRows', sheet.name, row, count]); },
    deleteRows(row, count) { sheet.rowCount -= count; calls.push(['deleteRows', sheet.name, row, count]); },
    addColumns(column, count) { sheet.columnCount += count; calls.push(['addColumns', sheet.name, column, count]); },
    deleteColumns(column, count) { sheet.columnCount -= count; calls.push(['deleteColumns', sheet.name, column, count]); },
    setColumnsWidth(list) { calls.push(['widths', sheet.name, list]); },
    setFrozenRowCount(count) { calls.push(['freeze', sheet.name, count]); },
    charts: { addFromSelection: async options => { calls.push(['chart', sheet.name, options?.chart?.title, options?.chart?.position]); return { chartId: 'c1', chartType: 'columnClustered' }; } },
  });
  const api = {
    workbook: {
      getWorksheets: () => book.map(worksheet),
      setActiveWorksheet: async id => { for (const sheet of book) sheet.active = sheet.id === id; calls.push(['activate', book.find(sheet => sheet.id === id).name]); },
      addWorksheet(name) { book.push({ id: `s${nextId++}`, name, active: false, rowCount: 50, columnCount: 10, cells: new Map() }); calls.push(['addWorksheet', name]); },
    },
  };
  return { api, book, calls };
}

test('A1 addresses convert both ways', () => {
  assert.equal(columnIndex('A'), 0); assert.equal(columnIndex('Z'), 25); assert.equal(columnIndex('AA'), 26); assert.equal(columnIndex('ab'), 27);
  assert.equal(columnLetters(0), 'A'); assert.equal(columnLetters(27), 'AB'); assert.equal(columnLetters(701), 'ZZ');
  assert.deepEqual(parseRange('B2'), { row: 1, column: 1, rowCount: 1, columnCount: 1 });
  assert.deepEqual(parseRange('d5:b2'), { row: 1, column: 1, rowCount: 4, columnCount: 3 });
  assert.equal(parseRange('B0'), undefined); assert.equal(parseRange('B2:C3:D4'), undefined); assert.equal(parseRange(3), undefined);
  assert.equal(rangeText({ row: 1, column: 1, rowCount: 4, columnCount: 3 }), 'B2:D5');
});

test('validation turns away bad batches before anything is written', () => {
  const book = { sheets: [{ name: 'Sheet1', active: true, rowCount: 20, columnCount: 5 }] };
  const bad = (ops, pattern) => assert.throws(() => validateXlsxOps(ops, book), error => error.code === 'invalid-edit' && pattern.test(error.message));
  bad([], /at least one/);
  bad([{ op: 'chart' }], /unknown operation/);
  bad([{ op: 'setValues', range: 'A1', values: [[1, 2], [3]] }], /same number/);
  bad([{ op: 'setValues', range: 'A1:B1', values: [[1, 2, 3]] }], /1×2 but values are 1×3/);
  bad([{ op: 'setValues', range: 'E20', values: [[1, 2]] }], /past column E/);
  bad([{ op: 'setValues', range: 'A20', values: [[1], [2]] }], /past row 20/);
  bad([{ op: 'setValues', sheet: 'Nope', range: 'A1', values: [[1]] }], /no worksheet "Nope"/);
  bad([{ op: 'setFormula', cell: 'A1', formula: 'SUM(A2)' }], /start with =/);
  bad([{ op: 'setFormula', cell: 'A1:A2', formula: '=1' }], /one cell/);
  bad([{ op: 'formatRange', range: 'A1' }], /at least one/);
  bad([{ op: 'formatRange', range: 'A1', background: 'red' }], /#RRGGBB/);
  bad([{ op: 'formatRange', range: 'A1', align: 'middle' }], /align/);
  bad([{ op: 'addSheet', name: 'Sheet1' }], /already a worksheet/);
  bad([{ op: 'addSheet', name: 'a/b' }], /without/);
  bad([{ op: 'insertColumns', column: 'B:C' }], /column letter/);
  bad([{ op: 'setColumnWidth', columns: 'B', width: 5 }], /width/);
  bad([{ op: 'setColumnWidth', columns: 'E:F', width: 50 }], /past the last column/);
  bad([{ op: 'setValues', range: 'A1', values: [['=']] }], /not a formula/);
  bad([{ op: 'setValues', range: 'A1', values: [['== 备注']] }], /not a formula/);
  bad([{ op: 'addSheet', name: 'sheet1' }], /already a worksheet/);
  bad([{ op: 'deleteRows', row: 19, count: 5 }], /not all in Sheet1/);
  bad([{ op: 'deleteRows', row: 1, count: 20 }], /every row/);
  bad([{ op: 'deleteColumns', column: 'E', count: 2 }], /not all in Sheet1/);
  bad([{ op: 'freezeRows', count: 20 }], /only 20 rows/);
  // A sheet added in the batch is where unnamed operations go next, as the editor shows it.
  const { steps } = validateXlsxOps([{ op: 'addSheet', name: '新表' }, { op: 'setValues', range: 'A1', values: [[1]] }], book);
  assert.equal(steps[1].sheet, '新表');
  assert.equal(validateXlsxOps([{ op: 'renameSheet', name: 'SHEET1' }], book).steps[0].name, 'SHEET1', 'a case-only rename is allowed');
});

test('validation follows sheets added, renamed and grown in the same batch, and flags deletions for the card', () => {
  const book = { sheets: [{ name: 'Sheet1', active: true, rowCount: 3, columnCount: 3 }] };
  const { steps, danger, reasons } = validateXlsxOps([
    { op: 'insertRows', row: 4, count: 2 },
    { op: 'setValues', range: 'A5', values: [['合计', '=SUM(B1:B4)']] },
    { op: 'addSheet', name: '汇总' },
    { op: 'setValues', sheet: '汇总', range: 'A1', values: [['x']] },
    { op: 'renameSheet', sheet: 'Sheet1', name: '明细' },
    { op: 'formatRange', sheet: '明细', range: 'A1:C1', bold: true, background: '#DDEBF7', numberFormat: '#,##0' },
    { op: 'deleteRows', sheet: '明细', row: 2 },
    { op: 'clearRange', sheet: '汇总', range: 'A1' },
  ], book);
  assert.deepEqual(steps[1], { op: 'setValues', sheet: 'Sheet1', range: { row: 4, column: 0, rowCount: 1, columnCount: 2 }, values: [['合计', '=SUM(B1:B4)']] });
  assert.equal(steps[5].sheet, '明细');
  assert.equal(danger, true);
  assert.deepEqual(reasons, ['deletes 1 row', 'clears A1']);
  assert.deepEqual(describeXlsxSteps(steps).map(row => [row.op, row.where]), [['insertRows', 'Sheet1!4'], ['setValues', 'Sheet1!A5:B5'], ['addSheet', '汇总'], ['setValues', '汇总!A1'], ['renameSheet', 'Sheet1'], ['formatRange', '明细!A1:C1'], ['deleteRows', '明细!2'], ['clearRange', '汇总!A1']]);
});

test('the workbook summary lists sheets, shown values and formulas, and cuts long reads', () => {
  const summary = summarizeWorkbook({ sheets: [
    { name: '表', active: true, rowCount: 100, columnCount: 10, cells: { range: { row: 0, column: 0, rowCount: 2, columnCount: 2 }, rows: [[{ text: '项目' }, { text: '数量' }], [{ text: '合计' }, { text: '8', formula: 'SUM(B1:B1)' }]] } },
    { name: '空', rowCount: 100, columnCount: 10 },
  ] }, { maxCells: 3 });
  assert.deepEqual(summary.sheets[0], { name: '表', active: true, rows: 100, columns: 10, used: 'A1:B2', values: [['项目', '数量'], ['合计']] });
  assert.deepEqual(summary.sheets[1], { name: '空', rows: 100, columns: 10, used: null });
  assert.equal(summary.truncated, true);
  assert.equal(xlsxProgress(0, 2), 'None of the 2 steps was applied.');
  assert.match(xlsxProgress(1, 2), /Steps 1–1 of 2 were applied and stay in the editor/);
});

test('steps write values and formulas, keep formulas when formatting, add and switch sheets, and select what they wrote', async () => {
  const fake = fakeSheet();
  const book = inspectWorkbook(fake.api);
  const { steps } = validateXlsxOps([
    { op: 'setValues', range: 'A1', values: [['项目', '数量'], ['苹果', 3], ['香蕉', 5], ['合计', '=SUM(B2:B3)']] },
    { op: 'formatRange', range: 'A1:B4', bold: true, numberFormat: '0' },
    { op: 'setColumnWidth', columns: 'A:B', width: 120 },
    { op: 'freezeRows', count: 1 },
    { op: 'addSheet', name: '汇总' },
    { op: 'setFormula', sheet: '汇总', cell: 'A1', formula: '=Sheet1!B4' },
  ], book);
  await runXlsxSteps(fake.api, steps);
  const cells = fake.book[0].cells;
  assert.deepEqual(cells.get('3,1'), { formula: '=SUM(B2:B3)', bold: true, format: '0' }, 'the formula stays, the format is added');
  assert.deepEqual(cells.get('1,1'), { value: 3, bold: true, format: '0' }, 'setData takes the raw value, not getData\'s { type, value }');
  assert.deepEqual(fake.calls.filter(call => call[0] === 'select').map(call => call.slice(1)), [['Sheet1', 'A1:B4'], ['Sheet1', 'A1:B4'], ['汇总', 'A1']]);
  assert.deepEqual(fake.calls.find(call => call[0] === 'widths')[2], [{ column: 0, width: 120 }, { column: 1, width: 120 }]);
  assert.ok(fake.calls.some(call => call[0] === 'activate' && call[1] === '汇总'), 'the new sheet is shown before it is written');
  assert.deepEqual(fake.book[1].cells.get('0,0'), { formula: '=Sheet1!B4' });
  const read = inspectWorkbook(fake.api);
  assert.deepEqual(structuredClone(read).sheets[0].cells.range, { row: 0, column: 0, rowCount: 4, columnCount: 2 });
  // Formatting again keeps the format a cell already has.
  const { steps: again } = validateXlsxOps([{ op: 'formatRange', sheet: 'Sheet1', range: 'B2', background: '#FFF2CC' }], read);
  await runXlsxSteps(fake.api, again);
  assert.deepEqual(cells.get('1,1'), { value: 3, bold: true, format: '0', background: '#FFF2CC' });
  // A big sheet is read only up to the cap, whole rows from the top; the summary still names the whole used range.
  const capped = inspectWorkbook(fake.api, { maxCells: 4 });
  assert.deepEqual([capped.sheets[0].cells.range.rowCount, capped.sheets[0].cells.used.rowCount], [2, 4]);
  const summary = summarizeWorkbook(capped);
  assert.deepEqual([summary.sheets[0].used, summary.truncated], ['A1:B4', true]);
});

/** A page with no Sheet DOM: the overlay has nothing to measure but the agent runs. */
function fakePage() {
  const element = () => ({ style: {}, dataset: {}, append() {}, remove() {}, setAttribute() {}, getBoundingClientRect: () => ({ width: 0 }) });
  return { defaultView: { requestAnimationFrame: () => 1, cancelAnimationFrame() {}, performance: { now: () => 0 }, addEventListener() {}, removeEventListener() {} },
    body: { append() {} }, createElement: element, createElementNS: element, addEventListener() {}, removeEventListener() {}, querySelectorAll: () => [], querySelector: () => null };
}

test('the agent applies a batch and is the AI until prepare has run', async () => {
  const fake = fakeSheet({ bounding: false });
  const order = [];
  const agent = createKitSheetAgent({ editor: { api: { sheet: fake.api }, prepare: async () => order.push('prepare') }, document: fakePage(), animate: () => false, wait: async () => {} });
  const pending = agent.apply({ ops: [{ op: 'setValues', range: 'A1', values: [[1]] }] });
  assert.equal(agent.applying, true);
  assert.deepEqual(await pending, { applied: 1 });
  assert.deepEqual(order, ['prepare']);
  assert.equal(agent.applying, false);
  await assert.rejects(agent.undo({}), error => error.code === 'unsupported-operation');
});

test('a failing or stopped batch keeps finished steps and says how far it got', async () => {
  const failing = fakeSheet({ failOn: 'setValue', bounding: false });
  const agent = createKitSheetAgent({ editor: { api: { sheet: failing.api } }, document: fakePage(), animate: () => false, wait: async () => {} });
  await assert.rejects(agent.apply({ ops: [{ op: 'addSheet', name: '新' }, { op: 'setValues', range: 'A1', values: [[1]] }] }),
    error => error.code === 'editor-failed' && /refused Steps 1–1 of 2 were applied/.test(error.message));
  assert.equal(failing.book.length, 2, 'the sheet added first stays');
  const fake = fakeSheet({ bounding: false });
  const abort = new AbortController();
  let pauses = 0;
  // Pauses: step 1's selection, the pause after it, step 2's selection (a stop there writes nothing more).
  const stopping = createKitSheetAgent({ editor: { api: { sheet: fake.api } }, document: fakePage(), animate: () => true, wait: async () => { if (++pauses === 3) abort.abort(); } });
  await assert.rejects(stopping.apply({ ops: [{ op: 'setValues', range: 'A1', values: [[1]] }, { op: 'setValues', range: 'A2', values: [[2]] }] }, abort.signal),
    error => error.code === 'cancelled' && /Stopped by the user\. Steps 1–1 of 2 were applied/.test(error.message));
  assert.deepEqual([...fake.book[0].cells.keys()], ['0,0']);
});

test('v1.1: the range is selected before it is written, rows go in one by one, and a stop mid-table says so', async () => {
  const fake = fakeSheet();
  const log = [];
  const { steps } = validateXlsxOps([{ op: 'setValues', range: 'A1', values: [['a', 1], ['b', 2], ['c', '=B1+B2']] }, { op: 'addSheet', name: '新' }], inspectWorkbook(fake.api));
  const original = fake.calls.push.bind(fake.calls);
  fake.calls.push = (...items) => { for (const item of items) log.push(item.slice(0, 3).join(' ')); return original(...items); };
  await runXlsxSteps(fake.api, steps, { rowByRow: true, onSelect: async (number, range) => log.push(`select ${number} ${range ? `${range.row},${range.rowCount}` : '-'}`), onFinish: number => log.push(`done ${number}`) });
  assert.deepEqual(log.filter(line => !line.startsWith('select Sheet1') && !line.startsWith('activate')), [
    'select 0 0,3', 'setValue Sheet1 A1:B1', 'setValue Sheet1 A2:B2', 'setValue Sheet1 A3:B3', 'formula Sheet1 B3', 'done 0',
    'select 1 -', 'addWorksheet 新', 'done 1',
  ]);
  const stopping = fakeSheet({ bounding: false });
  const abort = new AbortController();
  let pauses = 0;
  const agent = createKitSheetAgent({ editor: { api: { sheet: stopping.api } }, document: fakePage(), animate: () => true, wait: async () => { if (++pauses === 3) abort.abort(); } });
  await assert.rejects(agent.apply({ ops: [{ op: 'setValues', range: 'A1', values: [[1], [2], [3], [4]] }] }, abort.signal),
    error => error.code === 'cancelled' && /None of the 1 steps was applied\. Step 1 was partly applied/.test(error.message));
  assert.ok(stopping.book[0].cells.size >= 1 && stopping.book[0].cells.size < 4, 'rows already written stay');
});

test('v1.1: a stop during the selection writes nothing; a stop after the first row says the step is partly written', async () => {
  const during = fakeSheet({ bounding: false });
  const abort = new AbortController();
  const agent = createKitSheetAgent({ editor: { api: { sheet: during.api } }, document: fakePage(), animate: () => true, wait: async () => abort.abort() });
  await assert.rejects(agent.apply({ ops: [{ op: 'setValues', range: 'A1', values: [[1], [2]] }] }, abort.signal), error => error.code === 'cancelled' && /None of the 1 steps was applied\.$/.test(error.message));
  assert.equal(during.book[0].cells.size, 0);
  const after = fakeSheet({ bounding: false });
  const second = new AbortController();
  let pauses = 0;
  const agent2 = createKitSheetAgent({ editor: { api: { sheet: after.api } }, document: fakePage(), animate: () => true, wait: async () => { if (++pauses === 2) second.abort(); } });
  await assert.rejects(agent2.apply({ ops: [{ op: 'setValues', range: 'A1', values: [[1], [2]] }] }, second.signal), error => /Step 1 was partly applied/.test(error.message));
  assert.equal(after.book[0].cells.size, 1, 'row 1 stays');
});

test('addChart makes a column chart of the selected range, after the selection gesture', async () => {
  const fake = fakeSheet();
  const book = inspectWorkbook(fake.api);
  assert.throws(() => validateXlsxOps([{ op: 'addChart', range: 'A1:A5' }], book), /header row and a label column/);
  const { steps } = validateXlsxOps([{ op: 'addChart', range: 'A1:D9', title: '季度销售' }], book);
  assert.deepEqual(describeXlsxSteps(steps)[0], { op: 'addChart', where: 'Sheet1!A1:D9', after: '季度销售' });
  const log = [];
  await runXlsxSteps(fake.api, steps, { onSelect: async () => log.push('select') });
  assert.deepEqual(fake.calls.filter(call => call[0] === 'select' || call[0] === 'chart').map(call => call[0]), ['select', 'chart', 'select']);
  // Below the data, one blank row between, from the range's first column: the chart never covers what it plots.
  assert.deepEqual(fake.calls.find(call => call[0] === 'chart'), ['chart', 'Sheet1', '季度销售', { row: 10, column: 0, offsetX: 0, offsetY: 0 }]);
  assert.deepEqual(log, ['select']);
});

test('a chart goes one blank row under its range, from the range\'s first column, wherever the range starts', async () => {
  assert.deepEqual(chartPosition(parseRange('A1:D5')), { row: 6, column: 0, offsetX: 0, offsetY: 0 }, 'A1:D5 → row 7 (0.0.39-test.11 put it over B3:D5)');
  assert.deepEqual(chartPosition(parseRange('C4:F10')), { row: 11, column: 2, offsetX: 0, offsetY: 0 });
  const fake = fakeSheet();
  const { steps } = validateXlsxOps([{ op: 'addChart', range: 'B2:E6' }], inspectWorkbook(fake.api));
  await runXlsxSteps(fake.api, steps);
  const [, , title, position] = fake.calls.find(call => call[0] === 'chart');
  assert.equal(title, undefined, 'no title given, none sent');
  assert.deepEqual(position, { row: 7, column: 1, offsetX: 0, offsetY: 0 });
});

test('261109: a small range fills cell by cell, each cell selected first; a larger one row by row; formulas stay formulas', async () => {
  const fake = fakeSheet();
  const log = [];
  const { steps } = validateXlsxOps([{ op: 'setValues', range: 'A1', values: [['a', 1], ['b', '=B1*2']] }], inspectWorkbook(fake.api));
  const original = fake.calls.push.bind(fake.calls);
  fake.calls.push = (...items) => { for (const item of items) log.push(item.slice(0, 3).join(' ')); return original(...items); };
  const pauses = [];
  await runXlsxSteps(fake.api, steps, { rowByRow: true, cellByCell: 48, cellPause: 40, wait: async ms => pauses.push(ms) });
  assert.deepEqual(log.filter(line => !line.startsWith('activate')), [
    'select Sheet1 A1', 'setValue Sheet1 A1', 'select Sheet1 B1', 'setValue Sheet1 B1', 'select Sheet1 A2', 'setValue Sheet1 A2', 'select Sheet1 B2', 'formula Sheet1 B2',
    'select Sheet1 A1:B2',
  ], 'the whole range is selected again when the step is done');
  assert.deepEqual(pauses, [40, 40, 40], 'a pause between cells, none after the last');
  assert.equal(fake.book[0].cells.get('1,1').formula, '=B1*2');
  assert.deepEqual(['0,0', '0,1', '1,0'].map(at => fake.book[0].cells.get(at).value), [{ type: 'primitive', value: 'a' }, { type: 'primitive', value: 1 }, { type: 'primitive', value: 'b' }], 'the planned values are in the cells');
  const big = fakeSheet();
  const values = Array.from({ length: 7 }, (_, r) => Array.from({ length: 7 }, (_, c) => r * 7 + c));
  const { steps: bigSteps } = validateXlsxOps([{ op: 'setValues', range: 'A1', values }], inspectWorkbook(big.api));
  await runXlsxSteps(big.api, bigSteps, { rowByRow: true, cellByCell: 48 });
  assert.equal(big.calls.filter(call => call[0] === 'setValue').length, 7, '49 cells go in a row at a time');
  assert.equal(big.book[0].cells.size, 49);
});

test('261109: the agent fills cell by cell only with motion; a stop between cells says the step is partly written', async () => {
  const still = fakeSheet({ bounding: false });
  await createKitSheetAgent({ editor: { api: { sheet: still.api } }, document: fakePage(), animate: () => false, wait: async () => {} }).apply({ ops: [{ op: 'setValues', range: 'A1', values: [[1, 2], [3, 4]] }] });
  assert.deepEqual(still.calls.filter(call => call[0] === 'setValue').map(call => call[2]), ['A1:B2'], 'reduced motion writes the range at once');
  const moving = fakeSheet({ bounding: false });
  const abort = new AbortController();
  let pauses = 0;
  const agent = createKitSheetAgent({ editor: { api: { sheet: moving.api } }, document: fakePage(), animate: () => true, wait: async () => { if (++pauses === 3) abort.abort(); } });
  await assert.rejects(agent.apply({ ops: [{ op: 'setValues', range: 'A1', values: [[1, 2], [3, 4]] }] }, abort.signal),
    error => error.code === 'cancelled' && /Step 1 was partly applied \(part of it is written\)/.test(error.message));
  assert.ok(moving.book[0].cells.size >= 1 && moving.book[0].cells.size < 4, 'cells already written stay');
});
