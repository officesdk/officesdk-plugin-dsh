import { SELECTION_LIMITS, cpLength, planDocRewrite, planSheetRevert, planSheetRewrite, sheetAddress } from '../agent/selection.mjs';
import { docxReadToken } from '../agent/docx.mjs';
import { agentLabel } from './agent-status.mjs';
import { inspectDocument, paragraphLocator } from './document-agent.mjs';
import { textChunks } from './selection.mjs';
import { inspectWorkbook, keptStyle } from './sheet-agent.mjs';
import { validateXlsxOps } from '../agent/xlsx.mjs';

/**
 * Selection rewrite on a Kit page (spec 261109-selection-ai-edit §4.2, §5): reading what the user selected, replacing it
 * with the model's answer while the user watches, and taking that back. Plans come from agent/selection.mjs; the
 * agents (document-agent.mjs, sheet-agent.mjs) own the execution lock and the overlay.
 */
const agentError = (code, message) => Object.assign(new Error(message), { code });
/** Plain data only: postMessage cannot clone the SDK's objects or the editor's proxies. */
const plain = value => value === undefined ? undefined : JSON.parse(JSON.stringify(value));

/**
 * What the user selected in Writer, as a selection request's `{ kind, anchor, original }` plus `view` (the selected
 * paragraphs' whole texts, for drawing around them); `{ unsupported: reason }` for a selection that cannot be rewritten
 * ('table', 'too-long'); undefined for no selection.
 */
export async function readDocSelection(Word) {
  let result;
  await Word.run(async ctx => {
    const selection = ctx.document.getSelection();
    selection.load('text,isEmpty');
    const body = ctx.document.body.paragraphs;
    body.load('items/text,items/tableNestingLevel');
    const own = selection.paragraphs;
    own.load('items/text');
    await ctx.sync();
    if (selection.isEmpty || !String(selection.text ?? '').replace(/[\r\n]/g, '') || !own.items.length) return;
    const texts = Array.from(body.items, paragraph => String(paragraph.text ?? ''));
    const first = own.items[0], last = own.items[own.items.length - 1];
    // Where the selection's first and last paragraphs are among the body's: by text, and when several paragraphs read
    // the same, by location.
    const locate = async paragraph => {
      const candidates = texts.flatMap((text, index) => text === String(paragraph.text ?? '') ? [index] : []);
      if (candidates.length <= 1) return candidates[0] ?? -1;
      const results = candidates.map(index => body.items[index].getRange('Whole').compareLocationWith(paragraph.getRange('Whole')));
      await ctx.sync();
      return candidates[results.findIndex(result => result.value === 'Equal')] ?? -1;
    };
    let from = await locate(first), to = own.items.length === 1 ? from : await locate(last);
    if (from < 0 || to < 0 || to < from) return;
    const before = first.getRange('Start').expandTo(selection.getRange('Start'));
    const after = selection.getRange('End').expandTo(last.getRange('End'));
    before.load('text');
    after.load('text');
    await ctx.sync();
    const start = cpLength(before.text);
    let end = Math.max(0, cpLength(texts[to]) - cpLength(after.text));
    // A selection that runs into the next paragraph's start (a triple click takes the paragraph mark) ends where the
    // paragraph before it ends.
    if (to > from && end === 0) { to -= 1; end = cpLength(texts[to]); }
    if (body.items.slice(from, to + 1).some(paragraph => Number(paragraph.tableNestingLevel ?? 0) > 0)) { result = { unsupported: 'table' }; return; }
    const parts = from === to ? [Array.from(texts[from]).slice(start, end).join('')]
      : [Array.from(texts[from]).slice(start).join(''), ...texts.slice(from + 1, to), Array.from(texts[to]).slice(0, end).join('')];
    const original = parts.join('\n');
    if (!original.replace(/\n/g, '').trim()) return;
    if (to - from + 1 > SELECTION_LIMITS.maxParagraphs || cpLength(original) > SELECTION_LIMITS.maxText) { result = { unsupported: 'too-long' }; return; }
    result = { kind: 'docx', anchor: { paragraphs: [from, to], start, end }, original, view: { first: from, texts: texts.slice(from, to + 1) } };
  });
  return result;
}

/** Selects (best effort) the text under `after` so the user sees what changed; `texts` are the paragraphs now. */
export async function selectDocRange(Word, { paragraphs: [first, last], start, end }, texts) {
  try {
    await Word.run(async ctx => {
      const body = ctx.document.body.paragraphs;
      body.load('items');
      await ctx.sync();
      const head = body.items[first].getRange('Whole');
      if (first === last) { if (end > start) head.getSubrange(start, end - start).select(); }
      else {
        const from = cpLength(texts[first]) > start ? head.getSubrange(start, cpLength(texts[first]) - start) : head.getRange('End');
        const tail = body.items[last].getRange('Whole');
        from.expandTo(end > 0 ? tail.getSubrange(0, end) : tail.getRange('Start')).select();
      }
      await ctx.sync();
    });
  } catch {}
}

/**
 * Replaces code points [start, start + length) of paragraph `index` with `text`, in the parts `chunk` gives (each a
 * prefix of `text`, the last the whole): every write replaces what the previous one wrote, so the paragraph's other
 * text and formatting never move. `onWrite(text)` after each write.
 */
export async function replaceInParagraph(Word, { index, start, length, text, signal, chunk = value => [value], onWrite = () => {} }) {
  let current = length;
  const parts = text ? chunk(text) : [''];
  for (const part of parts) {
    if (signal?.aborted) throw agentError('cancelled', 'Stopped by the user.');
    await Word.run(async ctx => {
      const body = ctx.document.body.paragraphs;
      body.load('items');
      await ctx.sync();
      const whole = body.items[index].getRange('Whole');
      if (current > 0) {
        const range = whole.getSubrange(start, current);
        if (part) range.insertText(part, 'Replace'); else range.delete();
      } else if (part) {
        // Nothing there to replace (a rewrite to empty text being taken back): insert at the position.
        if (start > 0) whole.getSubrange(start - 1, 1).insertText(part, 'After');
        else whole.getRange('Start').insertText(part, 'Before');
      }
      await ctx.sync();
    });
    current = cpLength(part);
    onWrite(part);
  }
}

/**
 * Rewrites (or takes back) a Word selection on the agent's terms: `agent` gives `Word()`, the overlay, `apply` (a
 * document_edit batch with its animation) and the labels. `plan` is planDocRewrite's result.
 */
export async function runDocPlan({ Word, document, overlay, locale, apply, animate }, plan, doc, signal) {
  if (plan.mode === 'unchanged') return;
  if (plan.mode === 'paragraphs') {
    await apply({ ops: plan.ops, readToken: docxReadToken(doc) }, signal);
    return;
  }
  const text = doc.paragraphs[plan.paragraph]?.text ?? '';
  await overlay.select('t0', paragraphLocator(document, text, plan.paragraph), { label: agentLabel(locale, 'edit'), selecting: agentLabel(locale, 'selecting'), edge: false });
  try {
    if (signal?.aborted) throw agentError('cancelled', 'Stopped by the user.');
    const head = Array.from(text).slice(0, plan.start).join(''), tail = Array.from(text).slice(plan.start + plan.length).join('');
    await replaceInParagraph(Word, { index: plan.paragraph, start: plan.start, length: plan.length, text: plan.text, signal, chunk: animate ? textChunks : value => [value],
      onWrite: part => overlay.point('t0', paragraphLocator(document, `${head}${part}${tail}`, plan.paragraph)) });
  } finally { overlay.release('t0'); }
}

/** A spreadsheet cell as a selection records it: shown text, formula (no "=") and raw value. */
export function cellOf(data) {
  const value = data?.value;
  const raw = value && typeof value === 'object' && (value.type === 'primitive' || value.type === 'date') ? value.value : null;
  const text = String(data?.text ?? (raw === null ? (value && typeof value === 'object' && value.value && typeof value.value === 'object' ? value.value.error ?? '' : '') : raw));
  return { text, raw, ...(data?.formula && { formula: String(data.formula).replace(/^=/, '') }), ...(value?.type === 'date' && { date: true }) };
}

/**
 * What the user selected in the Sheet: `{ kind, anchor, cells, original, box }`, `{ unsupported: reason }` ('multiple',
 * 'not-cells', 'too-many', 'merged') or undefined.
 */
export function readSheetSelection(sheetApi) {
  const worksheet = sheetApi.workbook.getActiveWorksheet();
  const selections = worksheet?.getSelections?.() ?? sheetApi.selections;
  if (!selections?.length) return undefined;
  if (selections.length > 1) return { unsupported: 'multiple' };
  const range = selections[0].getRange();
  if (!range) return undefined;
  if (range.type !== 'cells') return { unsupported: 'not-cells' };
  if (range.rowCount * range.columnCount > SELECTION_LIMITS.maxCells) return { unsupported: 'too-many' };
  const data = plain(range.getData()) ?? [];
  if (data.some(row => (row ?? []).some(cell => cell && Object.keys(cell).some(key => /merge/i.test(key) && cell[key])))) return { unsupported: 'merged' };
  const anchor = { sheet: String(worksheet.name), row: range.row, column: range.column, rowCount: range.rowCount, columnCount: range.columnCount };
  const cells = Array.from({ length: anchor.rowCount }, (_, r) => Array.from({ length: anchor.columnCount }, (_, c) => cellOf(data[r]?.[c])));
  let box;
  try { box = range.getBounding?.() ?? undefined; } catch {}
  return { kind: 'xlsx', anchor, cells, ...(box && { box: plain(box) }) };
}

/** The cells under `anchor` now, as a selection records them. */
export function readSheetCells(sheetApi, anchor) {
  const worksheet = Array.from(sheetApi.workbook.getWorksheets() ?? []).find(sheet => sheet.name === anchor.sheet);
  if (!worksheet) throw agentError('source-changed', `Worksheet ${anchor.sheet} is no longer in the workbook.`);
  const range = worksheet.getRange({ type: 'cells', row: anchor.row, column: anchor.column, rowCount: anchor.rowCount, columnCount: anchor.columnCount });
  const data = plain(range.getData()) ?? [];
  return { cells: Array.from({ length: anchor.rowCount }, (_, r) => Array.from({ length: anchor.columnCount }, (_, c) => cellOf(data[r]?.[c]))), range, worksheet };
}

/** The spreadsheet_edit operation for a planned setValues step. */
export const sheetOp = step => ({ op: 'setValues', sheet: step.sheet, range: sheetAddress(step.range), values: step.values });

/** The screen box around paragraphs `view.texts` from `view.first` on (those drawn), or undefined. */
export function locateParagraphs(document, view) {
  let box;
  for (const [offset, text] of (view?.texts ?? []).entries()) {
    const found = paragraphLocator(document, text, view.first + offset)();
    if (!found) continue;
    box = box ? { left: Math.min(box.left, found.left), top: Math.min(box.top, found.top), right: Math.max(box.right, found.left + found.width), bottom: Math.max(box.bottom, found.top + found.height) }
      : { left: found.left, top: found.top, right: found.left + found.width, bottom: found.top + found.height };
  }
  return box && { left: box.left, top: box.top, width: box.right - box.left, height: box.bottom - box.top };
}

/**
 * The selection methods of the Word agent. `env`: `Word()`, the page document, the overlay, locale, `animate()`,
 * `apply` (the agent's document_edit apply, with its own lock) and `locked` (the agent's lock for writes outside apply).
 */
export function docSelectionMethods(env) {
  const run = async ({ anchor, original, text }, signal) => {
    const word = env.Word();
    const doc = await inspectDocument(word);
    const plan = planDocRewrite(doc.paragraphs, anchor, original, text);
    if (plan.mode === 'inline') await env.locked(() => runDocPlan({ ...env, Word: word, animate: env.animate() }, plan, doc, signal));
    else if (plan.mode === 'paragraphs') await env.apply({ ops: plan.ops, readToken: docxReadToken(doc) }, signal);
    const now = await inspectDocument(word);
    const texts = now.paragraphs.map(paragraph => paragraph.text);
    const [first, last] = plan.after.paragraphs;
    await selectDocRange(word, plan.after, texts);
    return { plan, now, view: { first, texts: texts.slice(first, last + 1) } };
  };
  return {
    readSelection: () => env.applying() ? Promise.resolve(undefined) : readDocSelection(env.Word()),
    locate: target => locateParagraphs(env.document, target?.view),
    async rewriteSelection(args, signal) {
      const { plan, now, view } = await run(args, signal);
      return {
        reply: { applied: plan.mode !== 'unchanged', mode: plan.mode, before: args.original, after: plan.afterText, readToken: docxReadToken(now) },
        revert: { anchor: plan.after, original: plan.afterText, text: args.original },
        located: { view },
      };
    },
    async revertSelection(revert, signal) { await run(revert, signal); },
  };
}

/**
 * Cells under `anchor` that hold a formula and get a plain value from `values`: writing a value does not remove a
 * formula in the Sheet SDK (measured: `setValue` with null or '' keeps it), so they are cleared first with `setData`,
 * keeping their format. Returns the cleared cells with their formulas, for restoreFormulas.
 */
export function clearFormulas(sheetApi, anchor, values) {
  const { range, worksheet } = readSheetCells(sheetApi, anchor);
  const data = plain(range.getData()) ?? [];
  const cleared = [];
  for (const [r, row] of values.entries()) for (const [c, value] of row.entries()) {
    const cell = data[r]?.[c];
    if (!cell?.formula || (typeof value === 'string' && value.startsWith('='))) continue;
    worksheet.getRange({ type: 'cells', row: anchor.row + r, column: anchor.column + c, rowCount: 1, columnCount: 1 }).setData([[{ ...keptStyle(cell), formula: null, value: null }]]);
    cleared.push({ row: anchor.row + r, column: anchor.column + c, formula: String(cell.formula).replace(/^=/, '') });
  }
  return cleared;
}

/** A write that failed after clearFormulas: the formulas go back into the cells still empty. */
export function restoreFormulas(sheetApi, anchor, cleared) {
  const { worksheet } = readSheetCells(sheetApi, anchor);
  for (const { row, column, formula } of cleared) {
    const now = cellOf(plain(worksheet.getRange({ type: 'cells', row, column, rowCount: 1, columnCount: 1 }).getData())?.[0]?.[0]);
    if (!now.formula && (now.raw === null || now.raw === '')) worksheet.getCell(row, column).setCellFormula(`=${formula}`);
  }
}

/** Cells written with their own value type (dates, text starting with "="), which setValues would change. */
export function writeSpecials(sheetApi, anchor, specials) {
  const { worksheet } = readSheetCells(sheetApi, anchor);
  for (const { r, c, value } of specials) worksheet.getRange({ type: 'cells', row: anchor.row + r, column: anchor.column + c, rowCount: 1, columnCount: 1 }).setValue([[value]]);
}

/** The selection methods of the Sheet agent (`env`: `api()`, `apply`, `locked`, `applying()`). */
export function sheetSelectionMethods(env) {
  const write = async (anchor, step, signal, specials = []) => {
    const op = sheetOp(step);
    // Refused values must not cost the formulas: check the batch before anything is cleared.
    validateXlsxOps([op], inspectWorkbook(env.api()));
    const cleared = await env.locked(async () => clearFormulas(env.api(), anchor, step.values));
    try { await env.apply({ ops: [op] }, signal); }
    catch (error) {
      if (cleared.length) await env.locked(async () => restoreFormulas(env.api(), anchor, cleared)).catch(() => {});
      throw error;
    }
    if (specials.length) await env.locked(async () => writeSpecials(env.api(), anchor, specials));
  };
  const locate = target => {
    if (target?.anchor) {
      try {
        const worksheet = Array.from(env.api().workbook.getWorksheets() ?? []).find(sheet => sheet.name === target.anchor.sheet);
        const { row, column, rowCount, columnCount } = target.anchor;
        const box = worksheet?.isActive ? worksheet.getRange({ type: 'cells', row, column, rowCount, columnCount }).getBounding() : undefined;
        if (box && box.width > 0) return plain(box);
      } catch {}
    }
    return target?.box;
  };
  return {
    readSelection: async () => env.applying() ? undefined : readSheetSelection(env.api()),
    locate,
    /** The Sheet keeps the keyboard for its grid and takes focus back; the request box asks it to let go while open. */
    async focusEditor(on) { try { await env.api().setFocus?.(on); } catch {} },
    async rewriteSelection({ anchor, cells, values }, signal) {
      const { cells: now } = readSheetCells(env.api(), anchor);
      const { step } = planSheetRewrite(now, anchor, cells, values);
      await write(anchor, step, signal);
      return {
        reply: { applied: true, range: `${anchor.sheet}!${sheetAddress(anchor)}`, before: cells.map(row => row.map(cell => cell.formula ? `=${cell.formula}` : cell.text)), after: step.values },
        revert: { anchor, written: step.values, cells },
        located: { anchor },
      };
    },
    async revertSelection({ anchor, written, cells }, signal) {
      const { cells: now } = readSheetCells(env.api(), anchor);
      const { step, specials } = planSheetRevert(now, anchor, written, cells);
      await write(anchor, step, signal, specials);
    },
  };
}
