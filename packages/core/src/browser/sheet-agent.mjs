import { XLSX_LIMITS, columnLetters, parseRange, validateXlsxOps, xlsxProgress } from '../agent/xlsx.mjs';
import { creationBands, createGenerationSurface } from './creation-surface.mjs';
import { createDocOverlay } from './doc-overlay.mjs';
import { agentLabel, xlsxStatus } from './agent-status.mjs';
import { sheetSelectionMethods } from './selection-edit.mjs';

/**
 * AI editing of a spreadsheet on a Kit page (spec 261105 §6), through `editor.api.sheet`, writable in edit mode in
 * office-kit's editor page (office-kit request K16, public since 0.3.0-rc.1). Each operation is one or a
 * few SDK calls; the range being written is selected (the Sheet's own selection is the highlight) and outlined with the
 * agent's frame. There is no rollback: a failed or stopped batch keeps the steps it finished and says how far it got.
 */
const agentError = (code, message) => Object.assign(new Error(message), { code });
const PAGE_CODES = new Set(['invalid-edit', 'cancelled', 'busy', 'read-only', 'unsupported-operation']);
/** Plain data only: postMessage cannot clone the SDK's objects or the editor's proxies. */
/** With motion, a range of at most this many cells is filled one cell at a time (larger ones a row at a time). */
export const XLSX_ANIM_CELLS = 48;
const plain = value => value === undefined ? undefined : JSON.parse(JSON.stringify(value));

export const SHEET_LABELS = {
  zh: { editing: 'AI 编辑中', step: (n, total) => `AI 正在改表 第 ${n}/${total} 步`, marks: { add: 'Agent · 写入中', edit: 'Agent · 修改中', delete: 'Agent · 删除中' } },
  en: { editing: 'AI editing', step: (n, total) => `AI is editing step ${n} of ${total}`, marks: { add: 'Agent · Writing', edit: 'Agent · Editing', delete: 'Agent · Deleting' } },
};

/**
 * Workbook as the page reads it: worksheets with size and the used range's cells (shown text and formula). Big sheets
 * are read only up to `maxCells` per sheet (whole rows from the top): every read crosses postMessage.
 */
export function inspectWorkbook(sheetApi, { maxCells = XLSX_LIMITS.maxReadCells } = {}) {
  const sheets = Array.from(sheetApi.workbook.getWorksheets() ?? [], worksheet => {
    const used = worksheet.getUsedRange?.();
    let cells;
    if (used && used.rowCount > 0 && used.columnCount > 0) {
      const usedRange = { row: used.row, column: used.column, rowCount: used.rowCount, columnCount: used.columnCount };
      const rows = Math.max(1, Math.min(used.rowCount, Math.floor(maxCells / used.columnCount)));
      const read = rows < used.rowCount ? worksheet.getRange({ type: 'cells', ...usedRange, rowCount: rows }) : used;
      const data = plain(read.getData()) ?? [];
      cells = { range: { ...usedRange, rowCount: rows }, ...(rows < used.rowCount && { used: usedRange }),
        rows: data.map(row => (row ?? []).map(cell => ({ text: String(cell?.text ?? (cell?.value && typeof cell.value === 'object' ? cell.value.value ?? '' : '')), ...(cell?.formula && { formula: String(cell.formula) }) }))) };
    }
    return { id: String(worksheet.id), name: String(worksheet.name), active: worksheet.isActive === true, rowCount: Number(worksheet.rowCount), columnCount: Number(worksheet.columnCount), ...(cells && { cells }) };
  });
  return { sheets };
}

const primitive = value => value === null ? null : { type: 'primitive', value };
/** The Sheet's cell-data keys for a formatRange step (fontSize and fontFamily are mapped by the SDK). */
function dataOf(format) {
  const data = {};
  for (const key of ['bold', 'italic', 'underline', 'color', 'background', 'fontSize', 'align', 'wrap']) if (format[key] !== undefined) data[key] = format[key];
  if (format.numberFormat !== undefined) data.format = format.numberFormat;
  return data;
}

/** Style keys getData reports and setData takes back (fontFamily/fontSize and borders are mapped by the SDK itself). */
const STYLE_KEYS = ['bold', 'italic', 'underline', 'strike', 'color', 'background', 'fontFamily', 'fontSize', 'align', 'vertical', 'wrap', 'format', 'borderTop', 'borderRight', 'borderBottom', 'borderLeft'];
export function keptStyle(cell) {
  const kept = {};
  for (const key of STYLE_KEYS) if (cell?.[key] !== undefined && cell[key] !== null) kept[key] = cell[key];
  return kept;
}

// Divide one validated values matrix into disjoint blocks. SDK mutations stay ordered while
// the cursors and their writes progress together; subsequent formatting and formulas keep their order.
async function writeGenerationValues(sheet, step, {signal, select, finish, wait, track, number}) {
  const {row, column, rowCount, columnCount} = step.range;
  const vertical = rowCount > 1;
  const length = vertical ? rowCount : columnCount;
  const count = Math.min(3, length);
  const bands = Array.from({length:count},(_,lane)=>{
    const lo = Math.floor(lane * length / count), hi = Math.floor((lane + 1) * length / count);
    return {row:vertical?row+lo:row,column:vertical?column:column+lo,rowCount:vertical?hi-lo:1,columnCount:vertical?columnCount:hi-lo};
  });
  const small = rowCount * columnCount <= XLSX_ANIM_CELLS;
  const ticks = small ? Math.max(...bands.map(b=>b.rowCount*b.columnCount)) : Math.min(16,Math.max(...bands.map(b=>b.rowCount)));
  try {
    await Promise.all(bands.map((b,lane)=>select(lane,sheet.getRange({type:'cells',...b}))));
    for (let tick=0;tick<ticks;tick++) {
      signal?.throwIfAborted();
      for (const b of bands) {
        const size = small ? b.rowCount*b.columnCount : b.rowCount;
        if (small && tick>=size) continue;
        const from = small ? Math.floor(tick/b.columnCount) : Math.floor(tick*size/ticks);
        const end = small ? from+1 : Math.floor((tick+1)*size/ticks);
        if (end===from) continue;
        const x = small ? b.column+tick%b.columnCount : b.column;
        const width = small ? 1 : b.columnCount;
        const y = b.row+from;
        const values=step.values.slice(y-row,y-row+end-from).map(line=>line.slice(x-column,x-column+width));
        const range=sheet.getRange({type:'cells',row:y,column:x,rowCount:values.length,columnCount:width});
        range.setValue(values.map(line=>line.map(v=>typeof v==='string'&&v.startsWith('=')?null:primitive(v))));
        track.partial=number;
        values.forEach((line,r)=>line.forEach((v,c)=>{
          if (typeof v==='string' && v.startsWith('=')) sheet.getCell(y+r,x+c).setCellFormula(v);
        }));
      }
      if (tick+1<ticks) await wait(65);
    }
    track.partial=undefined;
  } finally { bands.forEach((_,i)=>finish(i)); }
}

/**
 * Runs validated steps one at a time; `onStep(number, range?, worksheet?, tone)` draws where it acts. With `rowByRow`,
 * values go in a row at a time, or a cell at a time when the range has at most `cellByCell` cells (spec 261109 §4B).
 */
export async function runXlsxSteps(sheetApi, steps, { signal, onStep = () => {}, onSelect = async () => {}, onWrite = () => {}, onFinish = () => {}, local = false, creationTracks, rowByRow = false, cellByCell = 0, cellPause = 0, wait = () => Promise.resolve(), pause = 0, track = {} } = {}) {
  const book = () => sheetApi.workbook;
  const worksheet = name => {
    const found = Array.from(book().getWorksheets() ?? []).find(sheet => sheet.name === name);
    if (!found) throw agentError('editor-failed', `Worksheet ${name} is no longer in the workbook.`);
    return found;
  };
  /** The user sees the sheet being written: switch to it first. */
  const show = async sheet => { if (!local && !sheet.isActive) await book().setActiveWorksheet(sheet.id); };
  const rangeIn = (sheet, { row, column, rowCount, columnCount }) => sheet.getRange({ type: 'cells', row, column, rowCount, columnCount });
  track.done = 0;
  const stopped = () => { if (signal?.aborted) throw agentError('cancelled', 'Stopped by the user.'); };
  for (const [number, step] of steps.entries()) {
    stopped();
    // v1.1: the step's range is selected (the agent's gesture) before anything is written into it.
    if (creationTracks && step.op === 'setValues') {
      const sheet=worksheet(step.sheet);
      await show(sheet);
      await writeGenerationValues(sheet,step,{...creationTracks,signal,wait,track,number});
      track.done=number+1;
      onFinish(number);
      continue;
    }
    if (step.range) {
      const sheet = worksheet(step.sheet);
      await show(sheet);
      await onSelect(number, rangeIn(sheet, step.range), step);
    } else await onSelect(number, undefined, step);
    // A stop while the range was being selected writes nothing.
    stopped();
    let target, tone = 'edit';
    switch (step.op) {
      case 'setValues': {
        const sheet = worksheet(step.sheet);
        await show(sheet);
        target = rangeIn(sheet, step.range);
        const plainRow = row => row.map(value => typeof value === 'string' && value.startsWith('=') ? null : primitive(value));
        const formulas = (row, r) => row.forEach((value, c) => {
          if (typeof value === 'string' && value.startsWith('=')) sheet.getCell(step.range.row + r, step.range.column + c).setCellFormula(value);
        });
        const cells = step.values.reduce((sum, row) => sum + row.length, 0);
        if (rowByRow && cells <= cellByCell) {
          // A small range fills one cell at a time, each selected first: the Sheet's own selection is the highlight of
          // the cell being written, and the agent follows that same cell.
          let written = 0;
          for (const [r, row] of step.values.entries()) {
            for (const [c, value] of row.entries()) {
              if (written) stopped();
              const cell = rangeIn(sheet, { row: step.range.row + r, column: step.range.column + c, rowCount: 1, columnCount: 1 });
              if (!local) { try { cell.select(); } catch {} }
              onWrite(number, cell);
              if (typeof value === 'string' && value.startsWith('=')) sheet.getCell(step.range.row + r, step.range.column + c).setCellFormula(value);
              else cell.setValue([[primitive(value)]]);
              // From the first cell on, a stop leaves this step partly written.
              if (++written < cells) { track.partial = number; await wait(cellPause); }
            }
          }
        } else if (rowByRow) {
          // One real write per row, formulas cell by cell (spec §7.6), so the table fills while the user watches.
          for (const [r, row] of step.values.entries()) {
            if (r) stopped();
            const rowRange = rangeIn(sheet, { row: step.range.row + r, column: step.range.column, rowCount: 1, columnCount: row.length });
            onWrite(number, rowRange);
            rowRange.setValue([plainRow(row)]);
            formulas(row, r);
            // From the first row on, a stop leaves this step partly written.
            if (r < step.values.length - 1) track.partial = number;
            if (r < step.values.length - 1) await wait(pause / 2);
          }
        } else {
          // Plain values in one call, formulas cell by cell (they start with "=").
          target.setValue(step.values.map(plainRow));
          step.values.forEach(formulas);
        }
        track.partial = undefined;
        tone = 'add';
        break;
      }
      case 'setFormula': {
        const sheet = worksheet(step.sheet);
        await show(sheet);
        sheet.getCell(step.range.row, step.range.column).setCellFormula(step.formula);
        target = rangeIn(sheet, step.range);
        tone = 'add';
        break;
      }
      case 'formatRange': {
        const sheet = worksheet(step.sheet);
        await show(sheet);
        target = rangeIn(sheet, step.range);
        // setData writes each cell whole: keep its formula or value and the format it already has, then add the new
        // format. It takes the raw value (getData gives `{ type, value }`); an error cell keeps its shown text.
        const current = plain(target.getData()) ?? [];
        const extra = dataOf(step.format);
        target.setData(current.map(row => (row ?? []).map(cell => {
          const kept = keptStyle(cell);
          if (cell?.formula) return { ...kept, ...extra, formula: String(cell.formula).startsWith('=') ? cell.formula : `=${cell.formula}` };
          if (cell?.value && (cell.value.type === 'primitive' || cell.value.type === 'date')) return { ...kept, ...extra, value: cell.value.value };
          if (cell?.value && typeof cell.text === 'string' && cell.text) return { ...kept, ...extra, text: cell.text };
          return { ...kept, ...extra };
        })));
        break;
      }
      case 'clearRange': {
        const sheet = worksheet(step.sheet);
        await show(sheet);
        target = rangeIn(sheet, step.range);
        onStep(number, target, 'delete');
        await wait(pause);
        target.setValue(Array.from({ length: step.range.rowCount }, () => Array.from({ length: step.range.columnCount }, () => null)));
        tone = 'delete';
        break;
      }
      case 'addSheet': {
        book().addWorksheet(step.name);
        const sheet = worksheet(step.name);
        await show(sheet);
        break;
      }
      case 'renameSheet': worksheet(step.sheet).rename(step.name); break;
      case 'activateSheet': await book().setActiveWorksheet(worksheet(step.sheet).id); break;
      case 'insertRows': { const sheet = worksheet(step.sheet); await show(sheet); sheet.addRows(step.row, step.count); break; }
      case 'deleteRows': { const sheet = worksheet(step.sheet); await show(sheet); sheet.deleteRows(step.row, step.count); tone = 'delete'; break; }
      case 'insertColumns': { const sheet = worksheet(step.sheet); await show(sheet); sheet.addColumns(step.column, step.count); break; }
      case 'deleteColumns': { const sheet = worksheet(step.sheet); await show(sheet); sheet.deleteColumns(step.column, step.count); tone = 'delete'; break; }
      case 'setColumnWidth': {
        const sheet = worksheet(step.sheet);
        await show(sheet);
        sheet.setColumnsWidth(Array.from({ length: step.count }, (_, i) => ({ column: step.column + i, width: step.width })));
        break;
      }
      case 'freezeRows': { const sheet = worksheet(step.sheet); await show(sheet); sheet.setFrozenRowCount(step.count); break; }
      case 'addChart': {
        // The Sheet SDK makes a chart of the selected range (`charts`, office-kit K16). Left to
        // itself it anchors the chart inside the range (0.0.39-test.11: over B3:D5 of A1:D5) and hides the data, so it goes one
        // blank row below.
        const sheet = worksheet(step.sheet);
        await show(sheet);
        target = rangeIn(sheet, step.range);
        if (typeof sheet.charts?.addFromSelection !== 'function') throw agentError('unsupported-operation', 'This editor cannot add charts (its editor package lacks the Sheet chart members). Tell the user plainly.');
        target.select();
        await sheet.charts.addFromSelection({ chart: { ...(step.title && { title: step.title }), position: chartPosition(step.range) } });
        tone = 'add';
        break;
      }
    }
    track.done = number + 1;
    if (!local && target) { try { target.select(); } catch {} }
    onStep(number, target, tone);
    onFinish(number);
    if (number < steps.length - 1) await wait(pause);
  }
}

/** Where a chart of `range` goes: one blank row under it, from its first column (0-based, Sheet SDK chart position). */
export const chartPosition = ({ row, column, rowCount }) => ({ row: row + rowCount + 1, column, offsetX: 0, offsetY: 0 });

/** Follows `range` on screen: the SDK's own position of it in this page, clipped to the window. */
export function rangeLocator(range) {
  return () => {
    let box;
    try { box = range?.getBounding?.(); } catch { box = undefined; }
    return box && box.width > 0 && box.height > 0 ? { left: box.left, top: box.top, width: box.width, height: box.height } : undefined;
  };
}

export function createKitSheetAgent({ editor, document, animate = () => true, locale = 'zh', wait = ms => new Promise(resolve => setTimeout(resolve, ms)), pauseMs = 150, cellPauseMs = 40 }) {
  const api = () => editor.api.sheet;
  const labels = SHEET_LABELS[locale] ?? SHEET_LABELS.zh;
  const overlay = createDocOverlay(document, { animate, wait });
  const creation = createGenerationSurface(overlay,locale);
  let applying = false;
  /** apply's lock, for the selection rewrite's own writes (selection-edit.mjs). */
  const locked = async run => {
    if (applying) throw agentError('busy', 'Another edit is still running.');
    applying = true;
    try { return await run(); }
    finally { try { await editor.prepare?.(); } catch {} await wait(100); applying = false; }
  };
  const agent = {
    get applying() { return applying; },
    handOff() {},
    get creating() { return creation.active; },
    beginCreation(area) {
      const sheets = Array.from(api().workbook.getWorksheets());
      const sheet = sheets.find(item => item.isActive) ?? sheets[0];
      const planned = parseRange(area);
      const first = sheet.getRange({type:'cells',row:0,column:0,rowCount:1,columnCount:1});
      const target = planned && sheet.getRange({type:'cells',...planned});
      return creation.start(()=>{
        const cell = first.getBounding();
        if (target) return creationBands(rangeLocator(target)(),planned.rowCount);
        const canvas = document.elementFromPoint(cell.left+2,cell.top+2)?.closest('canvas')?.getBoundingClientRect()
          ?? Array.from(document.querySelectorAll('canvas')).map(el=>el.getBoundingClientRect()).find(r=>r.left<=cell.left+2 && r.right>cell.left+cell.width && r.top<=cell.top+2 && r.bottom>cell.top+cell.height);
        if (!canvas) return creationBands(cell);
        return creationBands({left:cell.left,top:cell.top,width:canvas.right-cell.left,height:canvas.bottom-cell.top,clip:canvas},Math.round((canvas.bottom-cell.top)/Math.max(1,cell.height)));
      });
    },
    async endGeneration() {
      if (!creation.finish()) return;
      const book=api().workbook, first=book.getWorksheets()[0];
      if (!first) return;
      if (!first.isActive) await book.setActiveWorksheet(first.id);
      first.getRange({type:'cells',row:0,column:0,rowCount:1,columnCount:1}).select();
    },
    inspect: async () => inspectWorkbook(api()),
    async apply({ ops }, signal, local = false) {
      if (applying) throw agentError('busy', 'Another edit is still running.');
      applying = true;
      if (!local) overlay.busy(true, creation.active ? undefined : labels.editing);
      const track = { done: 0 };
      let total = Array.isArray(ops) ? ops.length : 0;
      try {
        const sheetApi = api();
        const { steps } = validateXlsxOps(ops, inspectWorkbook(sheetApi));
        total = steps.length;
        if (!local && steps.some(step=>['setValues','setFormula','addChart'].includes(step.op))) creation.writing();
        const key = number => `t${number % 3}`;
        await runXlsxSteps(sheetApi, steps, {
          creationTracks: !local && creation.active && animate() ? {
            select:(lane,range)=>overlay.select('creating-'+lane,()=>creationBands(rangeLocator(range)()),{
              edge:true,writer:true,label:labels.marks.add,selecting:agentLabel(locale,'selecting')}),
            finish:lane=>overlay.drop('creating-'+lane),
          } : undefined,
          signal, wait, local, pause: animate() ? pauseMs : 0, track, rowByRow: animate(), cellByCell: animate() ? XLSX_ANIM_CELLS : 0, cellPause: animate() ? cellPauseMs : 0,
          onSelect: async (number, range, step) => {
            if (local) return;
            if (!creation.active) overlay.say(labels.step(number + 1, total));
            await overlay.select(key(number), range ? rangeLocator(range) : () => undefined, { label: agentLabel(locale, xlsxStatus(step)), selecting: agentLabel(locale, 'selecting'), edge: true });
          },
          onWrite: (number, range) => !local && overlay.point(key(number), rangeLocator(range)),
          onFinish: number => !local && overlay.release(key(number)),
        });
        return { applied: steps.length };
      } catch (error) {
        const code = signal?.aborted ? 'cancelled' : PAGE_CODES.has(error?.code) ? error.code : 'editor-failed';
        const reason = signal?.aborted ? 'Stopped by the user.' : (error?.message ?? String(error));
        throw agentError(code, `${reason} ${xlsxProgress(track.done, total)}${Number.isInteger(track.partial) ? ` Step ${track.partial + 1} was partly applied (part of it is written); check it with spreadsheet_read.` : ''}`);
      } finally {
        // Sheet changes reach the listener as the SDK commits them; finish editing and let late ones through as the AI's.
        try { await editor.prepare?.(); } catch {}
        await wait(100);
        applying = false;
        if (!local) overlay.busy(false);
      }
    },
    async undo() { throw agentError('unsupported-operation', 'Spreadsheet edits cannot be undone by the AI (there is no rollback); change the cells back with another spreadsheet_edit.'); },
    dispose() { overlay.dispose(); },
  };
  // Selection rewrite (spec 261109-selection-ai-edit): read the user's cells, replace them, take them back.
  Object.assign(agent, sheetSelectionMethods({ api, locked, overlay, locale, applying: () => applying, apply: (...args) => agent.apply(...args) }));
  return agent;
}

export { columnLetters };
