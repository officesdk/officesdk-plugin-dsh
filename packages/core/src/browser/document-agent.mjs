import { createTableStage } from './table-stage.mjs';
import { DOCX_ALIGN, DOCX_LIVE_MARK, DOCX_OPS, DOCX_OUTLINE_STYLES, DOCX_STYLES, DOCX_STYLE_FALLBACK, createDocxReadHistory, docxProgress, validateDocxOps } from '../agent/docx.mjs';
import { createGenerationSurface, creationBands, writerOpening, writerPaper } from './creation-surface.mjs';
import { DOC_LABELS, createDocOverlay } from './doc-overlay.mjs';
import { agentLabel, docxStatus } from './agent-status.mjs';
import { bodyBlock, findParagraphElements, paragraphRange, rememberWriterIndents, tableLocator, unionRects } from './writer-geometry.mjs';
export { findParagraphElements, boxOf } from './writer-geometry.mjs';
import { textChunks } from './selection.mjs';
import { docSelectionMethods, previousWriterChange, previewWriterPrefixes, startWriterWork } from './selection-edit.mjs';

/**
 * AI editing of a Word document on a Kit page (spec 261105 §3.3), through office-kit's public `editor.api.writer`
 * (Word JS API). Each operation is one `Word.run`; the user sees the document change step by step. There is no
 * rollback (office-kit K3): when a step fails or the user stops, the steps already done stay in the editor and the
 * error says how far it got. Changes reach Kit through the page's change queue like a hand edit, but while `applying`
 * they are marked as the AI's, so they wait for a save instead of autosaving.
 */
const agentError = (code, message) => Object.assign(new Error(message), { code });
const PROJECT_BLANK = '\u2800';
const isTemplateBlank = text => typeof text === 'string' && text.replace(/[\r\n]+/g,'') === PROJECT_BLANK;
/** Codes the page passes on as they are; anything else from the editor is `editor-failed`. */
const PAGE_CODES = new Set(['invalid-edit', 'source-changed', 'cancelled', 'busy', 'read-only', 'unsupported-operation']);

/** The document as the model reads it: paragraphs (text, style, list, inside a table) and tables. */
export async function inspectDocument(Word, document) {
  let doc;
  await Word.run(async ctx => {
    const paragraphs = ctx.document.body.paragraphs;
    paragraphs.load('items/text,items/style,items/isListItem,items/tableNestingLevel' + (document ? ',items/leftIndent,items/rightIndent' : ''));
    const tables = ctx.document.body.tables;
    tables.load('items/rowCount,items/values');
    await ctx.sync();
    rememberWriterIndents(document, paragraphs.items);
    // Plain copies only: office-kit's Writer API hands out arrays as proxies, which postMessage cannot clone.
    doc = {
      paragraphs: Array.from(paragraphs.items, paragraph => ({ text: String(paragraph.text ?? ''), style: paragraph.style ? String(paragraph.style) : undefined, list: paragraph.isListItem === true, table: Number(paragraph.tableNestingLevel ?? 0) > 0 })),
      tables: Array.from(tables.items, table => {
        const values = table.values ? Array.from(table.values, row => Array.from(row ?? [], cell => String(cell ?? ''))) : [];
        return { rows: Number(table.rowCount ?? values.length), columns: values[0]?.length ?? 0, values };
      }),
    };
  });
  return doc;
}

/**
 * Formatting through `set()`: office-kit's Writer API refuses property assignment (its proxies only allow method calls),
 * and takes the font through `font.set()` (a nested `font` in the paragraph's `set()` is not supported).
 */
function applyFormat(paragraph, format = {}) {
  const own = {}, font = {};
  if (format.align) own.alignment = DOCX_ALIGN[format.align] ?? format.align;
  for (const key of ['firstLineIndent', 'leftIndent', 'lineSpacing', 'spaceBefore', 'spaceAfter']) if (format[key] !== undefined) own[key] = format[key];
  for (const key of ['bold', 'italic', 'size', 'color']) if (format[key] !== undefined) font[key] = format[key];
  if (format.underline !== undefined) font.underline = format.underline ? 'Single' : 'None';
  if (format.font !== undefined) font.name = format.font;
  if (Object.keys(own).length) paragraph.set(own);
  if (Object.keys(font).length) paragraph.font.set(font);
}

/**
 * Runs `steps` (validateDocxOps output) one `Word.run` each. `order` holds, in document order, a key for every
 * paragraph: `#n` for paragraph n of the read, `@ref` or `+i` for paragraphs this batch made; each structural step
 * reads the paragraph count before and after and inserts (or removes) that many keys where it acted, so numbers from
 * the read keep naming the same paragraphs. `tableOrder` does the same for tables.
 */
export async function runDocxSteps(Word, steps, doc, { signal, onStep = () => {}, onDone = () => {}, onSelect = async () => {}, onWrite = () => {}, onFinish = () => {}, chunk = text => [text], tracks = 3, wait = () => Promise.resolve(), pause = 0, writingPause = 0, track = {}, projectCreation = false, templateCreation = false, meetingCreation = false, proposalCreation = false } = {}) {
  const order = doc.paragraphs.map((_, number) => `#${number}`);
  const tableOrder = doc.tables.map((_, number) => `#${number}`);
  const texts = new Map(doc.paragraphs.map((paragraph, number) => [`#${number}`, paragraph.text]));
  const tableParagraphs = new Set(order.filter((key, i) => doc.paragraphs[i].table));
  const tableRows = new Map(doc.tables.map((table, i) => [tableOrder[i], table.rows]));
  const blockOf = index => bodyBlock(order, index, key => tableParagraphs.has(key));
  const tableBlock = index => {
    const starts = order.flatMap((key, i) => tableParagraphs.has(key) && !tableParagraphs.has(order[i - 1]) ? [i] : []);
    return starts.length === tableOrder.length && starts[index] !== undefined ? blockOf(starts[index]) : undefined;
  };
  const refs = {};
  // The caller reads where the batch got to even when it stops part way.
  Object.assign(track, { order, tableOrder, refs });
  const state = { done: 0, styleFallback: 0, emptyStart: doc.paragraphs.length === 1 && !doc.paragraphs[0].text && !doc.tables.length };
  // The style id most body paragraphs of the read use: a paragraph whose built-in style is missing starts from it, so it
  // does not keep what it took over from its neighbour (a list paragraph's indent, a heading's size). The plugin templates'
  // own layout styles (redhead's Shimo Issuer / Shimo Doc Number: red 42 pt, red rule) are never body text.
  const bodyStyle = (() => {
    const counts = new Map();
    for (const paragraph of doc.paragraphs) if (paragraph.text && !paragraph.list && !paragraph.table && paragraph.style && !/^Shimo/i.test(paragraph.style)) counts.set(paragraph.style, (counts.get(paragraph.style) ?? 0) + 1);
    return [...counts].sort((a, b) => b[1] - a[1])[0]?.[0];
  })();
  const keyOf = target => target.ref !== undefined ? `@${target.ref}` : `#${target.paragraph}`;
  const indexOf = target => {
    const at = order.indexOf(keyOf(target));
    if (at < 0) throw agentError('editor-failed', `Paragraph ${target.ref !== undefined ? `@${target.ref}` : target.paragraph} is no longer in the document.`);
    return at;
  };
  const tableIndexOf = target => {
    const at = tableOrder.indexOf(target.ref !== undefined ? `@${target.ref}` : `#${target.table}`);
    if (at < 0) throw agentError('editor-failed', `Table ${target.ref !== undefined ? `@${target.ref}` : target.table} is no longer in the document.`);
    return at;
  };
  /** Where a step's new paragraphs start, in the current document, given its `where`. */
  const startOf = (where, count) => where.at === 'start' ? 0 : where.at === 'end' ? count : where.after ? indexOf(where.after) + 1 : indexOf(where.before);
  const load = async ctx => {
    const paragraphs = ctx.document.body.paragraphs;
    paragraphs.load('items/tableNestingLevel,items/isListItem');
    await ctx.sync();
    return paragraphs.items;
  };
  const expectCount = (step, before, after, { exact, atLeast }) => {
    const delta = after - before;
    if (exact !== undefined ? delta !== exact : delta < atLeast) throw agentError('editor-failed', `The editor did not ${step.op === 'deleteParagraph' ? 'delete' : 'insert'} that as planned (paragraph count changed by ${delta}).`);
    return delta;
  };
  /** Tables wholly before paragraph position `at` (a table starts where a table paragraph follows a body one). */
  const tablesBefore = (items, at) => {
    let count = 0;
    for (let i = 0; i < Math.min(at, items.length); i++) if ((items[i].tableNestingLevel ?? 0) > 0 && (i === 0 || (items[i - 1].tableNestingLevel ?? 0) === 0)) count++;
    return count;
  };
  let marked;
  const isMarked = async () => {
    if (marked === undefined) {
      try { await Word.run(async ctx => { const styles = ctx.document.getStyles(); styles.load('items/nameLocal'); await ctx.sync(); marked = Array.from(styles.items, item => String(item.nameLocal)).includes(DOCX_LIVE_MARK); }); }
      catch { marked = false; }
    }
    return marked;
  };
  // Body text in a document that is not a live template: Word copies the paragraph before into a new one, so text after a
  // directly formatted heading would come out large and centred. It is set back to the look of the document's own body
  // text (the first plain paragraph in the body style): alignment and size, not bold or italic.
  let reference;
  const bodyLook = async () => {
    if (reference !== undefined) return reference;
    reference = { bold: false, italic: false };
    const number = doc.paragraphs.findIndex(paragraph => paragraph.text && !paragraph.list && !paragraph.table && paragraph.style === bodyStyle);
    // Where that paragraph of the read is now (earlier steps may have inserted or deleted before it).
    const at = number < 0 ? -1 : order.indexOf(`#${number}`);
    if (at < 0) return reference;
    try {
      await Word.run(async ctx => {
        const paragraph = (await load(ctx))[at];
        paragraph.load('alignment,font/size');
        await ctx.sync();
        const align = Object.entries(DOCX_ALIGN).find(([, value]) => value === paragraph.alignment)?.[0];
        if (align) reference.align = align;
        if (typeof paragraph.font?.size === 'number' && paragraph.font.size > 0) reference.size = paragraph.font.size;
      });
    } catch {}
    return reference;
  };
  const style = async (index, name) => {
    // Styles apply by id; a document without that id (Chinese Word files use "1", "a3"…) gets direct formatting, and so
    // do titles and headings outside shimo-docx's own templates (their outline level would make Kit fail to save, K15).
    try {
      if (DOCX_OUTLINE_STYLES.includes(name) && !await isMarked()) throw new Error('not a live template');
      await Word.run(async ctx => { (await load(ctx))[index].set({ styleBuiltIn: DOCX_STYLES[name] }); await ctx.sync(); });
    } catch {
      state.styleFallback++;
      if (bodyStyle) { try { await Word.run(async ctx => { (await load(ctx))[index].set({ style: bodyStyle }); await ctx.sync(); }); } catch {} }
      await Word.run(async ctx => { applyFormat((await load(ctx))[index], DOCX_STYLE_FALLBACK[name]); await ctx.sync(); });
    }
    if ((name === 'body' || name === 'quote') && !await isMarked()) {
      const look = await bodyLook();
      await Word.run(async ctx => { applyFormat((await load(ctx))[index], name === 'quote' ? { ...look, italic: true } : look); await ctx.sync(); });
    }
  };
  const list = async (index, kind) => {
    await Word.run(async ctx => {
      const items = await load(ctx);
      const paragraph = items[index], previous = items[index - 1];
      let parent;
      if (previous?.isListItem) { parent = previous.list; parent.load('id,levelTypes'); await ctx.sync(); }
      // A list item continues the list right before it when that list is the same kind; otherwise it starts its own.
      const same = parent && (parent.levelTypes?.[0] === undefined || (String(parent.levelTypes[0]).toLowerCase() === 'bullet') === (kind === 'bullet'));
      if (same) {
        if (!paragraph.isListItem) paragraph.attachToList(parent.id, 0);
      } else {
        if (paragraph.isListItem) { paragraph.detachFromList(); await ctx.sync(); }
        const created = paragraph.startNewList();
        await ctx.sync();
        if (kind === 'bullet') created.setLevelBullet(0, 'Solid');
      }
      await ctx.sync();
    });
  };
  const unlist = async index => {
    // A paragraph inserted next to a list item joins that list; one that should not be in a list leaves it.
    await Word.run(async ctx => { const paragraph = (await load(ctx))[index]; if (paragraph.isListItem) { paragraph.detachFromList(); await ctx.sync(); } });
  };
  const cursorTo = async index => {
    // The editor scrolls its caret into view: the paragraph being written stays on screen. Best effort.
    try { await Word.run(async ctx => { (await load(ctx))[index].getRange('End').select(); await ctx.sync(); }); } catch {}
  };
  /**
   * v1.1 (spec §7.6): a new paragraph is written in parts (`chunk`); up to `tracks` consecutive new paragraphs are
   * written round-robin, each by its own agent track, so they grow together. `pending` holds the parts still to write.
   */
  const pending = [];
  // The caller reads, after a stop or failure, every paragraph left half written and every step finished out of order.
  track.pending = pending;
  // Steps done = the leading steps that are complete (a paragraph still being written holds the count back).
  const complete = new Set();
  const finished = number => { complete.add(number); while (complete.has(state.done)) state.done++; track.done = state.done; track.complete = complete; onFinish(number); };
  const flush = async () => {
    while (pending.some(entry => entry.parts.length)) {
      for (const entry of pending) {
        if (!entry.parts.length) continue;
        if (signal?.aborted) { track.partial = pending.find(item => item.parts.length)?.number; throw agentError('cancelled', 'Stopped by the user.'); }
        const text = entry.parts[0];
        const at = order.indexOf(entry.key);
        if (at < 0) throw agentError('editor-failed', 'A paragraph being written is no longer in the document.');
        // Each part adds to the end of the paragraph, so the formatting set when it was inserted stays.
        const added = text.slice((texts.get(entry.key) ?? '').length);
        await Word.run(async ctx => { (await load(ctx))[at].insertText(added, 'End'); await ctx.sync(); });
        entry.parts.shift();
        texts.set(entry.key, text);
        onWrite(entry.number, { index: at, block: blockOf(at), text });
        if (!entry.parts.length) {
          await cursorTo(at);
          finished(entry.number);
        }
      }
      if (writingPause && pending.some(entry=>entry.parts.length)) await wait(writingPause);
    }
    pending.length = 0;
    track.partial = undefined;
  };
  /** What a step's agent selects before writing: an existing paragraph, or the place a new one will go. */
  const selectTarget = step => {
    if (step.op === 'insertParagraph' || step.op === 'insertPageBreak' || step.op === 'insertTable') {
      const where = step.where;
      const index = where.at === 'start' ? 0 : where.at === 'end' ? order.length - 1 : indexOf(where.after ?? where.before);
      return { insert: where.before || where.at === 'start' ? 'before' : 'after', index, text: texts.get(order[index]) ?? '' };
    }
    if (step.target) {
      const index = indexOf(step.target), text = texts.get(order[index]) ?? '';
      if (step.find !== undefined) {
        const start = Array.from(text.slice(0, text.indexOf(step.find))).length;
        return { index, text, start, end: start + Array.from(step.find).length };
      }
      return { index, text };
    }
    return undefined;
  };
  for (const [number, step] of steps.entries()) {
    if (signal?.aborted) { track.partial = pending.find(item => item.parts.length)?.number; throw agentError('cancelled', 'Stopped by the user.'); }
    // Anything but another new paragraph, or a full set of tracks, first finishes the paragraphs being written.
    if (pending.length && (step.op !== 'insertParagraph' || pending.length >= tracks)) await flush();
    let target = selectTarget(step);
    if (step.op === 'setCell') {
      const table = tableIndexOf(step.table);
      target = { tableBlock: tableBlock(table), row: step.row, column: step.column, rows: tableRows.get(tableOrder[table]) };
    } else if (target) target.block = blockOf(target.index);
    const staged = await onSelect(number, target, step, () => order.map(id=>texts.get(id)));
    // A stop while the target was being selected writes nothing.
    if (signal?.aborted) { track.partial = pending.find(item => item.parts.length)?.number; throw agentError('cancelled', 'Stopped by the user.'); }
    let focus, grouped = false;
    switch (step.op) {
      case 'insertParagraph': {
        let at;
        const parts = chunk(step.text, startOf(step.where, order.length));
        await Word.run(async ctx => {
          const items = await load(ctx);
          const before = items.length;
          if (state.emptyStart && step.where.at) {
            // A new document's only, empty paragraph takes the first text instead of staying blank above it.
            items[0].insertText(parts[0], 'Replace');
            await ctx.sync();
            expectCount(step, before, (await load(ctx)).length, { exact: 0 });
            at = 0;
            order[0] = step.ref ? `@${step.ref}` : `+${number}`;
          } else {
            at = startOf(step.where, before);
            if (step.where.at) ctx.document.body.insertParagraph(parts[0], step.where.at === 'start' ? 'Start' : 'End');
            else items[step.where.after ? at - 1 : at].insertParagraph(parts[0], step.where.after ? 'After' : 'Before');
            await ctx.sync();
            expectCount(step, before, (await load(ctx)).length, { exact: 1 });
            order.splice(at, 0, step.ref ? `@${step.ref}` : `+${number}`);
          }
          // From here the paragraph is in the document, even if its style or list fails below.
          track.partial = number;
          if (step.ref) refs[step.ref] = 'paragraph';
        });
        await style(at, step.style ?? 'body');
        if (step.list) await list(at, step.list); else await unlist(at);
        if (step.format) await Word.run(async ctx => { applyFormat((await load(ctx))[at], step.format); await ctx.sync(); });
        texts.set(order[at], parts[0]);
        onWrite(number, { index: at, block: blockOf(at), text: parts[0] });
        if (parts.length > 1) { pending.push({ key: order[at], parts: parts.slice(1), number }); grouped = true; }
        focus = { index: at, text: parts[0], tone: 'add' };
        break;
      }
      case 'setText': {
        const at = indexOf(step.target);
        const projectBlank = (projectCreation || templateCreation) && isTemplateBlank(texts.get(order[at])) && !tableParagraphs.has(order[at]);
        const projectBlock = blockOf(at);
        // A whole rewrite grows in parts like a new paragraph (in this step, not on a track); a replaceText (`find`)
        // changes only part of the paragraph and is written once: growing it would first wipe the rest of the text.
        const parts = staged?.parts ?? (step.find === undefined ? chunk(step.text, at) : [step.text]);
        for (const [part, text] of parts.entries()) {
          if (part) {
            if (signal?.aborted) throw agentError('cancelled', 'Stopped by the user.');
            await wait(pause / 2);
            if (signal?.aborted) throw agentError('cancelled', 'Stopped by the user.');
          }
          const added = part ? text.slice(parts[part - 1].length) : text;
          await Word.run(async ctx => {
            (await load(ctx))[at].insertText(added, part ? 'End' : 'Replace');
            await ctx.sync();
            if (projectBlank && !part) {
              const proposalStyle = doc.paragraphs[at]?.style;
              const color = meetingCreation && [6,9].includes(projectBlock) ? '#30363C' : proposalCreation ? (['Style14','21'].includes(proposalStyle)?'#357CA2':proposalStyle==='Style17'?'#434343':'#000000') : templateCreation === 'letter' ? '#000000' : (meetingCreation ? false : templateCreation ? projectBlock >= 5 && projectBlock % 2 === 1 : [5, 7, 11].includes(projectBlock)) ? '#53606A' : '#202124';
              (await load(ctx))[at].font.set({color});
              await ctx.sync();
            }
          });
          texts.set(order[at], text);
          // From the first part on, a stop leaves this paragraph half rewritten.
          if (part < parts.length - 1) track.partial = number;
          const start = step.find === undefined ? 0 : Array.from(step.before.slice(0, step.before.indexOf(step.find))).length;
          onWrite(number, { index: at, block: blockOf(at), text, start, end: start + Array.from(step.replace ?? text).length });
        }
        // Split the finished line into its template label and value. The marker starts in the label's
        // bold run; reset the appended range so only the label retains emphasis.
        if (meetingCreation && projectBlank && projectBlock >= 1 && projectBlock <= 4) {
          const label = ['会议主题', '会议时间', '参会人', '记录人'][projectBlock - 1];
          if (step.text.startsWith(label)) await Word.run(async ctx => {
            const paragraph = (await load(ctx))[at];
            paragraph.insertText(label, 'Replace');
            await ctx.sync();
            const value = (await load(ctx))[at].insertText(step.text.slice(label.length), 'End');
            value.font.set({ bold: false });
            await ctx.sync();
          });
        }
        focus = { index: at, text: step.text, tone: 'edit' };
        break;
      }
      case 'formatParagraph': {
        const at = indexOf(step.target);
        if (step.style) { await style(at, step.style); track.partial = number; }
        if (step.list) { await list(at, step.list); track.partial = number; }
        if (step.format) await Word.run(async ctx => { applyFormat((await load(ctx))[at], step.format); await ctx.sync(); });
        focus = { index: at, text: texts.get(order[at]), tone: 'edit' };
        break;
      }
      case 'deleteParagraph': {
        const at = indexOf(step.target);
        // Selected first (onSelect above), then it goes.
        onStep(number, { index: at, text: step.before, tone: 'delete' });
        await Word.run(async ctx => {
          const items = await load(ctx);
          const before = items.length;
          items[at].delete();
          await ctx.sync();
          expectCount(step, before, (await load(ctx)).length, { exact: -1 });
          order.splice(at, 1);
        });
        focus = undefined;
        break;
      }
      case 'insertTable': {
        let at;
        await Word.run(async ctx => {
          const items = await load(ctx);
          const before = items.length;
          at = startOf(step.where, before);
          const rows = step.values.length, columns = step.values[0].length;
          const table = step.where.at
            ? ctx.document.body.insertTable(rows, columns, step.where.at === 'start' ? 'Start' : 'End', step.values)
            : items[step.where.after ? at - 1 : at].insertTable(rows, columns, step.where.after ? 'After' : 'Before', step.values);
          if (step.header) for (let column = 0; column < columns; column++) table.getCell(0, column).body.paragraphs.getFirst().font.set({ bold: true });
          await ctx.sync();
          track.partial = number;
          const after = await load(ctx);
          const delta = expectCount(step, before, after.length, { atLeast: rows * columns });
          order.splice(at, 0, ...Array.from({ length: delta }, (_, i) => `+${number}.${i}`));
          for (const key of order.slice(at, at + delta)) tableParagraphs.add(key);
          const tableKey = step.ref ? `@${step.ref}` : `+${number}`;
          tableOrder.splice(tablesBefore(after, at), 0, tableKey);
          tableRows.set(tableKey, rows);
          if (step.ref) refs[step.ref] = 'table';
        });
        onWrite(number, { tableBlock: blockOf(at) });
        focus = { index: at, text: step.values[0][0], tone: 'add' };
        break;
      }
      case 'setCell': {
        const at = tableIndexOf(step.table);
        await Word.run(async ctx => {
          const tables = ctx.document.body.tables;
          tables.load('items');
          await ctx.sync();
          // `TableCell.value` is accepted but not applied by the Writer API; the cell's first paragraph is.
          tables.items[at].getCell(step.row, step.column).body.paragraphs.getFirst().insertText(step.text, 'Replace');
          await ctx.sync();
        });
        focus = { index: 0, text: step.text, tone: 'edit', cell: true };
        break;
      }
      case 'insertPageBreak': {
        await Word.run(async ctx => {
          const items = await load(ctx);
          const before = items.length;
          const at = startOf(step.where, before);
          const anchor = step.where.at === 'start' ? items[0] : step.where.at === 'end' ? items[before - 1] : items[step.where.after ? at - 1 : at];
          anchor.insertBreak('Page', step.where.at === 'start' || step.where.before ? 'Before' : 'After');
          await ctx.sync();
          track.partial = number;
          const delta = expectCount(step, before, (await load(ctx)).length, { atLeast: 0 });
          order.splice(at, 0, ...Array.from({ length: delta }, (_, i) => `+${number}.${i}`));
        });
        break;
      }
    }
    // Only a batch's very first insert may fill a new document's empty paragraph; any step ends that.
    state.emptyStart = false;
    if (focus && !focus.cell) await cursorTo(focus.index);
    await onDone(number, focus);
    if (!grouped) { if (!pending.length) track.partial = undefined; finished(number); }
    if (number < steps.length - 1) await wait(pause);
  }
  await flush();
  return { refs: refsNow(track), styleFallback: state.styleFallback, state };
}

/** Each ref of a batch as a number in the document now: `{ paragraph: n }` or `{ table: n }`, for the next call. */
export function refsNow(track) {
  const out = {};
  for (const [ref, kind] of Object.entries(track?.refs ?? {})) {
    const at = (kind === 'table' ? track.tableOrder : track.order)?.indexOf(`@${ref}`) ?? -1;
    if (at >= 0) out[ref] = { [kind]: at };
  }
  return out;
}

export function createKitDocumentAgent({ editor, document, animate = () => true, locale = 'zh', wait = ms => new Promise(resolve => setTimeout(resolve, ms)), pauseMs = 120, enabled = DOCX_OPS }) {
  const Word = () => editor.api.writer.Word;
  const reads = createDocxReadHistory();
  const labels = DOC_LABELS[locale] ?? DOC_LABELS.zh;
  const overlay = createDocOverlay(document, { animate, wait, showStatus: false, followParagraphPaint: true });
  const creation = createGenerationSurface(overlay,locale);
  let tableStage, tableFrozen = false, projectMode = false, templateMode = false, meetingMode = false, proposalMode = false;
  let applying = false, settled, release;
  const beginWrite = () => {
    applying = true;
    settled = new Promise(resolve => { release = resolve; });
  };
  const finishWrite = () => {
    applying = false;
    release();
  };
  /** The lock and overlay of apply, for the selection rewrite's in-paragraph writes (selection-edit.mjs). */
  const locked = async (run, signal) => {
    while (applying) await settled;
    if (signal?.aborted) throw agentError('cancelled', 'Stopped by the user.');
    beginWrite();
    overlay.busy(true, labels.editing);
    try { return await run(); }
    finally { try { await editor.prepare?.(); } catch {} await wait(100); finishWrite(); overlay.busy(false); }
  };
  const agent = {
    /** True while an AI edit is writing: its changes are not the user's. */
    get applying() { return applying; },
    handOff() {},
    get creating() { return creation.active; },
    beginTemplateCreation(mode='template') { templateMode = mode; },
    beginMeetingCreation() { meetingMode = true; },
    beginProposalCreation() { proposalMode = true; },
    beginCreation(area) {
      return creation.start(()=>{
        const page = writerPaper(document);
        if (!page) return;
        if (area === 'opening') return creationBands(writerOpening(document,page));
        const bounds = page.getBoundingClientRect();
        return creationBands({left:bounds.left,top:bounds.top,width:bounds.width,height:bounds.height,
          clip:page.closest('.writer-document-preview__scroll')?.getBoundingClientRect()},Math.ceil(bounds.height/27));
      });
    },
    async beginTableStage() {
      if (tableStage?.active) throw agentError('invalid-edit', 'A table is already being prepared.');
      const doc = await inspectDocument(Word(), document);
      const index = doc.paragraphs.findIndex(p => p.table);
      if (doc.tables.length !== 1 || index < 0) throw agentError('invalid-edit', 'The project plan needs exactly one existing table.');
      const block = bodyBlock(doc.paragraphs, index);
      const locate = tableLocator(document, block);
      // An absent paint target must fail closed; never present an unmasked placeholder as completed.
      for (let attempt = 0; attempt < 40 && !locate(); attempt++) await wait(50);
      if (!locate()) throw agentError('editor-failed', 'The table is not visible yet; try again after layout.');
      tableStage = createTableStage({ document, overlay, locate });
      await tableStage.begin();
      projectMode = true;
    },
    async fillProjectTable(milestones, kind = 'project') {
      if (!tableStage?.active) throw agentError('invalid-edit', 'No table is being prepared.');
      if (!Array.isArray(milestones) || !milestones.length) throw agentError('invalid-edit', 'Provide milestones.');
      if (projectMode && (await inspectDocument(Word(), document)).paragraphs.some(p => !p.table && isTemplateBlank(p.text)))
        throw agentError('invalid-edit', 'Complete the project plan paragraphs before finishing the table.');
      beginWrite();
      try {
        await Word().run(async ctx => {
          const tables=ctx.document.body.tables;tables.load('items');await ctx.sync();
          const table=tables.items[0];table.load('rowCount,values');await ctx.sync();
          if(table.rowCount !== milestones.length+1) throw agentError('invalid-edit','The table row count changed.');
          const fields=kind==='meeting'?['decision','owner','deadline']:kind==='proposal'?['description','quantity','unitPrice','amount']:['stage','deliverable','owner','deadline'];
          if ((kind==='meeting') !== meetingMode || (kind==='proposal') !== proposalMode || !Array.isArray(table.values) || table.values.slice(1).some(row => !Array.isArray(row) || row.length !== fields.length || row.some(value => value !== '待填写')))
            throw agentError('source-changed','The table was edited during generation; its contents were kept.');
          for(let i=0;i<milestones.length;i++) for(let col=0;col<fields.length;col++) {
            const value=milestones[i][fields[col]] || '待填写';
            table.getCell(i+1,col).body.paragraphs.getFirst().insertText(value,'Replace');
          }
          await ctx.sync();
        });
      } finally { finishWrite(); }
    },
    async freezeTableStage() {
      if (!tableStage?.active) throw agentError('invalid-edit', 'No table is being prepared.');
      await tableStage.activate();
      tableFrozen = true; overlay.busy(true, labels.editing);
    },
    endTableStage() {
      tableStage?.end(); tableStage = undefined; tableFrozen = false; overlay.busy(false);
    },
    async endGeneration() {
      if (!tableFrozen && tableStage?.active) { tableStage.end(); tableStage = undefined; }
      if (!creation.finish()) return;
      // Clear only untouched markers, including when a generation is stopped. User text is never replaced here.
      const remaining = (projectMode || templateMode) ? await inspectDocument(Word(), document) : {paragraphs:[]};
      const blanks = remaining.paragraphs.flatMap((p, i) => !p.table && isTemplateBlank(p.text) ? [i] : []);
      projectMode = false;
      templateMode = false;
      meetingMode = false;
      proposalMode = false;
      if (blanks.length) {
        beginWrite();
        try { await Word().run(async ctx => {
          const ps=ctx.document.body.paragraphs;ps.load('items/text');await ctx.sync();
          for (const i of blanks) if (isTemplateBlank(ps.items[i]?.text)) ps.items[i].insertText('', 'Replace');
          await ctx.sync();
        }); } finally { finishWrite(); }
      }
      await overlay.waitForParagraphCursors();
      if (creation.active) return;
      if (document.hasFocus()) {
        await Word().run(async ctx => {
          ctx.document.body.paragraphs.getFirst().getRange('Start').select();
          await ctx.sync();
        });
        // Writer scrolls the new caret into view on its next paint. Let that settle first.
        await wait(45);
      }
      const scroll = document.querySelector('.writer-document-preview__scroll');
      if (scroll) {
        const first = writerPaper(document)?.getBoundingClientRect();
        const top = first ? scroll.scrollTop + first.top - scroll.getBoundingClientRect().top : 0;
        scroll.scrollTo({ top: Math.max(0, top), behavior: 'instant' });
      }
    },
    async inspect({ checkEdit } = {}) {
      const doc = await inspectDocument(Word(), document);
      if (checkEdit) {
        reads.check(checkEdit.ops, doc, checkEdit.readToken);
        return { ...doc, checkedEdit: true };
      }
      reads.remember(doc);
      return doc;
    },
    async apply({ ops, readToken, editId }, signal, selectionWork, chatBatch) {
      while (applying) await settled;
      if (signal?.aborted) throw agentError('cancelled', 'Stopped by the user.');
      beginWrite();
      overlay.busy(true, labels.editing);
      let total = Array.isArray(ops) ? ops.length : 0;
      const track = { done: 0 };
      const batchFrames=new Map(), writtenChanges=new Map(), previousChanges=new Map();
      try {
        const word = Word();
        const doc = await inspectDocument(word, document);
        // Checked here, after the approval card: the user may have edited while it was open.
        reads.check(ops, doc, readToken);
        const { steps } = validateDocxOps(ops, doc, { enabled });
        total = steps.length;
        if (!selectionWork && steps.some(step=>['insertParagraph','setText','replaceText','insertTable','setCell'].includes(step.op))) creation.writing();
        // v1.1: each step's agent (one of three tracks) selects first, then writes; new paragraphs grow in parts.
        // A step keeps its track until it finishes; a new step takes a free one.
        const owners = new Map();
        const key = number => {
          if (!owners.has(number)) owners.set(number, ['t0', 't1', 't2'].find(name => ![...owners.values()].includes(name)) ?? `t${number % 3}`);
          return owners.get(number);
        };
        const locate = (target, workBands = false) => {
          if (!target) return () => undefined;
          if ('tableBlock' in target) return target.tableBlock === undefined ? () => undefined : tableLocator(document, target.tableBlock, target);
          const at = target.insert === 'before' ? 0 : Array.from(target.text).length;
          return paragraphRange(document, target.text, target.block, target.insert ? at : target.start, target.insert ? at : target.end, true, workBands);
        };
        const writtenRanges = new Map();
        const followWritten = (number, written) => {
          const step=steps[number];
          if(chatBatch && editId && step.op==='setText' && written.text===step.text){
            const before=step.find ?? step.before, after=step.replace ?? step.text;
            if(before!==after && before && !doc.paragraphs[written.index]?.table){
              const anchor={paragraphs:[written.index,written.index],start:written.start,end:written.end};
              writtenChanges.set(number,{id:editId+'-'+number,kind:'docx',before,after,appliedAt:Date.now(),
                revert:{anchor,original:after,text:before},located:{anchor,view:{first:written.index,texts:[written.text],blocks:[written.block]}}});
            }
          }
          if (!selectionWork) return overlay.point(batchFrames.get(number)?.key ?? key(number), locate(written, creation.active || batchFrames.has(number)));
          const {after} = selectionWork;
          const [first, last] = after.paragraphs;
          const length = Array.from(written.text).length;
          writtenRanges.set(written.index, paragraphRange(document,written.text,written.block,
            written.index === first ? Math.min(after.start,length) : 0, written.index === last ? Math.min(after.end,length) : length,true,true));
          overlay.point(selectionWork.key, () => {
            const ranges = [...writtenRanges].sort(([a],[b]) => a-b).map(([,find]) => find()).filter(Boolean);
            const rects = ranges.flatMap(range => range.rects);
            return rects.length ? { ...unionRects(rects), rects, clip:ranges[0].clip } : undefined;
          });
        };
        const moving=animate() && !selectionWork?.reserved;
        const chunk=(text,index)=>{
          if (!moving) return [text];
          const parts=textChunks(text);
          if (!selectionWork?.previewLines) return parts;
          const offset=index-selectionWork.after.paragraphs[0], shown=selectionWork.previewLines[offset];
          if (shown === undefined) return parts;
          const count=Array.from(shown).length+(offset===0 ? selectionWork.after.start : 0);
          return parts.filter((part,i)=>i===parts.length-1 || Array.from(part).length>=count);
        };
        const outcome = await runDocxSteps(word, steps, doc, {
          signal, wait, pause:moving ? pauseMs : 0, writingPause:moving && creation.active ? pauseMs / 2 : 0, track, chunk, projectCreation: creation.active && !!tableStage?.active, templateCreation: creation.active ? templateMode : false, meetingCreation: creation.active && meetingMode, proposalCreation: creation.active && proposalMode,
          onSelect: async (number, target, step, currentTexts) => {
            if (selectionWork) return;
            overlay.say(labels.step(number + 1, total));
            if (creation.active && (tableStage?.active || templateMode) && step.op === 'setText' && isTemplateBlank(target?.text)) {
              const find=paragraphRange(document,PROJECT_BLANK,target.block,0,1,true,true);
              await overlay.select(key(number), () => {
                const base=find(), element=findParagraphElements(document,PROJECT_BLANK,target.block,true)?.[0];
                if (!base || !element) return base;
                const font=document.defaultView.getComputedStyle(element), scale=element.getScreenCTM()?.a || 1;
                const canvas=document.createElement('canvas'), ctx=canvas.getContext('2d');
                ctx.font=`${font.fontWeight} ${parseFloat(font.fontSize)*scale}px ${font.fontFamily}`;
                let lines=1, width=0;
                for (const char of Array.from(step.text)) {
                  const advance=ctx.measureText(char).width;
                  if (char === '\n' || (width && width+advance > base.width)) { lines++; width=0; }
                  if (char !== '\n') width+=advance;
                }
                const height=base.rects[0].height;
                const rects=Array.from({length:Math.min(lines,30)},(_,i)=>({left:base.left,top:base.top+i*height,width:base.width,height}));
                return {...unionRects(rects),rects,clip:base.clip,page:base.page};
              }, {label:agentLabel(locale,target.block === 0 ? 'title' : 'body'),selecting:agentLabel(locale,'selecting'),writer:true,edge:true});
              return;
            }
            if(chatBatch && step.op==='setText' && target?.text){
              const original=step.find ?? target.text, start=target.start ?? 0;
              const end = start + Array.from(original).length;
              previousChanges.set(number, previousWriterChange(chatBatch.changes,currentTexts(),{paragraphs:[target.index,target.index],start,end},original));
              const work = chatBatch.claim({ kind:'docx', original, anchor:{ paragraphs:[target.index,target.index], start, end } })
                ?? startWriterWork({overlay,locale},locate(target, true),true);
              batchFrames.set(number,work);
              await work.ready;await overlay.phase(work.key,'clear',160);await overlay.phase(work.key,'write');
              if (!moving) return;
              const targetText=step.replace ?? step.text;
              const staged=await previewWriterPrefixes(overlay,work.key,textChunks(targetText),signal);
              if (staged.reserved) overlay.preview(work.key, undefined);
              // A partial replacement writes its whole SDK paragraph in one operation.
              return step.find === undefined ? staged : {parts:[step.text]};
            }
            if (creation.active && target?.insert) {
              const before = locate({...target,insert:undefined},true);
              await overlay.select(key(number),()=>{
                const found=before();
                const opening=writerOpening(document,found?.page);
                if (!found) return creationBands(opening);
                const line=found.rects?.at(target.insert==='before'?0:-1);
                return creationBands({left:opening?.left ?? found.left,top:target.insert==='before'?found.top:found.top+found.height,
                  width:opening?.width ?? found.width,height:line?.height ?? found.height,clip:found.clip});
              },{label:agentLabel(locale,docxStatus(step)),selecting:agentLabel(locale,'selecting'),writer:true,edge:true});
              return;
            }
            await overlay.select(key(number), locate(target), { label: agentLabel(locale, docxStatus(step)), selecting: agentLabel(locale, 'selecting'), edge: true });
          },
          onWrite: followWritten,
          onDone: async number => {
            const tableAt=doc.paragraphs.findIndex(p=>p.table), step=steps[number];
            if (tableStage?.active && !tableStage.framed && tableAt > 0 && step.op === 'setText' && step.target?.paragraph === tableAt-1 && step.text.trim()) await tableStage.activate();
          },
          onFinish: number => {
            if(writtenChanges.has(number)){chatBatch.record(writtenChanges.get(number),previousChanges.get(number));writtenChanges.delete(number);}
            const work = batchFrames.get(number);
            if (work) {
              overlay.holdParagraphCursor(work.key);
              work.dispose();
              batchFrames.delete(number);
            } else if (!selectionWork) {
              const slot = key(number);
              if (['insertParagraph', 'setText'].includes(steps[number].op)) overlay.holdParagraphCursor(slot);
              else overlay.release(slot);
              owners.delete(number);
            }
          },
        });
        if (selectionWork) overlay.holdParagraphCursor(selectionWork.key);
        const refs = outcome.refs;
        const after = await inspectDocument(word, document);
        const tableAt = after.paragraphs.findIndex(p => p.table);
        const preceding = after.paragraphs[tableAt-1]?.text;
        if (tableStage?.active && !tableStage.framed && tableAt > 0 && !isTemplateBlank(preceding) && preceding?.trim()) await tableStage.activate();
        return { applied: steps.length, ...(Object.keys(refs).length && { refs }), readToken: reads.remember(after), ...(outcome.styleFallback && { styleFallback: outcome.styleFallback }) };
      } catch (error) {
        const code = signal?.aborted ? 'cancelled' : PAGE_CODES.has(error?.code) ? error.code : 'editor-failed';
        const reason = signal?.aborted ? 'Stopped by the user.' : (error?.message ?? String(error));
        // Every paragraph still half written is named, and steps finished after the first unfinished one too.
        const halves = (track.pending ?? []).filter(entry => entry.parts.length).map(entry => entry.number);
        const own = Number.isInteger(track.partial) && !track.complete?.has(track.partial) && !halves.includes(track.partial) ? [track.partial] : [];
        const partial = [...halves, ...own].sort((a, b) => a - b);
        const extra = [...(track.complete ?? [])].filter(number => number >= (track.done ?? 0)).sort((a, b) => a - b);
        throw agentError(code, `${reason} ${docxProgress(track.done ?? 0, total, refsNow(track), partial, extra)}`);
      } finally {
        for(const frame of batchFrames.values())frame.dispose();
        // Writer reports local commits to the change listener as it makes them; settle input and let any late
        // report through while the changes still count as the AI's (spec 261105 §1.3).
        try { await editor.prepare?.(); } catch {}
        await wait(100);
        finishWrite();
        overlay.busy(false);
      }
    },
    async undo() { throw agentError('unsupported-operation', 'Word edits cannot be undone by the AI (there is no rollback); change the text back with another document_edit.'); },
    dispose() { overlay.dispose(); },
  };
  // Selection rewrite (spec 261109-selection-ai-edit): read the user's selection, replace it, take it back.
  Object.assign(agent, docSelectionMethods({ Word, document, reads, overlay, locale, animate, locked, applying: () => applying, apply: (...args) => agent.apply(...args) }));
  return agent;
}
