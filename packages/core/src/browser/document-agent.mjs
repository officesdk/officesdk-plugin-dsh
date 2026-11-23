import { DOCX_ALIGN, DOCX_LIVE_MARK, DOCX_OPS, DOCX_OUTLINE_STYLES, DOCX_STYLES, DOCX_STYLE_FALLBACK, checkDocxToken, docxProgress, docxReadToken, validateDocxOps } from '../agent/docx.mjs';
import { DOC_LABELS, createDocOverlay } from './doc-overlay.mjs';
import { agentLabel, docxStatus } from './agent-status.mjs';
import { textChunks } from './selection.mjs';
import { docSelectionMethods } from './selection-edit.mjs';

/**
 * AI editing of a Word document on a Kit page (spec 261105 §3.3), through office-kit's public `editor.api.writer`
 * (Word JS API). Each operation is one `Word.run`; the user sees the document change step by step. There is no
 * rollback (office-kit K3): when a step fails or the user stops, the steps already done stay in the editor and the
 * error says how far it got. Changes reach Kit through the page's change queue like a hand edit, but while `applying`
 * they are marked as the AI's, so they wait for a save instead of autosaving.
 */
const agentError = (code, message) => Object.assign(new Error(message), { code });
/** Codes the page passes on as they are; anything else from the editor is `editor-failed`. */
const PAGE_CODES = new Set(['invalid-edit', 'source-changed', 'cancelled', 'busy', 'read-only', 'unsupported-operation']);

/** The document as the model reads it: paragraphs (text, style, list, inside a table) and tables. */
export async function inspectDocument(Word) {
  let doc;
  await Word.run(async ctx => {
    const paragraphs = ctx.document.body.paragraphs;
    paragraphs.load('items/text,items/style,items/isListItem,items/tableNestingLevel');
    const tables = ctx.document.body.tables;
    tables.load('items/rowCount,items/values');
    await ctx.sync();
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

const normalize = text => String(text ?? '').replace(/\s+/g, '');

/**
 * Where paragraph `text` is drawn now: Writer draws each line as SVG `text[data-writer-document-content]` whose
 * `data-paint-key` is `fragment:<story>/<block>/…`. Lines of one block are joined and compared with the text; of
 * several matches the one whose block number is nearest `near` wins. Read only; undefined when the paragraph is not
 * drawn (pages are drawn near the view only) or has no text. Depends on Writer's DOM (office-kit K4 has no API).
 */
export function findParagraphElements(document, text, near = 0) {
  const wanted = normalize(text);
  if (!wanted) return undefined;
  const blocks = new Map();
  for (const element of document.querySelectorAll('text[data-writer-document-content]')) {
    const match = /^fragment:(\d+)\/(\d+)\//.exec(element.getAttribute('data-paint-key') ?? '');
    if (!match || match[1] !== '0') continue;
    const block = Number(match[2]);
    const entry = blocks.get(block) ?? { block, elements: [], text: '' };
    entry.elements.push(element);
    entry.text += element.textContent ?? '';
    blocks.set(block, entry);
  }
  let best;
  for (const entry of blocks.values()) {
    if (normalize(entry.text) !== wanted) continue;
    if (!best || Math.abs(entry.block - near) < Math.abs(best.block - near)) best = entry;
  }
  return best?.elements;
}

/** The screen box of `elements` (their union), clipped to Writer's scrolling view when it is there. */
export function boxOf(document, elements) {
  let left = Infinity, top = Infinity, right = -Infinity, bottom = -Infinity;
  for (const element of elements ?? []) {
    if (!element.isConnected) return undefined;
    const rect = element.getBoundingClientRect();
    if (!rect.width && !rect.height) continue;
    left = Math.min(left, rect.left); top = Math.min(top, rect.top); right = Math.max(right, rect.right ?? rect.left + rect.width); bottom = Math.max(bottom, rect.bottom ?? rect.top + rect.height);
  }
  if (!Number.isFinite(left)) return undefined;
  const scroller = document.querySelector('.writer-document-preview__scroll')?.getBoundingClientRect();
  return { left, top, width: right - left, height: bottom - top, ...(scroller && { clip: { left: scroller.left, top: scroller.top, right: scroller.right ?? scroller.left + scroller.width, bottom: scroller.bottom ?? scroller.top + scroller.height } }) };
}

/**
 * Where a new paragraph will go before it exists (spec §7.6 11): one line below (or above) the paragraph it is inserted
 * next to, as wide as that paragraph. Undefined when that paragraph is not drawn.
 */
export function placeLocator(document, { text, index, insert }) {
  let elements, misses = 0;
  return () => {
    if (!elements || elements.some(element => !element.isConnected)) {
      if (elements === null && misses++ % 8) return undefined;
      elements = findParagraphElements(document, text, index) ?? null;
    }
    const box = elements && boxOf(document, elements);
    if (!box) return undefined;
    const line = Math.max(12, Math.min(40, elements[0].getBoundingClientRect().height || 18));
    return { ...box, top: insert === 'before' ? box.top - line - 4 : box.top + box.height + 4, height: line };
  };
}

/** Follows paragraph `text` near block `near` every frame; finds its elements again when Writer redraws them. */
export function paragraphLocator(document, text, near) {
  let elements, misses = 0;
  return () => {
    if (!elements || elements.some(element => !element.isConnected)) {
      // Looking again costs a query over the drawn lines: at most every few frames while it is not found.
      if (elements === null && misses++ % 8) return undefined;
      elements = findParagraphElements(document, text, near) ?? null;
    }
    return elements ? boxOf(document, elements) : undefined;
  };
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
export async function runDocxSteps(Word, steps, doc, { signal, onStep = () => {}, onDone = () => {}, onSelect = async () => {}, onWrite = () => {}, onFinish = () => {}, chunk = text => [text], tracks = 3, wait = () => Promise.resolve(), pause = 0, track = {} } = {}) {
  const order = doc.paragraphs.map((_, number) => `#${number}`);
  const tableOrder = doc.tables.map((_, number) => `#${number}`);
  const texts = new Map(doc.paragraphs.map((paragraph, number) => [`#${number}`, paragraph.text]));
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
        const text = entry.parts.shift();
        const at = order.indexOf(entry.key);
        if (at < 0) throw agentError('editor-failed', 'A paragraph being written is no longer in the document.');
        // Each part adds to the end of the paragraph, so the formatting set when it was inserted stays.
        const added = text.slice((texts.get(entry.key) ?? '').length);
        await Word.run(async ctx => { (await load(ctx))[at].insertText(added, 'End'); await ctx.sync(); });
        texts.set(entry.key, text);
        onWrite(entry.number, { index: at, text });
        if (!entry.parts.length) await cursorTo(at);
      }
    }
    for (const entry of pending.splice(0)) finished(entry.number);
    track.partial = undefined;
  };
  /** What a step's agent selects before writing: an existing paragraph, or the place a new one will go. */
  const selectTarget = step => {
    if (step.op === 'insertParagraph' || step.op === 'insertPageBreak' || step.op === 'insertTable') {
      const where = step.where;
      const index = where.at === 'start' ? 0 : where.at === 'end' ? order.length - 1 : indexOf(where.after ?? where.before);
      return { insert: where.before || where.at === 'start' ? 'before' : 'after', index, text: texts.get(order[index]) ?? '' };
    }
    if (step.target) { const index = indexOf(step.target); return { index, text: texts.get(order[index]) ?? '' }; }
    return undefined;
  };
  for (const [number, step] of steps.entries()) {
    if (signal?.aborted) { track.partial = pending.find(item => item.parts.length)?.number; throw agentError('cancelled', 'Stopped by the user.'); }
    // Anything but another new paragraph, or a full set of tracks, first finishes the paragraphs being written.
    if (pending.length && (step.op !== 'insertParagraph' || pending.length >= tracks)) await flush();
    await onSelect(number, selectTarget(step), step);
    // A stop while the target was being selected writes nothing.
    if (signal?.aborted) { track.partial = pending.find(item => item.parts.length)?.number; throw agentError('cancelled', 'Stopped by the user.'); }
    let focus, grouped = false;
    switch (step.op) {
      case 'insertParagraph': {
        let at;
        const parts = chunk(step.text);
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
        onWrite(number, { index: at, text: parts[0] });
        if (parts.length > 1) { pending.push({ key: order[at], parts: parts.slice(1), number }); grouped = true; }
        focus = { index: at, text: parts[0], tone: 'add' };
        break;
      }
      case 'setText': {
        const at = indexOf(step.target);
        // A whole rewrite grows in parts like a new paragraph (in this step, not on a track); a replaceText (`find`)
        // changes only part of the paragraph and is written once: growing it would first wipe the rest of the text.
        const parts = step.find === undefined ? chunk(step.text) : [step.text];
        for (const [part, text] of parts.entries()) {
          if (part) {
            if (signal?.aborted) throw agentError('cancelled', 'Stopped by the user.');
            await wait(pause / 2);
            if (signal?.aborted) throw agentError('cancelled', 'Stopped by the user.');
          }
          const added = part ? text.slice(parts[part - 1].length) : text;
          await Word.run(async ctx => { (await load(ctx))[at].insertText(added, part ? 'End' : 'Replace'); await ctx.sync(); });
          texts.set(order[at], text);
          // From the first part on, a stop leaves this paragraph half rewritten.
          if (part < parts.length - 1) track.partial = number;
          if (parts.length > 1) onWrite(number, { index: at, text });
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
          tableOrder.splice(tablesBefore(after, at), 0, step.ref ? `@${step.ref}` : `+${number}`);
          if (step.ref) refs[step.ref] = 'table';
        });
        // The frame goes round the first cell's text (each cell is its own drawn block).
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
    onDone(number, focus);
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
  const labels = DOC_LABELS[locale] ?? DOC_LABELS.zh;
  const overlay = createDocOverlay(document, { animate, wait });
  let applying = false;
  /** The lock and overlay of apply, for the selection rewrite's in-paragraph writes (selection-edit.mjs). */
  const locked = async run => {
    if (applying) throw agentError('busy', 'Another edit is still running.');
    applying = true;
    overlay.busy(true, labels.editing);
    try { return await run(); }
    finally { try { await editor.prepare?.(); } catch {} await wait(100); applying = false; overlay.busy(false); }
  };
  const agent = {
    /** True while an AI edit is writing: its changes are not the user's. */
    get applying() { return applying; },
    handOff() {},
    inspect: () => inspectDocument(Word()),
    async apply({ ops, readToken }, signal) {
      if (applying) throw agentError('busy', 'Another edit is still running.');
      applying = true;
      overlay.busy(true, labels.editing);
      let total = Array.isArray(ops) ? ops.length : 0;
      const track = { done: 0 };
      try {
        const word = Word();
        const doc = await inspectDocument(word);
        // Checked here, after the approval card: the user may have edited while it was open.
        checkDocxToken(ops, doc, readToken);
        const { steps } = validateDocxOps(ops, doc, { enabled });
        total = steps.length;
        // v1.1: each step's agent (one of three tracks) selects first, then writes; new paragraphs grow in parts.
        // A step keeps its track until it finishes; a new step takes a free one.
        const owners = new Map();
        const key = number => {
          if (!owners.has(number)) owners.set(number, ['t0', 't1', 't2'].find(name => ![...owners.values()].includes(name)) ?? `t${number % 3}`);
          return owners.get(number);
        };
        const outcome = await runDocxSteps(word, steps, doc, {
          signal, wait, pause: animate() ? pauseMs : 0, track, chunk: animate() ? textChunks : text => [text],
          onSelect: async (number, target, step) => {
            overlay.say(labels.step(number + 1, total));
            const find = !target ? undefined : step.op === 'insertParagraph' ? placeLocator(document, target) : paragraphLocator(document, target.text, target.index);
            await overlay.select(key(number), find ?? (() => undefined), { label: agentLabel(locale, docxStatus(step)), selecting: agentLabel(locale, 'selecting'), edge: false });
          },
          onWrite: (number, written) => overlay.point(key(number), paragraphLocator(document, written.text, written.index)),
          onFinish: number => { overlay.release(key(number)); owners.delete(number); },
        });
        const refs = outcome.refs;
        const after = await inspectDocument(word);
        return { applied: steps.length, ...(Object.keys(refs).length && { refs }), readToken: docxReadToken(after), ...(outcome.styleFallback && { styleFallback: outcome.styleFallback }) };
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
        // Writer reports local commits to the change listener as it makes them; settle input and let any late
        // report through while the changes still count as the AI's (spec 261105 §1.3).
        try { await editor.prepare?.(); } catch {}
        await wait(100);
        applying = false;
        overlay.busy(false);
      }
    },
    async undo() { throw agentError('unsupported-operation', 'Word edits cannot be undone by the AI (there is no rollback); change the text back with another document_edit.'); },
    dispose() { overlay.dispose(); },
  };
  // Selection rewrite (spec 261109-selection-ai-edit): read the user's selection, replace it, take it back.
  Object.assign(agent, docSelectionMethods({ Word, document, overlay, locale, animate, locked, applying: () => applying, apply: (args, signal) => agent.apply(args, signal) }));
  return agent;
}
