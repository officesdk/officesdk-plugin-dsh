import { rebaseChangeText, sameChangeAnchor } from './change-history.mjs';
import {chatTarget, needsSuiteTools} from '../agent/chat-edit-target.mjs';
import { selectionHighlightBounds } from './writer-selection-highlight.mjs';
import { SELECTION_LIMITS, cpLength, planDocRevert, selectedText, planDocRewrite, planSheetRevert, planSheetRewrite, sheetAddress } from '../agent/selection.mjs';
import { checkDocxToken, docxReadToken } from '../agent/docx.mjs';
import { agentLabel } from './agent-status.mjs';
import { inspectDocument } from './document-agent.mjs';
import { bodyBlock, locateWriterRange, paragraphRange, rememberWriterIndents, writerRangeLocator } from './writer-geometry.mjs';
export { unionRects, locateWriterRange } from './writer-geometry.mjs';
import { textChunks } from './selection.mjs';
import { inspectWorkbook, keptStyle } from './sheet-agent.mjs';
import { parseRange, validateXlsxOps } from '../agent/xlsx.mjs';

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
export async function readDocSelection(Word, document) {
  let result;
  await Word.run(async ctx => {
    const selection = ctx.document.getSelection();
    selection.load('text,isEmpty');
    const body = ctx.document.body.paragraphs;
    body.load('items/text,items/tableNestingLevel' + (document ? ',items/leftIndent,items/rightIndent' : ''));
    const own = selection.paragraphs;
    own.load('items/text');
    await ctx.sync();
    if (selection.isEmpty || !String(selection.text ?? '').replace(/[\r\n]/g, '') || !own.items.length) return;
    rememberWriterIndents(document, body.items);
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
    const block = bodyBlock(body.items, from, p => Number(p.tableNestingLevel) > 0);
    result = { kind: 'docx', anchor: { paragraphs: [from, to], start, end }, original,
      view: { first: from, texts: texts.slice(from, to + 1), blocks: parts.map((_, i) => block + i) } };
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
export async function replaceInParagraph(Word, { index, start, length, text, expected, wholeParagraph = false, signal, chunk = value => [value], onWrite = () => {} }) {
  let current = length;
  const parts = text ? chunk(text) : [''];
  for (const part of parts) {
    if (signal?.aborted) throw agentError('cancelled', 'Stopped by the user.');
    await Word.run(async ctx => {
      const body = ctx.document.body.paragraphs;
      body.load(expected === undefined ? 'items' : 'items/text');
      await ctx.sync();
      const paragraph = body.items[index];
      if (expected !== undefined) {
        const live = paragraph?.text;
        if (typeof live !== 'string' || (wholeParagraph ? live : Array.from(live).slice(start, start + current).join('')) !== expected)
          throw agentError('source-changed', 'The target text changed before it could be written.');
      }
      const whole = paragraph.getRange('Whole');
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
    if (expected !== undefined) expected = part;
    current = cpLength(part);
    await onWrite(part);
  }
}

/**
 * Rewrites (or takes back) a Word selection on the agent's terms: `agent` gives `Word()`, the overlay, `apply` (a
 * document_edit batch with its animation) and the labels. `plan` is planDocRewrite's result.
 */
export function startWriterWork({overlay, locale}, find, persistent = false, awaitingResult = false) {
  const key = Symbol('writer-selection');
  const ready = (async () => {
    await overlay.select(key, find, { label:agentLabel(locale,'edit'), selecting:agentLabel(locale,'selecting'), edge:true, writer:true, persistent });
    await overlay.phase?.(key, 'edge', 200);
    // A primed track breathes until its writer arrives. Give even a cached reply one visible placeholder frame.
    await overlay.phase?.(key, 'placeholder', awaitingResult ? 32 : 1000);
  })();
  return { key, ready, point:locate=>overlay.point(key,locate), dispose: () => overlay.drop(key) };
}

function writingPrefixes(text) {
  return textChunks(text,{min:1,max:Math.min(36,Math.max(1,Math.ceil(cpLength(text)/2)))});
}

// Display fitting prefixes over the unchanged document. At overflow the real editor takes over
// with the complete prefix in one write. A fitting replacement commits only after its last frame.
export async function previewWriterPrefixes(overlay, key, parts, signal) {
  if (!overlay.preview) return {parts};
  for (let i=0; i<parts.length; i++) {
    if (signal?.aborted) throw agentError('cancelled', 'Stopped by the user.');
    if (!overlay.preview(key, parts[i])) {
      overlay.preview(key, undefined);
      return {parts:parts.slice(i)};
    }
    await overlay.phase(key, 'write', 36);
  }
  return {parts:parts.slice(-1), reserved:true};
}

export async function runDocPlan({ Word, document, overlay, locale, apply, animate }, plan, doc, signal, prepared) {
  if (plan.mode === 'unchanged') return;
  if (plan.mode === 'paragraphs') {
    const work = prepared ?? (plan.source && startWriterWork({overlay,locale},writerRangeLocator(document,plan.source,true),true));
    try {
      let staged;
      if (work) {
        await work.ready;
        await overlay.phase(work.key, 'clear', 160);
        await overlay.phase(work.key, 'write');
        if (animate) staged=await previewWriterPrefixes(overlay,work.key,writingPrefixes(plan.afterText),signal);
        if (staged?.reserved) overlay.preview(work.key, undefined);
      }
      await apply({ ops: plan.ops, readToken: docxReadToken(doc) }, signal, work && {
        key:work.key, after:plan.after, reserved:staged?.reserved, previewLines:staged?.parts[0]?.split('\n')
      });
    } finally { if (!prepared) work?.dispose(); }
    return;
  }
  const text = doc.paragraphs[plan.paragraph]?.text ?? '';
  const block = bodyBlock(doc.paragraphs, plan.paragraph);
  const original=paragraphRange(document, text, block, plan.start, plan.start + plan.length, true, true);
  const work = prepared ?? startWriterWork({overlay, locale}, original);
  const key = work.key;
  try {
    await work.ready;
    await overlay.phase?.(key,'clear',160);
    if (signal?.aborted) throw agentError('cancelled', 'Stopped by the user.');
    const head = Array.from(text).slice(0, plan.start).join(''), tail = Array.from(text).slice(plan.start + plan.length).join('');
    const initial=original(), page=initial?.page, base=page?.getBoundingClientRect();
    let live, previous;
    const find=()=>{
      const next=live?.(), current=page?.getBoundingClientRect();
      let held=initial;
      if (initial && current?.width && base?.width) {
        const scale=current.width/base.width;
        const move=r=>({left:current.left+(r.left-base.left)*scale,top:current.top+(r.top-base.top)*scale,width:r.width*scale,height:r.height*scale});
        held={...move(initial),rects:initial.rects.map(move),clip:page.closest('.writer-document-preview__scroll')?.getBoundingClientRect()};
      }
      if (!held) return next ?? previous;
      if (next) previous=next;
      const drawn=next ?? previous;
      const end=held.rects?.at(-1), tip=drawn?.rects?.at(-1);
      if (drawn && (drawn.height>held.height+1 || drawn.width>held.width+1 || (tip && end && tip.left+tip.width>end.left+end.width+1 && tip.top>=end.top-1))) return drawn;
      return held;
    };
    await overlay.phase?.(key,'write');
    const parts=writingPrefixes(plan.text);
    const staged=animate ? await previewWriterPrefixes(overlay,key,parts,signal) : {parts:[plan.text]};
    if (!staged.reserved) overlay.point(key,find);
    let written=null;
    try {
      await replaceInParagraph(Word,{index:plan.paragraph,start:plan.start,length:plan.length,text:plan.text,expected:Array.from(text).slice(plan.start,plan.start+plan.length).join(''),wholeParagraph:!head && !tail,signal,chunk:()=>staged.parts,
        onWrite:async part=>{
          written=part;
          live=paragraphRange(document,head+part+tail,block,plan.start,plan.start+cpLength(part),true,true);
          if(animate && part!==plan.text) await overlay.phase?.(key,'write',36);
        }});
      overlay.holdParagraphCursor(key, live);
    } catch(error) {
      // Only restore the bytes this operation actually wrote; the source is locked during these writes.
      if(written!==null) await replaceInParagraph(Word,{index:plan.paragraph,start:plan.start,length:cpLength(written),expected:written,wholeParagraph:!head && !tail,text:Array.from(text).slice(plan.start,plan.start+plan.length).join('')});
      throw error;
    }
  } finally { if (!prepared) overlay.release(key); }
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
  return sheetRangeReference(worksheet,range);
}

function sheetRangeReference(worksheet,range) {
  if (!range) return undefined;
  if (range.type !== 'cells') return { unsupported: 'not-cells' };
  if (range.rowCount * range.columnCount > SELECTION_LIMITS.maxCells) return { unsupported: 'too-many' };
  const data = plain(range.getData()) ?? [];
  if (data.some(row => (row ?? []).some(cell => cell && Object.keys(cell).some(key => /merge/i.test(key) && cell[key])))) return { unsupported: 'merged' };
  const anchor = { sheet: String(worksheet.name), row: range.row, column: range.column, rowCount: range.rowCount, columnCount: range.columnCount };
  const cells = Array.from({ length: anchor.rowCount }, (_, r) => Array.from({ length: anchor.columnCount }, (_, c) => cellOf(data[r]?.[c])));
  let box;
  try { box = range.getBounding?.() ?? undefined; } catch {}
  return { kind: 'xlsx', anchor, cells, original:cells.map(row=>row.map(c=>c.formula ? '='+c.formula : c.text).join('\t')).join('\n'), ...(box && { box: plain(box) }) };
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
    const found = paragraphRange(document, text, view.first + offset)();
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
function resolveChangeAnchor(texts,revert,located) {
  const {anchor,original}=revert, captured=located?.view?.texts;
  const [first,last]=anchor.paragraphs;
  const fail=()=>{throw agentError('source-changed','这处文字的位置或内容已变化，无法安全定位。');};
  if (!captured) {
    if(last>=texts.length || selectedText(texts,anchor)!==original) fail();
    return anchor;
  }
  const matches=[];
  for(let i=0;i<=texts.length-captured.length;i++) if(captured.every((line,n)=>texts[i+n]===line)) matches.push(i);
  if(matches.length===1 || matches.includes(first)) {
    const at=matches.includes(first)?first:matches[0];
    return {...anchor,paragraphs:[at,at+last-first]};
  }
  // Edits elsewhere in the same paragraph may shift this range. Require a unique match and unchanged neighbouring text.
  if(first===last) {
    const shifted=rebaseChangeText(texts[first],original,captured[0],anchor);
    if (shifted) return shifted;
  }
  fail();
}

export function previousWriterChange(records, texts, anchor, original) {
  return records?.findLast(record => {
    if (record.kind !== 'docx' || record.revert?.original !== original) return false;
    try { return sameChangeAnchor(resolveChangeAnchor(texts,record.revert,record.located),anchor); }
    catch { return false; }
  })?.id;
}

// Resolve the location against the same revision the model read. Writer uses code-point offsets.
export function modelEditRange(doc, { paragraph, find, readToken }, check = checkDocxToken) {
  if (!Number.isInteger(paragraph) || paragraph < 0) throw agentError('invalid-edit', 'Choose a paragraph index from document_read.');
  check([{ op: find === undefined ? 'setText' : 'replaceText', paragraph, find }], doc, readToken);
  const entry = doc.paragraphs[paragraph];
  if (!entry?.text || entry.table) throw agentError('invalid-edit', 'Choose a nonempty body paragraph.');
  const original = find === undefined ? entry.text : find;
  if (typeof original !== 'string' || !original) throw agentError('invalid-edit', 'find must be nonempty text.');
  const index = entry.text.indexOf(original);
  if (index < 0 || index !== entry.text.lastIndexOf(original)) throw agentError('invalid-edit', 'The target must occur exactly once in this paragraph.');
  const start = cpLength(entry.text.slice(0, index));
  return { kind: 'docx', original, anchor: { paragraphs: [paragraph, paragraph], start, end: start + cpLength(original) },
    view: { first: paragraph, texts: [entry.text], blocks: [bodyBlock(doc.paragraphs, paragraph)] } };
}

export function docSelectionMethods(env) {
  const perform = async ({ anchor, original, text }, signal, revert=false, prepared, captured, records) => {
    const word = env.Word();
    const doc = await inspectDocument(word, env.document);
    if (captured?.view) anchor=resolveChangeAnchor(doc.paragraphs.map(p=>p.text),{anchor:captured.anchor,original},{view:captured.view});
    env.reads?.remember(doc);
    const plan = (revert?planDocRevert:planDocRewrite)(doc.paragraphs, anchor, original, text);
    if (plan.mode === 'paragraphs' && !prepared) {
      const [first,last] = anchor.paragraphs;
      plan.source={anchor,view:{first,texts:doc.paragraphs.slice(first,last+1).map(p=>p.text),
        blocks:Array.from({length:last-first+1},(_,offset)=>bodyBlock(doc.paragraphs,first+offset))}};
    }
    const previousId = previousWriterChange(records,doc.paragraphs.map(p=>p.text),anchor,original);
    await runDocPlan({ ...env, Word: word, animate: env.animate() }, plan, doc, signal, prepared);
    const now = await inspectDocument(word, env.document);
    const texts = now.paragraphs.map(paragraph => paragraph.text);
    const [first, last] = plan.after.paragraphs;
    if (!prepared) await selectDocRange(word, plan.after, texts);
    return { plan, now, previousId, view: { first, texts: texts.slice(first, last + 1) } };
  };
  // Inspect and resolve offsets after taking the writer lock. Multiline plans acquire
  // that same lock inside apply(), where their read token is checked before writing.
  const run = (args, signal, ...rest) => {
    const inline = args.anchor.paragraphs[0] === args.anchor.paragraphs[1] && !/[\r\n]/.test(args.text);
    const performEdit = () => perform(args, signal, ...rest);
    return inline ? env.locked(performEdit, signal) : performEdit();
  };
  return {
    modelEditRange: (doc, target) => modelEditRange(doc, target, env.reads?.check),
    async chatSelection(instruction,selection) {
      const {paragraphs}=await inspectDocument(env.Word(), env.document);
      let ref=chatTarget(instruction,paragraphs,selection);
      if(!ref && !selection && /(?:这[一]?段|当前段落|本段|this paragraph|current paragraph)/i.test(instruction)){
        let index=-1;
        await env.Word().run(async ctx=>{
          const body=ctx.document.body.paragraphs, here=ctx.document.getSelection().paragraphs;
          body.load('items');here.load('items');await ctx.sync();
          if(here.items.length!==1)return;
          const range=here.items[0].getRange('Whole');
          const matches=body.items.map(p=>p.getRange('Whole').compareLocationWith(range));
          await ctx.sync();index=matches.findIndex(item=>item.value==='Equal');
        });
        const text=paragraphs[index]?.text;
        if(text && !paragraphs[index].table)ref={kind:'docx',original:text,anchor:{paragraphs:[index,index],start:0,end:cpLength(text)},view:{first:index,texts:[text]}};
      }
      if (ref) ref.view.blocks=ref.view.texts.map((_,i)=>bodyBlock(paragraphs,ref.view.first+i));
      return ref;
    },
    beginSelectionWork: ref => startWriterWork(env, writerRangeLocator(env.document,ref,true), true, true),
    readSelection: () => env.applying() ? Promise.resolve(undefined) : readDocSelection(env.Word(), env.document),
    locate: target => locateWriterRange(env.document, target),
    locateActiveSelection: target => selectionHighlightBounds(env.document,locateWriterRange(env.document,target)),
    async describeSelection(args) {
      const doc = await inspectDocument(env.Word(), env.document);
      const [first,last] = args.anchor.paragraphs;
      return { ...args, view:{ first, texts:doc.paragraphs.slice(first,last+1).map(p=>p.text) } };
    },
    async rewriteSelection(args, signal, prepared, captured, records) {
      const { plan, now, view, previousId } = await run(args, signal, false, prepared, captured, records);
      return {
        previousId,
        reply: { applied: plan.mode !== 'unchanged', mode: plan.mode, before: args.original, after: plan.afterText, readToken: env.reads?.remember(now) ?? docxReadToken(now) },
        revert: { anchor: plan.after, original: plan.afterText, text: args.original },
        located: { view, anchor: plan.after },
      };
    },
    async changeLocations(records, workBands=false) {
      const {paragraphs}=await inspectDocument(env.Word(), env.document);
      const texts=paragraphs.map(p=>p.text), locations=new Map();
      for (const record of records) {
        if (record.kind!=='docx' || typeof record.revert?.original!=='string') continue;
        try {
          const anchor=resolveChangeAnchor(texts,record.revert,record.located);
          const [first,last]=anchor.paragraphs;
          const view={first,texts:texts.slice(first,last+1),blocks:Array.from({length:last-first+1},(_,offset)=>bodyBlock(paragraphs,first+offset))};
          locations.set(record.id,writerRangeLocator(env.document,{anchor,view},workBands,!workBands));
        } catch { /* A later edit can replace the text this card described. */ }
      }
      return locations;
    },
    async selectChange({revert,located}) {
      const doc=await inspectDocument(env.Word(), env.document), texts=doc.paragraphs.map(p=>p.text);
      const anchor=resolveChangeAnchor(texts,revert,located);
      await selectDocRange(env.Word(),anchor,texts);
      const [first,last]=anchor.paragraphs;
      if (cpLength(revert.original)>SELECTION_LIMITS.maxText || last-first+1>SELECTION_LIMITS.maxParagraphs) return {unsupported:'too-long'};
      return {kind:'docx',anchor,original:revert.original,view:{first,texts:texts.slice(first,last+1)}};
    },
    async revertSelection(revert, signal, located) {
      const doc=await inspectDocument(env.Word(), env.document);
      const anchor=resolveChangeAnchor(doc.paragraphs.map(p=>p.text),revert,located);
      await run({...revert,anchor}, signal,true);
    },
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
  const write = async (anchor, step, signal, specials = [], local = false) => {
    const op = sheetOp(step);
    // Refused values must not cost the formulas: check the batch before anything is cleared.
    validateXlsxOps([op], inspectWorkbook(env.api()));
    const cleared = await env.locked(async () => clearFormulas(env.api(), anchor, step.values));
    try { await env.apply({ ops: [op] }, signal, local); }
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
        if (box && box.width > 0 && box.height > 0) {
          const rect = {left:box.left, top:box.top, width:box.width, height:box.height};
          return {...rect, rects:[rect]};
        }
        return undefined;
      } catch { return undefined; }
    }
    return target?.box;
  };
  return {
    readSelection: async () => env.applying() ? undefined : readSheetSelection(env.api()),
    locate,
    beginSelectionWork: ref => startWriterWork(env, () => locate(ref), true, true),
    async chatSelection(instruction, selection) {
      if (needsSuiteTools(instruction)) return;
      // Only use an explicit cell address on the named or active worksheet, or a selection supplied by the composer.
      const subject=instruction.split(/(?:改成|改为|替换为|换成|\bto\b|\bwith\b)/i)[0];
      const match=subject.match(/(?<![a-zA-Z0-9_])(?:(?:'([^']+)'|([\w.\u4e00-\u9fff]+))!)?(\$?[A-Z]{1,3}\$?[1-9]\d*(?::\$?[A-Z]{1,3}\$?[1-9]\d*)?)(?![a-zA-Z0-9_])/i);
      if (!match) return selection?.kind==='xlsx' ? selection : undefined;
      const grid=parseRange(match[3].replaceAll('$','').toUpperCase());
      if (!grid || grid.rowCount*grid.columnCount > SELECTION_LIMITS.maxCells) return;
      const api=env.api(), named=match[1] ?? match[2];
      const sheets=Array.from(api.workbook.getWorksheets());
      const sheetName=match[1] ?? named?.replace(/^(?:(?:请|把|将|修改|改写|更新|编辑|替换|表格|工作表)\s*)+/u,'');
      const sheet=named ? sheets.find(s=>s.name===named) ?? sheets.find(s=>s.name.toLowerCase()===sheetName.toLowerCase()) : api.activeSheet;
      if(sheet && grid.row+grid.rowCount<=1048576 && grid.column+grid.columnCount<=16384) return sheetRangeReference(sheet,sheet.getRange({type:'cells',...grid}));
    },
    /** The Sheet keeps the keyboard for its grid and takes focus back; the request box asks it to let go while open. */
    async focusEditor(on) { try { await env.api().setFocus?.(on); } catch {} },
    async rewriteSelection({ anchor, cells, values }, signal, prepared, _captured, records) {
      const work=prepared ?? (env.overlay && startWriterWork(env,()=>locate({anchor}),true));
      try {
        if (work) {
          await work.ready;
          await env.overlay.phase(work.key,'clear',160);
          await env.overlay.phase(work.key,'write');
        }
        if (signal?.aborted) throw agentError('cancelled','Stopped by the user.');
        const { cells: now } = readSheetCells(env.api(), anchor);
        const { step } = planSheetRewrite(now, anchor, cells, values);
        const previous = records?.findLast(r => {
          if (r.kind !== 'xlsx' || !sameChangeAnchor(r.revert?.anchor,anchor)) return false;
          try { planSheetRevert(now,anchor,r.revert.written,r.revert.cells); return true; }
          catch { return false; }
        });
        await write(anchor, step, signal, [], true);
        return {
          previousId:previous?.id,
          reply: { applied: true, range: `${anchor.sheet}!${sheetAddress(anchor)}`, before: cells.map(row => row.map(cell => cell.formula ? `=${cell.formula}` : cell.text)), after: step.values },
          revert: { anchor, written: step.values, cells },
          located: { anchor },
        };
      } finally { if (!prepared) work?.dispose(); }
    },
    changeLocations(records) {
      const locations=new Map();
      for (const r of records) {
        if (r.kind!=='xlsx' || !r.revert) continue;
        try {
          const {anchor,written,cells}=r.revert;
          planSheetRevert(readSheetCells(env.api(),anchor).cells,anchor,written,cells);
          locations.set(r.id,()=>locate({anchor}));
        } catch { /* Skip cells changed since the recorded edit. */ }
      }
      return locations;
    },
    async selectChange({revert:{anchor,written,cells}}) {
      const {cells:now,range,worksheet}=readSheetCells(env.api(),anchor);
      planSheetRevert(now,anchor,written,cells);
      worksheet.activate?.();range.select();
      return sheetRangeReference(worksheet,range);
    },
    async revertSelection({ anchor, written, cells }, signal) {
      const { cells: now } = readSheetCells(env.api(), anchor);
      const { step, specials } = planSheetRevert(now, anchor, written, cells);
      await write(anchor, step, signal, specials);
    },
  };
}
