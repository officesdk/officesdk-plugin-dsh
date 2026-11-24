import { writerWorkBands } from './writer-selection-highlight.mjs';
const normalize = text => String(text ?? '').replace(/\s+/g, '');

// The SDK reports paragraph indents in points; painted page coordinates use CSS pixels at 100%.
const indentsByDocument = new WeakMap();
export function rememberWriterIndents(document, paragraphs) {
  if (!document) return;
  const blocks = new Map();
  let block = -1, previousCell = false;
  for (const paragraph of paragraphs) {
    const cell = Number(paragraph.tableNestingLevel) > 0;
    if (!cell || !previousCell) block++;
    if (!cell) blocks.set(block, { left: Number(paragraph.leftIndent) || 0, right: Number(paragraph.rightIndent) || 0 });
    previousCell = cell;
  }
  indentsByDocument.set(document, blocks);
}

function fullParagraphBands(document, elements, text, index) {
  const pages = new Map();
  for (const element of elements) {
    const page = element.closest('svg');
    if (!pages.has(page)) pages.set(page, []);
    pages.get(page).push(element);
  }
  const bands = [];
  for (const [page, runs] of pages) {
    const rows = writerWorkBands(document, runs, lineRects(document, runs, text, 0, Infinity, true));
    if (!rows.length) continue;
    const drawn = unionRects(rows);
    let left = drawn.left, right = drawn.left + drawn.width;
    const start = page?.querySelector('path[data-paint-key$=":corner:top-left"]');
    const end = page?.querySelector('path[data-paint-key$=":corner:top-right"]');
    const scale = start?.getScreenCTM()?.a;
    const indent = indentsByDocument.get(document)?.get(index);
    if (start && end && scale > 0 && indent) {
      const pageLeft = start.getBoundingClientRect().right + scale / 2;
      const pageRight = end.getBoundingClientRect().left + scale / 2;
      const lo = pageLeft + indent.left * scale * 4 / 3;
      const hi = pageRight - indent.right * scale * 4 / 3;
      // A neighbouring column or floating text must stay visible. If this paragraph shares a
      // band, limit the expansion on that side to its own runs rather than spanning the gutter.
      const own = new Set(runs);
      const adjacent = [...page.querySelectorAll('text[data-writer-document-content]')]
        .filter(el => !own.has(el)).map(el => el.getBoundingClientRect())
        .filter(r => r.width && r.top < drawn.top + drawn.height && r.bottom > drawn.top);
      if (hi > lo) {
        // Hanging punctuation may paint into the margin. Only layout sets the paragraph frame.
        if (!adjacent.some(r => r.right <= left + scale && r.right > lo)) left = lo;
        if (!adjacent.some(r => r.left >= right - scale && r.left < hi)) right = hi;
      }
    }
    for (const row of rows) bands.push({ ...row, inkLeft: row.left, inkRight: row.left + row.width, left, width: right - left });
  }
  return bands;
}

/**
 * Where paragraph `text` is drawn now: Writer draws each line as SVG `text[data-writer-document-content]` whose
 * `data-paint-key` is `fragment:<story>/<block>/…`. Lines of one block are joined and compared with the text; of
 * several matches the one whose block number is nearest `near` wins. Read only; undefined when the paragraph is not
 * drawn (pages are drawn near the view only) or has no text. Depends on Writer's DOM (office-kit K4 has no API).
 */
export function findParagraphElements(document, text, near = 0, strict = false) {
  const wanted = normalize(text);
  if (!wanted) return undefined;
  const blocks = new Map();
  const painted = new Set();
  const selector = strict ? `text[data-writer-document-content][data-paint-key^="fragment:0/${near}/"]` : 'text[data-writer-document-content]';
  for (const element of document.querySelectorAll(selector)) {
    const key = element.getAttribute('data-paint-key') ?? '';
    const match = /^fragment:(\d+)\/(\d+)\//.exec(key);
    if (!match || match[1] !== '0') continue;
    // Whole-page zoom also mounts an invisible copy of each SVG fragment for measurements.
    // Only the painted copy supplies screen coordinates.
    const rect = element.getBoundingClientRect();
    if ((!rect.width && !rect.height) || painted.has(key)) continue;
    painted.add(key);
    const block = Number(match[2]);
    const entry = blocks.get(block) ?? { block, elements: [], text: '' };
    entry.elements.push(element);
    entry.text += element.textContent ?? '';
    blocks.set(block, entry);
  }
  let best;
  for (const entry of blocks.values()) {
    if (normalize(entry.text) !== wanted || (strict && entry.block !== near)) continue;
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

/** A collapsed range uses the adjacent glyph, including at a soft line wrap. Offsets in the model count code points. */
function lineRects(document, elements, text, start, end, forWrite = false, reviewInk = false) {
  const lines = [];
  let position = 0;
  const collapsed = start === end;
  const afterLast = collapsed && start === Array.from(text).length;
  const from = afterLast ? start - 1 : start;
  const to = collapsed ? from + 1 : end;
  for (const element of elements) {
    const walker = document.createTreeWalker(element, 4);
    for (let node = walker.nextNode(); node; node = walker.nextNode()) {
      const chars = Array.from(node.textContent);
      const lo = Math.max(0, from - position), hi = Math.min(chars.length, to - position);
      position += chars.length;
      if (hi <= lo) continue;
      const range = document.createRange();
      range.setStart(node, chars.slice(0, lo).join('').length);
      range.setEnd(node, chars.slice(0, hi).join('').length);
      for (const rect of range.getClientRects()) {
        if (!rect.width || !rect.height) continue;
        const left = collapsed && afterLast ? rect.right : rect.left;
        const width = collapsed ? 1 : rect.width;
        const prev = lines.at(-1);
        if (prev && (!reviewInk || prev.inkNode === node.parentElement) && Math.abs(prev.top - rect.top) < 0.75 && Math.abs(prev.height - rect.height) < 0.75 &&
            left >= prev.left && left <= prev.left + prev.width + 1) {
          prev.width = Math.max(prev.width, left + width - prev.left);
        } else lines.push({ left, top: rect.top, width, height: rect.height, ...(forWrite && { run:element }), ...(reviewInk && { inkNode:node.parentElement }) });
      }
    }
  }
  return lines;
}

export function unionRects(rects) {
  if (!rects.length) return undefined;
  const left = Math.min(...rects.map(r => r.left)), top = Math.min(...rects.map(r => r.top));
  const right = Math.max(...rects.map(r => r.left + r.width)), bottom = Math.max(...rects.map(r => r.top + r.height));
  const last = rects.at(-1);
  return { left, top, width: right-left, height: bottom-top, endX: last.left+last.width, endY: last.top };
}

// Cache only the text nodes. All screen measurements are read after the current scroll, zoom and layout.
export function paragraphRange(document, text, index, start = 0, end = Array.from(text).length, strict = false, workBands = false, reviewInk = false) {
  const chars = workBands && end > start ? Array.from(text) : undefined;
  const whole = chars && !chars.slice(0,start).join('').trim() && !chars.slice(end).join('').trim();
  let elements;
  return () => {
    if (!elements?.every(el => el.isConnected && el.getBoundingClientRect().height) ||
        elements.map(el => el.textContent).join('') !== text) {
      elements = findParagraphElements(document, text, index, strict);
    }
    if (!elements || elements.map(el => el.textContent).join('') !== text) return undefined;
    let rects;
    if (whole) {
      rects = fullParagraphBands(document, elements, text, index);
    } else {
      rects = lineRects(document, elements, text, start, end, workBands, reviewInk);
      if (workBands) rects = writerWorkBands(document, elements, rects);
    }
    const bounds = unionRects(rects);
    if (!bounds) return undefined;
    const clip = elements[0].closest('.writer-document-preview__scroll')?.getBoundingClientRect();
    return { ...bounds, rects, clip, page:elements[0].closest('svg') };
  };
}

export function writerRangeLocator(document, target, workBands = false, reviewInk = false) {
  const { view, anchor } = target ?? {};
  if (!view || !anchor) return () => undefined;
  const ranges = view.texts.map((text, offset) => {
    const index = view.first + offset;
    const from = index === anchor.paragraphs[0] ? anchor.start : 0;
    const to = index === anchor.paragraphs[1] ? anchor.end : Array.from(text).length;
    return paragraphRange(document, text, view.blocks?.[offset] ?? index, from, to, !!view.blocks, workBands, reviewInk);
  });
  return () => {
    const visible = ranges.map(find => find()).filter(Boolean);
    if (visible.length <= 1) return visible[0];
    const rects = visible.flatMap(range => range.rects);
    return { ...unionRects(rects), rects, clip:visible[0].clip };
  };
}

export function locateWriterRange(document, target) {
  return writerRangeLocator(document, target)();
}


// Writer's body paragraph collection includes the paragraphs inside a table. The painted body
// counts that table as one block. Keep the two indexes separate, including after live inserts.
export function bodyBlock(paragraphs, index, isCell = p => p.table) {
  let block = -1, previousCell = false;
  for (let i = 0; i <= index && i < paragraphs.length; i++) {
    const cell = isCell(paragraphs[i]);
    if (!cell || !previousCell) block++;
    previousCell = cell;
  }
  return block;
}

export function tableLocator(document, block, { row, column, rows } = {}) {
  return () => {
    const cells = new Map();
    const prefix = `table-cell:0/${block}/`;
    for (const el of document.querySelectorAll(`[data-paint-key^="${prefix}"]`)) {
      const match = /^(\d+)\/(\d+):/.exec(el.getAttribute('data-paint-key').slice(prefix.length));
      if (!match) continue;
      const bounds = el.getBoundingClientRect();
      if (!bounds.width && !bounds.height) continue;
      const key = Number(match[1]), col = Number(match[2]);
      if (!cells.has(key)) cells.set(key, new Map());
      const line = cells.get(key);
      if (!line.has(col)) line.set(col, []);
      line.get(col).push(el);
    }
    const keys = [...cells.keys()].sort((a,b) => a-b);
    if (row !== undefined && rows !== keys.length) return undefined;
    const elements = row === undefined ? [...cells.values()].flatMap(r => [...r.values()].flat()) : cells.get(keys[row])?.get(column);
    return boxOf(document, elements);
  };
}
