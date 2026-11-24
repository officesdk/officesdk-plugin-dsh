import { PPTX_LIMITS, placeholderKey, validatePptxOps } from '../agent/pptx.mjs';
import { AttentionOverlay } from './attention-overlay.mjs';
import { AGENT_LAYER, selectionAt, selectionMs, textChunks } from './selection.mjs';
import { agentLabel, conclusionIndex, pptStatus } from './agent-status.mjs';

const EMU_PER_POINT = 12700;
const CARET = '▍';
const agentError = (code, message) => Object.assign(new Error(message), { code });

/** The deck as the user sees it: ids, geometry in points, text, and the current selection. */
export async function inspectDeck(api) {
  const deck = { slides: [], selectedSlideIds: [], selectedShapeIds: [] };
  await api.PowerPoint.run(async ctx => {
    const slides = ctx.presentation.slides;
    slides.load('items');
    await ctx.sync();
    const collections = slides.items.map(slide => { const shapes = slide.shapes; shapes.load('items/id,items/name,items/type,items/left,items/top,items/width,items/height'); return shapes; });
    await ctx.sync();
    deck.slides = slides.items.map((slide, index) => ({ id: slide.id, shapes: collections[index].items.map(shape => ({ id: shape.id, name: shape.name, type: shape.type, left: shape.left, top: shape.top, width: shape.width, height: shape.height })) }));
    // Layouts the model can name in addSlide, and the slide size it positions shapes in; optional on older editors.
    try {
      const masters = ctx.presentation.slideMasters; masters.load('items/id,items/name');
      const setup = ctx.presentation.pageSetup; setup.load('slideWidth,slideHeight');
      await ctx.sync();
      const layoutSets = masters.items.map(master => { const layouts = master.layouts; layouts.load('items/id,items/name'); return layouts; });
      await ctx.sync();
      deck.layouts = layoutSets.flatMap(layouts => layouts.items.map(layout => ({ id: layout.id, name: layout.name })));
      deck.slideSize = { width: setup.slideWidth, height: setup.slideHeight };
      // Converted decks may lose layout names: each layout also says which placeholders addSlide can fill.
      try {
        const shapeSets = layoutSets.flatMap(layouts => layouts.items.map(layout => { const shapes = layout.shapes; shapes.load('items/id,items/type'); return shapes; }));
        await ctx.sync();
        const formats = shapeSets.map(shapes => shapes.items.filter(shape => shape.type === 'Placeholder').map(shape => { const format = shape.placeholderFormat; format.load('type'); return format; }));
        await ctx.sync();
        deck.layouts.forEach((layout, index) => { layout.placeholders = [...new Set(formats[index].map(format => placeholderKey(format.type)).filter(Boolean))]; });
      } catch {}
    } catch {}
    const selectedSlides = ctx.presentation.getSelectedSlides(); selectedSlides.load('items/id');
    const selectedShapes = ctx.presentation.getSelectedShapes(); selectedShapes.load('items/id');
    await ctx.sync().catch(() => {});
    try { deck.selectedSlideIds = selectedSlides.items.map(item => item.id); } catch {}
    try { deck.selectedShapeIds = selectedShapes.items.map(item => item.id); } catch {}
  });
  // Text is read per slide: one shape without a text frame must not hide the rest of the slide.
  for (const slide of deck.slides) {
    const read = async shapes => api.PowerPoint.run(async ctx => {
      const ranges = shapes.map(shape => { const range = ctx.presentation.slides.getItem(slide.id).shapes.getItem(shape.id).textFrame.textRange; range.load('text'); return range; });
      await ctx.sync();
      shapes.forEach((shape, index) => { shape.text = ranges[index].text; });
    });
    try { await read(slide.shapes); }
    catch { for (const shape of slide.shapes) await read([shape]).catch(() => {}); }
  }
  return deck;
}

const GLOW = 'linear-gradient(135deg, #438bfa 0%, #9464ec 32%, #d96eb9 64%, #18b8c7 100%)';
const GLIDE_MS = 250;
const PRESS = new Set(['pointerdown', 'mousedown', 'touchstart']);
const MOVE = new Set(['pointermove', 'mousemove', 'touchmove']);
const RELEASE = new Set(['pointerup', 'mouseup', 'touchend', 'click']);
const POINTER = [...PRESS, ...MOVE, ...RELEASE, 'dblclick', 'auxclick', 'contextmenu', 'dragstart', 'drop'];

/**
 * The slide drawn in the editor's main view: its screen box, points-to-pixels scale, the slide id it shows
 * and the area that cuts it off. Thumbnails reuse the surface class; the editor names the main one's
 * definitions `mop-main-…`. Without that name the largest surface is taken (zoomed far out it can be
 * smaller than a thumbnail, so the name comes first).
 */
export function locateSlide(document) {
  let surface, box, main = false;
  for (const item of document.querySelectorAll('svg.ppt-slide-surface')) {
    const named = String(item.getAttribute?.('data-mop-svg-definition-namespace') ?? '').startsWith('mop-main');
    const rect = item.getBoundingClientRect();
    if (!box || (named && !main) || (named === main && rect.width * rect.height > box.width * box.height)) { surface = item; box = rect; main = named; }
  }
  const view = surface?.viewBox?.baseVal;
  if (!surface || !view?.width || !box.width) return undefined;
  return { surface, box, named: main, scale: box.width / (view.width / EMU_PER_POINT), shown: surface.getAttribute?.('data-mop-svg-slide') ?? undefined, area: clipArea(surface) };
}

/** The editor's views as `PowerPoint.Public.setViewMode` names them (SDK, 2026-11-05; office-kit K17). */
export const SORTER = 'slideSorter';
export const NORMAL = 'normal';

/**
 * `PowerPoint.Public.setViewMode` when the editor has it, else undefined (older SDKs: slides are built in the normal
 * view as before). Its product actions may be unbound on some hosts and then fail when called (`InvalidContext`), so
 * the switch returns a promise that rejects on a thrown error as well as on a rejected one.
 */
export function viewSwitch(api) {
  try {
    const host = api?.PowerPoint?.Public;
    const setViewMode = host?.setViewMode;
    if (typeof setViewMode !== 'function') return undefined;
    return mode => Promise.resolve().then(() => setViewMode.call(host, mode));
  } catch { return undefined; }
}

/** A surface the browser lays out: one of its sides is 0 while it is hidden. */
const shown = rect => Boolean(rect?.width && rect.height);
const namespaceOf = surface => String(surface.getAttribute?.('data-mop-svg-definition-namespace') ?? '');

/**
 * The view the editor shows, read from its page (no public getViewMode yet, K17; the same recorded dependency as K9):
 * `normal` while the main view's slide (`mop-main-…`) is drawn, `slideSorter` while only slide cards
 * (`mop-thumbnail-…`) are, undefined when it cannot tell.
 */
export function readView(document) {
  if (typeof document?.querySelectorAll !== 'function') return undefined;
  let cards = false;
  for (const surface of document.querySelectorAll('svg.ppt-slide-surface')) {
    if (!shown(surface.getBoundingClientRect?.()) || surface.closest?.('[inert]')) continue;
    const namespace = namespaceOf(surface);
    if (namespace.startsWith('mop-main')) return NORMAL;
    if (namespace.startsWith('mop-thumbnail')) cards = true;
  }
  return cards ? SORTER : undefined;
}

/**
 * Where the slide sorter draws `slideId`'s card: the largest thumbnail surface of that slide (the sorter's cards are
 * the slide list's thumbnails laid out as a grid; K9 dependency, to be confirmed on the SDK with setViewMode) and the
 * list that cuts it off. Only the whole card is outlined, so the card's own coordinate system is not needed.
 */
export function locateCard(document, slideId) {
  if (!slideId || typeof document?.querySelectorAll !== 'function') return undefined;
  let surface, box;
  for (const item of document.querySelectorAll('svg.ppt-slide-surface')) {
    if (!namespaceOf(item).startsWith('mop-thumbnail') || item.getAttribute?.('data-mop-svg-slide') !== slideId) continue;
    const rect = item.getBoundingClientRect();
    if (shown(rect) && (!box || rect.width * rect.height > box.width * box.height)) { surface = item; box = rect; }
  }
  return surface && { box, area: listArea(surface) };
}

/** Where `geometry` (points) of `slideId` is on screen in `slide`; undefined while another slide is shown. */
export function project(slide, geometry, slideId) {
  if (!slide || !geometry || (slideId && slide.shown && slide.shown !== slideId)) return undefined;
  if (geometry.id && slide.surface?.querySelectorAll) {
    const shape = Array.from(slide.surface.querySelectorAll('[data-mop-move-preview-id]'))
      .find(node => node.getAttribute('data-mop-move-preview-id') === geometry.id);
    const bounds = shape?.getBoundingClientRect();
    if (bounds?.width && bounds.height) return { left: bounds.left, top: bounds.top, width: bounds.width, height: bounds.height, clip: slide.area?.getBoundingClientRect() };
    if (!sized(geometry)) return undefined;
  }
  const matrix = slide.surface?.getScreenCTM?.();
  if (matrix) {
    const x = geometry.left * EMU_PER_POINT, y = geometry.top * EMU_PER_POINT;
    const w = geometry.width * EMU_PER_POINT, h = geometry.height * EMU_PER_POINT;
    const corners = [[x,y], [x+w,y], [x,y+h], [x+w,y+h]].map(([u,v]) => ({
      x: matrix.a*u + matrix.c*v + matrix.e, y: matrix.b*u + matrix.d*v + matrix.f
    }));
    const xs = corners.map(p => p.x), ys = corners.map(p => p.y);
    const left = Math.min(...xs), top = Math.min(...ys);
    return { left, top, width: Math.max(...xs) - left, height: Math.max(...ys) - top, clip: slide.area?.getBoundingClientRect() };
  }
  const { box, scale } = slide;
  return { left: box.left + geometry.left * scale, top: box.top + geometry.top * scale, width: Math.max(geometry.width * scale, 8), height: Math.max(geometry.height * scale, 8), clip: slide.area?.getBoundingClientRect() };
}

/** Screen rectangle of a shape on the slide drawn in the main view; see `project`. */
export function screenRect(document, geometry, slideId) { return project(locateSlide(document), geometry, slideId); }

const clipAreas = new WeakMap();
/** The nearest ancestor that cuts off what overflows it, so the box is cut off with the slide. */
function clipArea(element) {
  if (clipAreas.has(element)) return clipAreas.get(element);
  const view = element.ownerDocument?.defaultView;
  let area;
  for (let node = element.parentElement; node && view?.getComputedStyle; node = node.parentElement) {
    const style = view.getComputedStyle(node);
    if (/(hidden|auto|scroll|clip)/.test(`${style.overflow} ${style.overflowX} ${style.overflowY}`)) { area = node; break; }
  }
  clipAreas.set(element, area);
  return area;
}

const listAreas = new WeakMap();
/**
 * The slide list a thumbnail sits in: its nearest scrolling ancestor. Each thumbnail also has a wrapper of its
 * own size that hides overflow, so the nearest clipping ancestor is that wrapper, not the list: cutting a ring
 * off there would hide all of it. Without a scrolling ancestor, the nearest clipping one.
 */
function listArea(thumbnail) {
  if (listAreas.has(thumbnail)) return listAreas.get(thumbnail);
  const view = thumbnail.ownerDocument?.defaultView;
  let area;
  for (let node = thumbnail.parentElement; node && view?.getComputedStyle; node = node.parentElement) {
    const style = view.getComputedStyle(node);
    if (/(auto|scroll)/.test(`${style.overflow} ${style.overflowX} ${style.overflowY}`)) { area = node; break; }
  }
  area ??= clipArea(thumbnail);
  listAreas.set(thumbnail, area);
  return area;
}

/** A slide in the list: the editor names its thumbnails' definitions `mop-thumbnail-…`; other surfaces (the main one, any elsewhere) are not. */
const isThumbnail = (surface, slide) => surface !== slide?.surface && String(surface.getAttribute?.('data-mop-svg-definition-namespace') ?? '').startsWith('mop-thumbnail');
/** Controls the list may hold over a thumbnail; a press on one is not a click on the slide. */
const CONTROLS = 'button, a, input, select, textarea, [role=button], [role=menuitem]';
const inside = (rect, x, y) => Boolean(rect) && x >= rect.left && x <= rect.left + (rect.width ?? rect.right - rect.left) && y >= rect.top && y <= rect.top + (rect.height ?? rect.bottom - rect.top);
/**
 * What a pointer event at (x, y) during an AI edit may reach: `thumbnail` (the visible part of a slide in
 * the list: a click selects it, a drag would reorder) and `scrollbar` (the main view's scroll bars) pass;
 * the slide, the handles around a selected shape, the ribbon and other controls are `other` and stopped.
 * Judged by position: the editor lays its own elements over the thumbnails, so the target is not the slide.
 * Thumbnails only pass when the main view is known by name: otherwise a small main slide could pass as one;
 * and only surfaces named as thumbnails, where no control (a button in the list) is under the pointer.
 */
export function pointerTarget(document, x, y, slide = locateSlide(document)) {
  if (!slide || !Number.isFinite(x) || !Number.isFinite(y)) return 'other';
  if (inside(slide.box, x, y)) return 'other';
  if (slide.named) {
    for (const surface of document.querySelectorAll('svg.ppt-slide-surface')) {
      if (!isThumbnail(surface, slide) || !inside(surface.getBoundingClientRect(), x, y)) continue;
      // Only where the list shows the thumbnail (it may be scrolled partly out) and nothing floats over it
      // there (a ribbon menu, a context menu): what is under the pointer belongs to the list.
      const list = listArea(surface), hit = document.elementFromPoint?.(x, y);
      if (list && inside(list.getBoundingClientRect(), x, y) && (!hit || list === hit || (list.contains?.(hit) && !hit.closest?.(CONTROLS)))) return 'thumbnail';
    }
  }
  const area = slide.area, rect = area?.getBoundingClientRect();
  if (!rect || !inside(rect, x, y)) return 'other';
  const width = area.clientWidth, height = area.clientHeight, left = rect.left + (area.clientLeft ?? 0), top = rect.top + (area.clientTop ?? 0);
  return (Number.isFinite(width) && x >= left + width) || (Number.isFinite(height) && y >= top + height) ? 'scrollbar' : 'other';
}

/** True when `target` is inside the slide list (the nearest scrolling ancestor of a thumbnail). */
function inSlideList(document, target, slide) {
  if (!target || !slide?.named) return false;
  for (const surface of document.querySelectorAll('svg.ppt-slide-surface')) {
    const list = isThumbnail(surface, slide) && listArea(surface);
    if (list && (list === target || list.contains?.(target))) return true;
  }
  return false;
}

const NAVIGATION_KEYS = new Set(['ArrowUp', 'ArrowDown', 'ArrowLeft', 'ArrowRight', 'PageUp', 'PageDown', 'Home', 'End']);
/**
 * Keys that may reach the editor during an AI edit: navigation while the slide list has focus (anywhere else
 * an arrow could nudge a selected shape), copy, Escape and bare modifiers. Zoom shortcuts are not among them
 * (Ctrl+= is also a text format); the wheel and pinch zoom instead.
 */
export function keyAllowedWhileBusy(event, document, slide = locateSlide(document)) {
  if (event.type !== 'keydown') return false;
  if (event.metaKey || event.ctrlKey) return !event.shiftKey && !event.altKey && String(event.key).toLowerCase() === 'c';
  if (NAVIGATION_KEYS.has(event.key)) return !event.shiftKey && !event.altKey && inSlideList(document, event.target, slide);
  return event.key === 'Escape' || event.key === 'Shift' || event.key === 'Control' || event.key === 'Meta' || event.key === 'Alt';
}

/**
 * Highlight boxes, typed captions, thumbnail rings and input guard over the editor; purely visual, never part of
 * the document. While the AI applies an edit, editing keys and pointer input on the slide and the ribbon are stopped
 * before the editor sees them; scrolling, zooming (wheel, pinch) and picking a slide in the list still work, and
 * between edits the user edits the deck directly. Each track (`key`; one per slide when several slides are written
 * at once) has its own box and caption, which follow their target — `{ slideId, geometry }` in points — on every
 * animation frame, so scrolling, zooming, resizing or switching slides moves them or hides them with their slide:
 * only the tracks of the slide in the main view are drawn. `mark` rings the thumbnails of the slides being written.
 */
export function createOverlay(document, { animate: motion = () => true, pause = ms => new Promise(resolve => setTimeout(resolve, ms)) } = {}) {
  const view = document.defaultView;
  const now = () => view?.performance?.now?.() ?? Date.now();
  const makeTrack = () => {
    // The box covers what is outlined (a `text` target plus its gap, or an `edge` target exactly); the outline is
    // the AttentionOverlay drawn inside it, made when the box is first shown and again when its kind changes.
    const box = document.createElement('div');
    box.dataset.shimoAgent = 'frame';
    Object.assign(box.style, { position: 'fixed', zIndex: AGENT_LAYER.selection, display: 'none', pointerEvents: 'none', boxSizing: 'border-box' });
    // The agent's cursor: an arrow at a corner of the drawn box and its status beside it. Not inside the box, so the
    // box's clip never cuts it in half: it hides as a whole instead.
    const cursor = document.createElement('div');
    cursor.dataset.shimoAgent = 'cursor';
    Object.assign(cursor.style, { position: 'fixed', zIndex: AGENT_LAYER.cursor, display: 'none', width: '0', height: '0', pointerEvents: 'none' });
    const arrow = document.createElementNS(SVG_NS, 'svg');
    arrow.setAttribute('viewBox', '0 0 20 20');
    Object.assign(arrow.style, { position: 'absolute', left: '0', top: '0', width: '20px', height: '20px', overflow: 'visible', transformOrigin: '1px 1px', filter: 'drop-shadow(0 1px 1px #0002)' });
    const tip = document.createElementNS(SVG_NS, 'path');
    for (const [name, value] of Object.entries({ d: 'M1 1 18 7 11 11 7 18Z', fill: '#292d34', stroke: 'white', 'stroke-width': '1.3', 'stroke-linejoin': 'round' })) tip.setAttribute(name, value);
    arrow.append(tip);
    const pill = document.createElement('span');
    Object.assign(pill.style, { position: 'absolute', left: '10px', top: '10px', display: 'flex', alignItems: 'center', height: '28px', boxSizing: 'border-box', padding: '3px 10px',
      border: '1px solid #282c32', borderRadius: '24px', background: '#282c32', color: '#fff', whiteSpace: 'nowrap', font: '500 12px/18px -apple-system, BlinkMacSystemFont, "PingFang SC", system-ui, sans-serif' });
    cursor.append(arrow, pill);
    // Typed text drawn over the shape while the AI writes; the document only takes the final value.
    const caption = document.createElement('div');
    caption.dataset.shimoAgent = 'caption';
    Object.assign(caption.style, { position: 'fixed', zIndex: AGENT_LAYER.selection, display: 'none', pointerEvents: 'none', overflow: 'hidden', whiteSpace: 'pre-wrap', wordBreak: 'break-word',
      font: '16px/1.3 system-ui, sans-serif', color: '#111827', background: 'rgba(255,255,255,.94)', padding: '2px 4px', boxSizing: 'border-box' });
    return { box, cursor, arrow, pill, caption, outline: undefined, sized: undefined, kind: 'text', label: '', selecting: '', pillWidth: 0, above: false,
      framed: undefined, written: undefined, text: '', glide: undefined, drawn: undefined, start: undefined, motion: false, settle: undefined };
  };
  const main = makeTrack();
  const tracks = new Map([[MAIN_TRACK, main]]);
  const trackOf = key => {
    let track = tracks.get(key ?? MAIN_TRACK);
    if (!track) { track = makeTrack(); tracks.set(key, track); document.body.append(track.box, track.cursor, track.caption); }
    return track;
  };
  /** Releases a track's outline: its frames and listeners go with it. */
  const dropOutline = track => { track.outline?.dispose(); track.outline = undefined; track.sized = undefined; };
  const status = document.createElement('div');
  status.setAttribute('role', 'status');
  Object.assign(status.style, { position: 'fixed', left: '50%', bottom: '16px', transform: 'translateX(-50%)', zIndex: '2147483647', display: 'none', padding: '6px 12px',
    font: '12px/18px system-ui, sans-serif', color: '#fff', background: 'rgba(17,24,39,.86)', borderRadius: '14px', pointerEvents: 'none' });
  let busy = false, press;
  const stopEvent = event => { event.preventDefault?.(); event.stopImmediatePropagation?.(); event.stopPropagation?.(); };
  const guardKeys = event => { if (busy && !keyAllowedWhileBusy(event, document)) stopEvent(event); };
  const guardPointer = event => {
    if (!busy) return;
    const point = event.touches?.[0] ?? event;
    if (PRESS.has(event.type)) {
      // Only a plain primary press: another button or Ctrl-click opens a menu.
      const plain = (event.button ?? 0) === 0 && !event.ctrlKey && !event.metaKey;
      const kind = plain ? pointerTarget(document, point.clientX, point.clientY) : 'other';
      press = kind === 'other' ? undefined : { kind, x: point.clientX, y: point.clientY, dragging: false };
      if (!press) stopEvent(event);
    } else if (MOVE.has(event.type)) {
      // A press on a thumbnail may click it; once it moves it would drag the slide to a new place. A held
      // button without a press let through (a drag begun before this edit) moves nothing either.
      if (press?.kind === 'thumbnail' && !press.dragging && Math.hypot(point.clientX - press.x, point.clientY - press.y) > 4) press.dragging = true;
      if (press ? press.dragging : (event.buttons ?? 0) !== 0 || event.type === 'touchmove') stopEvent(event);
    } else if (RELEASE.has(event.type)) {
      // The release of a press let through always reaches the editor, so it never stays pressed; with the
      // moves in between stopped it sees a click at most. Known limit: a press made before the edit began
      // has its release stopped like any other, so a drag held across the start of an edit stays with the
      // editor until the next click.
      if (!press) stopEvent(event);
      // The click ends the press (pointerup and mouseup come before it).
      if (event.type === 'click' || event.type === 'touchend') press = undefined;
    } else stopEvent(event);
  };
  const KEYS = ['keydown', 'keypress', 'beforeinput', 'paste', 'cut'];
  const target = view ?? document;
  document.body.append(main.box, main.cursor, main.caption, status);
  for (const type of KEYS) document.addEventListener(type, guardKeys, true);
  for (const type of POINTER) target.addEventListener?.(type, guardPointer, true);
  let frameId, marked = [], mode = NORMAL;
  const rings = [];
  // Insets from the element's own box to the area the slide is drawn in; negative ones keep the glow.
  const clipTo = (element, rect, clip) => {
    element.style.clipPath = clip ? `inset(${clip.top - rect.top}px ${rect.left + rect.width - clip.right}px ${rect.top + rect.height - clip.bottom}px ${clip.left - rect.left}px)` : '';
  };
  const active = () => marked.length > 0 || [...tracks.values()].some(track => track.framed || track.written);
  /**
   * The cursor at the drawn box `rect`: its arrow points at the bottom-right corner, or at the top-right one when
   * the status would not fit below within `clip` (the view without one). Hidden while the corner or the status
   * would leave that area, rather than cut off or pointing away from its box.
   */
  const placeCursor = (track, rect, clip, dragging = false) => {
    const { cursor, arrow, pill } = track;
    if (!rect || !pill.textContent) { cursor.style.display = 'none'; return; }
    const bounds = clip ?? { left: 0, top: 0, right: view?.innerWidth ?? Infinity, bottom: view?.innerHeight ?? Infinity };
    const right = rect.left + rect.width, bottom = rect.top + rect.height;
    // While the selection is pulled open the cursor stays on the dragged corner (handoff v1.1 02).
    const above = !dragging && bottom + CURSOR_ROOM > bounds.bottom;
    const y = above ? rect.top : bottom;
    if (above !== track.above) {
      track.above = above;
      arrow.style.transform = above ? 'scaleY(-1)' : '';
      Object.assign(pill.style, above ? { top: '-8px', transform: 'translateY(-100%)' } : { top: '10px', transform: '' });
    }
    Object.assign(cursor.style, { display: 'block', left: `${right}px`, top: `${y}px` });
    // Measured once the cursor is laid out and again while that gave nothing; the text only changes with frame().
    if (!track.pillWidth) track.pillWidth = pill.getBoundingClientRect?.().width || 0;
    const width = Math.max(ARROW_PX, PILL_LEFT + track.pillWidth);
    const top = above ? y - CURSOR_ROOM + 2 : y, end = above ? y : y + CURSOR_ROOM;
    const inside = right >= bounds.left && right + width <= bounds.right && top >= bounds.top && end <= bounds.bottom;
    cursor.style.visibility = inside ? 'visible' : 'hidden';
  };
  /** Puts one track's box, cursor and caption over their targets as `slide` is drawn now; hides them when it is not shown. */
  const placeTrack = (track, slide) => {
    const { box, caption, framed, written, pill } = track;
    let rect, clip;
    if (framed?.whole) {
      // The whole slide, edge to edge, as the main view draws it; like `project`, only while that slide is shown.
      // In the slide sorter: its card, where the sorter shows it now.
      track.drawn = { slideId: framed.slideId, whole: true };
      if (mode === SORTER) {
        const card = locateCard(document, framed.slideId);
        if (card) { rect = { left: card.box.left, top: card.box.top, width: card.box.width, height: card.box.height }; clip = card.area?.getBoundingClientRect(); }
      } else if (slide && (!slide.shown || slide.shown === framed.slideId)) {
        rect = { left: slide.box.left, top: slide.box.top, width: slide.box.width, height: slide.box.height };
        clip = slide.area?.getBoundingClientRect();
      }
    } else {
      let geometry = framed?.geometry;
      if (track.glide) {
        // A new target on the same slide glides there in slide coordinates, so it also follows the view meanwhile.
        const t = Math.min(1, (now() - track.glide.start) / GLIDE_MS), ease = 1 - (1 - t) ** 3;
        geometry = Object.fromEntries(['left', 'top', 'width', 'height'].map(key => [key, track.glide.from[key] + (framed.geometry[key] - track.glide.from[key]) * ease]));
        if (t >= 1) track.glide = undefined;
      }
      track.drawn = framed && geometry && { slideId: framed.slideId, geometry };
      const outline = framed && project(slide, geometry, framed.slideId);
      if (outline) {
        // A text target is outlined around it, an edge target on its own edge.
        const gap = track.kind === 'edge' ? 0 : TEXT_GAP;
        rect = { left: outline.left - gap, top: outline.top - gap, width: outline.width + 2 * gap, height: outline.height + 2 * gap };
        clip = outline.clip;
      }
    }
    // v1.1 selection gesture: cursor at the top-left, then the outline pulled open, then the task's status.
    const stage = selectionAt(track.start, now(), { animate: track.motion });
    const shown = stage.phase === 'done' ? track.label : track.selecting || track.label;
    if (shown !== pill.textContent) { pill.textContent = shown; track.pillWidth = 0; }
    let cursorAt = rect, dragging = false;
    if (!rect || stage.phase === 'cursor') {
      box.style.display = 'none';
      track.outline?.hide(true);
      track.sized = undefined;
      cursorAt = rect && { left: rect.left, top: rect.top, width: 0, height: 0 };
      dragging = true;
    } else {
      Object.assign(box.style, { display: 'block', left: `${rect.left}px`, top: `${rect.top}px`, width: `${rect.width}px`, height: `${rect.height}px` });
      clipTo(box, rect, clip);
      if (!track.outline) {
        track.outline = new AttentionOverlay(box, { edge: track.kind === 'edge' });
        // The whole selection, its light included, is drawn at 70%; the cursor is not.
        track.outline.svg.style.opacity = '0.7';
      }
      const radiusPx = track.kind === 'edge' ? 0 : TEXT_RADIUS;
      if (stage.phase === 'selecting') {
        track.sized = undefined;
        track.outline.focus({ whole: true }, { immediate: true, radiusPx, selectionProgress: stage.progress });
        cursorAt = { left: rect.left, top: rect.top, width: rect.width * stage.progress, height: rect.height * stage.progress };
        dragging = true;
      } else {
        // The box moves with its style; the drawing only changes with its size (its light runs on the outline's own frames).
        const size = `${rect.width}x${rect.height}`;
        if (size !== track.sized) { track.sized = size; track.outline.focus({ whole: true }, { immediate: true, radiusPx }); }
      }
    }
    placeCursor(track, cursorAt, clip, dragging);
    const area = written && project(slide, written.geometry, written.slideId);
    if (!area) caption.style.display = 'none';
    else {
      const size = Math.max(10, Math.min(area.height * 0.55, 40));
      Object.assign(caption.style, { display: 'block', left: `${area.left}px`, top: `${area.top}px`, width: `${area.width}px`, height: `${area.height}px`, fontSize: `${size}px` });
      if (caption.textContent !== track.text) caption.textContent = track.text;
      clipTo(caption, area, area.clip);
    }
  };
  const ringFor = index => {
    if (rings[index]) return rings[index];
    const ring = document.createElement('div');
    ring.dataset.shimoAgent = 'mark';
    // A gradient border only: the mask cuts out its middle (and would cut off a shadow, so there is none).
    Object.assign(ring.style, { position: 'fixed', zIndex: AGENT_LAYER.selection, display: 'none', pointerEvents: 'none', borderRadius: '6px', boxSizing: 'border-box', padding: '3px', background: GLOW, ...CUT_OUT });
    document.body.append(ring);
    return (rings[index] = ring);
  };
  /** Rings each marked slide's thumbnail where the list shows it now; no ring where the list does not show one. */
  const placeMarks = slide => {
    const surfaces = marked.length ? [...document.querySelectorAll('svg.ppt-slide-surface')] : [];
    marked.forEach((slideId, index) => {
      const ring = ringFor(index);
      const thumbnail = surfaces.find(surface => isThumbnail(surface, slide) && surface.getAttribute?.('data-mop-svg-slide') === slideId);
      const box = thumbnail?.getBoundingClientRect();
      if (!box?.width || !box.height) { ring.style.display = 'none'; return; }
      const rect = { left: box.left - 4, top: box.top - 4, width: box.width + 8, height: box.height + 8 };
      Object.assign(ring.style, { display: 'block', left: `${rect.left}px`, top: `${rect.top}px`, width: `${rect.width}px`, height: `${rect.height}px` });
      clipTo(ring, rect, listArea(thumbnail)?.getBoundingClientRect());
    });
    for (let index = marked.length; index < rings.length; index++) rings[index].style.display = 'none';
  };
  const place = () => {
    if (mode === SORTER) {
      // No main view: the largest surface would be a card, so nothing is projected onto it; only whole-card frames are
      // drawn, and no rings (the frame on the card is the mark).
      for (const track of tracks.values()) placeTrack(track, undefined);
      for (const ring of rings) ring.style.display = 'none';
      return;
    }
    const slide = active() ? locateSlide(document) : undefined;
    // Slide tracks exist only while several slides are written at once. Without the slide id of the main view (K9)
    // any of them could land on the wrong slide, so none is drawn for the rest of that edit.
    const unknown = slide && !slide.shown && [...tracks.keys()].some(key => key !== MAIN_TRACK);
    for (const [key, track] of tracks) placeTrack(track, unknown && key !== MAIN_TRACK ? undefined : slide);
    placeMarks(slide);
  };
  const loop = () => {
    frameId = undefined;
    if (!active()) return;
    place();
    frameId = view?.requestAnimationFrame?.(loop);
  };
  const follow = () => {
    place();
    if (frameId === undefined && active()) frameId = view?.requestAnimationFrame?.(loop);
  };
  const stop = () => {
    for (const [key, track] of tracks) {
      Object.assign(track, { framed: undefined, written: undefined, glide: undefined, drawn: undefined });
      // Tracks of other slides last one edit.
      if (key !== MAIN_TRACK) { dropOutline(track); track.box.remove(); track.cursor.remove(); track.caption.remove(); tracks.delete(key); }
    }
    marked = [];
    mode = NORMAL;
    if (frameId !== undefined) view?.cancelAnimationFrame?.(frameId);
    frameId = undefined;
    place();
  };
  return {
    busy(on, message) { busy = Boolean(on); press = undefined; status.style.display = on && message ? 'block' : 'none'; if (message) status.textContent = message; if (!on) stop(); },
    /** Typed text over `target` (`{ slideId, geometry }`) on track `key`; no target or text hides it. */
    write(target, value, key) {
      const track = trackOf(key);
      track.written = target?.geometry && value !== undefined ? target : undefined;
      track.text = value ?? '';
      follow();
    },
    say(message) { status.textContent = message; status.style.display = message ? 'block' : 'none'; },
    /**
     * Outlines `target` on track `key` until its next call, with the agent's cursor and `text` as its status (no
     * text, no cursor); no target hides both. `target` is `{ slideId, geometry }` (points) or `{ slideId, whole: true }`
     * for the whole slide. `kind` is `text` (outlined around it, rounded) or `edge` (on its edge, square; a whole
     * slide always is). `tone` (edit, delete) only chooses the caller's text and does not change the look.
     */
    frame(target, { text: title = '', kind = 'text', key, select = false, selecting = '' } = {}) {
      const track = trackOf(key);
      const whole = Boolean(target?.whole && target.slideId);
      const next = whole ? { slideId: target.slideId, whole: true } : target?.geometry ? target : undefined;
      const look = whole || kind === 'edge' ? 'edge' : 'text';
      // From where the box is drawn now, also when the last glide has not finished; only between two shapes of
      // one slide in one look: to or from a whole slide, or to another look (gap and corners change), it jumps.
      const from = track.drawn;
      track.glide = !select && from?.geometry && next?.geometry && from.slideId === next.slideId && look === track.kind && track.box.style.display !== 'none' ? { from: { ...from.geometry }, start: now() } : undefined;
      if (look !== track.kind) { dropOutline(track); track.kind = look; }
      track.framed = next;
      track.label = title ?? '';
      track.selecting = selecting;
      track.settle?.();
      // `select`: the v1.1 gesture before writing; the promise resolves when the outline is fully open (timed, so a
      // target that is not on screen never holds the writing up). Without it the frame is simply there.
      track.motion = select && Boolean(next) && motion();
      track.start = track.motion ? now() : undefined;
      if (track.motion) track.glide = undefined;
      follow();
      const ms = track.motion ? selectionMs() : 0;
      return new Promise(resolve => {
        let settled = false;
        track.settle = () => { if (!settled) { settled = true; track.settle = undefined; resolve(); } };
        if (!ms) track.settle(); else void pause(ms).then(() => track.settle?.());
      });
    },
    /**
     * The view the editor was switched to (`slideSorter` or `normal`): in the sorter, whole-slide frames are drawn on
     * the slides' cards and nothing else is projected. Back to `normal` when the edit ends.
     */
    view(next) {
      mode = next === SORTER ? SORTER : NORMAL;
      follow();
    },
    /** Rings the thumbnails of `slideIds` (the slides being written) until the next call; none clears them. */
    mark(slideIds) {
      marked = [...new Set((slideIds ?? []).filter(Boolean))];
      follow();
    },
    dispose() {
      stop(); busy = false;
      for (const type of KEYS) document.removeEventListener(type, guardKeys, true);
      for (const type of POINTER) target.removeEventListener?.(type, guardPointer, true);
      dropOutline(main);
      main.box.remove(); main.cursor.remove(); main.caption.remove(); status.remove();
      for (const ring of rings.splice(0)) ring.remove();
    },
  };
}

const MAIN_TRACK = 'main';
const SVG_NS = 'http://www.w3.org/2000/svg';
/** A text target's outline: this far outside it, with these corners (design handoff v1). */
const TEXT_GAP = 9;
const TEXT_RADIUS = 10;
/** The cursor: the arrow's size, where its status starts, and the height both take below (or above) the corner. */
const ARROW_PX = 20;
const PILL_LEFT = 10;
const CURSOR_ROOM = 38;
const CUT_OUT = { WebkitMask: 'linear-gradient(#000 0 0) content-box, linear-gradient(#000 0 0)', WebkitMaskComposite: 'xor', maskComposite: 'exclude' };

/** Page numbers as a range ("3–6") when they follow each other, else listed. */
const pages = (numbers, separator) => {
  const sorted = [...numbers].sort((a, b) => a - b);
  return sorted.length > 1 && sorted.every((value, index) => index === 0 || value === sorted[index - 1] + 1) ? `${sorted[0]}–${sorted.at(-1)}` : sorted.join(separator);
};

export const LABELS = {
  zh: { editing: 'AI 编辑中', deleting: 'AI 删除中', adding: 'AI 新增', marks: { edit: 'Agent · 修改中', add: 'Agent · 新增中', delete: 'Agent · 删除中' }, slide: n => `AI 正在修改第 ${n} 页`, slides: numbers => `AI 正在生成第 ${pages(numbers, '、')} 页`, saving: '正在完成…',
    sorter: '幻灯片浏览 · 创建空白页', toNormal: n => `切换普通视图，展开第 ${n} 页` },
  en: { editing: 'AI editing', deleting: 'AI deleting', adding: 'AI adding', marks: { edit: 'Agent · Editing', add: 'Agent · Adding', delete: 'Agent · Deleting' }, slide: n => `AI is editing slide ${n}`, slides: numbers => `AI is building slides ${pages(numbers, ', ')}`, saving: 'Finishing…',
    sorter: 'Slide sorter · Creating blank slides', toNormal: n => `Switching to normal view, slide ${n}` },
};

/** Below this width of the main view's area two readable cards do not fit side by side (spec Q4 1). */
const SORTER_MIN_WIDTH = 520;
/** The track of the frame on the card being created in the slide sorter. */
const SORTER_TRACK = 'sorter';

/** At most this many slides are written at the same time: one batch's new slides (spec §11.2). */
export const MAX_TOGETHER = PPTX_LIMITS.maxNewSlides;

/**
 * Whether a batch's slides play at the same time (spec §11.3): animated, typed and outlined on the overlay (the
 * editor's own outline only reaches the slide shown; typing into the document would write every intermediate
 * state), building new slides, 2–`MAX_TOGETHER` slides (each new slide plus each existing slide it changes), and
 * no slide deleted or moved (their order matters). Anything else runs step by step as before.
 */
export function playsTogether(steps, { animate = true, typing = 'document', outline = 'overlay' } = {}) {
  if (!animate || (typing !== 'overlay' && typing !== 'chunks') || outline !== 'overlay') return false;
  if (steps.some(step => step.op === 'deleteSlide' || step.op === 'moveSlide')) return false;
  const added = steps.filter(step => step.op === 'addSlide').length;
  if (!added) return false;
  const existing = new Set(steps.filter(step => step.op !== 'addSlide' && step.slideRef === undefined && step.slideId !== undefined).map(step => step.slideId));
  const count = added + existing.size;
  return count >= 2 && count <= MAX_TOGETHER;
}

/**
 * Runs validated steps with fixed Office.js calls. With `animate`, each step is shown where it lands:
 * the slide is selected, the area outlined, and text is cleared and typed before its exact final value.
 * Steps that target an object created earlier in the batch (`slideRef` / `shapeRef`) get its real id here;
 * `refs` collects them for the result. The outline is the overlay's frame, which follows the slide on screen.
 * `outline: 'editor'` uses the editor's own `focusAttention` instead: office-kit K8 — its outline leaves the
 * shape after the slide repaints — so it stays off until that is fixed.
 *
 * A batch that `playsTogether` first builds all its new slides in order, selects the first of its slides once,
 * then writes every slide at the same time, one track each (its own outline and caption, its thumbnail ringed).
 * Every editor call of the batch goes through one queue, so tracks only interleave between calls; once a track
 * fails the others stop at their next call and the error is thrown after all of them have stopped, so the caller
 * can undo the whole batch. `report.together` says which way the batch ran.
 */
export async function runSteps(api, steps, { overlay, document, animate = true, signal, wait = ms => new Promise(resolve => setTimeout(resolve, ms)), labels = LABELS.zh, geometryOf = () => undefined, typing = 'document', outline = 'overlay', refs = {}, report = {}, locale, viewOf = readView, settleMs = 240, planned = false } = {}) {
  let halted;
  // v1.1 (spec §7): status texts made from what each step does, when the caller names a locale.
  const statusOf = (step, geometry) => locale ? agentLabel(locale, pptStatus(step, { shape: geometry, conclusion: conclusions.has(step) })) : undefined;
  const selecting = locale ? agentLabel(locale, 'selecting') : '';
  const conclusions = new Set();
  /** Throws once the batch is cancelled, or once another track failed. */
  const check = () => { signal?.throwIfAborted(); if (halted) throw halted; };
  let queue = Promise.resolve();
  const call = body => { const next = queue.then(() => { check(); return api.PowerPoint.run(body); }); queue = next.catch(() => {}); return next; };
  const run = body => call(async ctx => { await body(ctx); await ctx.sync(); });
  const slideRefs = new Map(), shapeRefs = new Map(), born = new Map(), fills = new Map();
  const shapeOf = (ctx, step) => ctx.presentation.slides.getItem(step.slideId).shapes.getItem(step.shapeId);
  const order = async () => { let ids = []; await call(async ctx => { const slides = ctx.presentation.slides; slides.load('items'); await ctx.sync(); ids = slides.items.map(slide => slide.id); }); return ids; };
  let attention = outline === 'editor', attended = false;
  /** The editor's outline over `geometry` on the active slide (null clears it); false when it has none. */
  const attend = async geometry => {
    if (!attention) return false;
    try { await run(ctx => { ctx.presentation.focusAttention(geometry ? { left: geometry.left, top: geometry.top, width: geometry.width, height: geometry.height } : null); }); attended = true; return true; }
    catch { attention = false; return false; }
  };
  /** The track a step runs on: the main one selects each slide it works on; one of several slides selects nothing. */
  const MAIN = { key: undefined, select: true };
  let focused;
  const focus = async (slideId, lane) => {
    if (!slideId || !lane.select || (lane.same && focused === slideId)) return;
    focused = slideId;
    await run(ctx => ctx.presentation.setSelectedSlides([slideId])).catch(() => {});
    const position = (await order()).indexOf(slideId);
    overlay?.say(labels.slide(position + 1));
    overlay?.mark?.([slideId]);
    if (animate) await wait(180);
  };
  /** The agent's status on the overlay (`labels.marks`; captions passed without them get the Chinese ones). */
  const marks = labels.marks ?? LABELS.zh.marks;
  /**
   * Outlines `geometry` of the step's slide with the cursor's `text`: on its edge for a chart (`kind` edge; the editor
   * reports the shape's type), around it for anything else.
   */
  const frame = async (step, geometry, tone, text, lane, kind = 'edge') => {
    if (!overlay) return;
    // A box without size (a placeholder that only inherits its layout's position reads 0,0,0,0) has no place to point at.
    if (!sized(geometry) && !geometry?.id) geometry = undefined;
    // A deletion always uses the overlay, so its "deleting" cursor shows; with `outline: 'editor'` the rest uses the editor's outline.
    if (tone !== 'delete' && sized(geometry) && lane === MAIN && await attend(geometry)) overlay.frame(undefined);
    else if (locale) {
      // v1.1: the cursor pulls the selection open; nothing is written before it is complete.
      await overlay.frame(geometry && { slideId: step.slideId, geometry }, { tone, text: statusOf(step, geometry) ?? text, kind, select: animate, selecting, ...(lane.key !== undefined && { key: lane.key }) });
      return;
    } else overlay.frame(geometry && { slideId: step.slideId, geometry }, { tone, text, kind, ...(lane.key !== undefined && { key: lane.key }) });
    if (animate) await wait(220);
  };
  /** Outlines the whole slide on its edge, always on the overlay (the editor's outline only takes a shape's box). */
  const frameSlide = async (slideId, tone, text, lane) => {
    if (!overlay) return;
    // The editor's outline of the last shape goes, so the two are not shown together.
    if (attended && lane === MAIN) { await attend(null); attended = false; }
    if (locale) { await overlay.frame({ slideId, whole: true }, { tone, text: agentLabel(locale, tone === 'delete' ? 'delete' : 'shape'), kind: 'edge', select: animate, selecting, ...(lane.key !== undefined && { key: lane.key }) }); return; }
    overlay.frame({ slideId, whole: true }, { tone, text, kind: 'edge', ...(lane.key !== undefined && { key: lane.key }) });
    if (animate) await wait(220);
  };
  const typeInto = async (slideId, shapeId, to, geometry, lane) => {
    if (!animate) return;
    // `chunks` (v1.1): the text grows in the shape itself in a few real writes (at most four); the caller writes the
    // whole text last. A shape not created yet has nothing to write into: it is created with its first part.
    if (typing === 'chunks') {
      if (!shapeId) return;
      const parts = planned ? textChunks(to, { min: 4 }) : textChunks(to);
      for (const part of parts.slice(0, -1)) {
        check();
        await run(ctx => { ctx.presentation.slides.getItem(slideId).shapes.getItem(shapeId).textFrame.textRange.text = part; });
        await wait(planned ? 110 : 60);
      }
      return;
    }
    // `overlay`: the typing is drawn over the shape and every intermediate state stays out of the
    // document (a Kit record would keep each one as a change).
    if (typing === 'overlay') {
      const target = sized(geometry) ? { slideId, geometry } : undefined;
      const chunks = Math.min(Math.max(to.length, 1), 36);
      for (let index = 1; index <= chunks; index++) {
        check();
        overlay?.write(target, to.slice(0, Math.ceil(to.length * index / chunks)) + (index < chunks ? CARET : ''), lane.key);
        await wait(28);
      }
      return;
    }
    const target = ctx => ctx.presentation.slides.getItem(slideId).shapes.getItem(shapeId).textFrame.textRange;
    await run(ctx => { target(ctx).text = ''; });
    await wait(120);
    const chunks = Math.min(Math.max(to.length, 1), 36);
    for (let index = 1; index <= chunks; index++) {
      check();
      const prefix = to.slice(0, Math.ceil(to.length * index / chunks));
      await run(ctx => { target(ctx).text = prefix + (index < chunks ? CARET : ''); });
      await wait(28);
    }
  };
  const clearCaption = lane => overlay?.write(undefined, undefined, lane.key);
  /** Writes text into a shape just created: typed first on the overlay, or typed into the shape. */
  const fillText = async (slideId, shapeId, text, geometry, lane) => {
    if (!text) return;
    if (animate && typing !== 'overlay') await typeInto(slideId, shapeId, text, geometry, lane);
    check();
    await run(ctx => { ctx.presentation.slides.getItem(slideId).shapes.getItem(shapeId).textFrame.textRange.text = text; });
    clearCaption(lane);
  };
  const styleText = (shape, step) => {
    if (step.font) Object.assign(shape.textFrame.textRange.font, step.font);
    if (step.align) shape.textFrame.textRange.paragraphFormat.horizontalAlignment = step.align;
  };
  /** Adds one shape with `add(shapes)` on the step's slide; returns its id and remembers its ref and box. */
  const create = async (step, add) => {
    let id;
    await call(async ctx => {
      const shape = add(ctx.presentation.slides.getItem(step.slideId).shapes, ctx);
      shape.load('id');
      await ctx.sync();
      id = shape.id;
    });
    step.createdId = id;
    born.set(`${step.slideId}/${id}`, step.geometry);
    if (step.ref) { shapeRefs.set(step.ref, id); refs[step.ref] = id; }
    return id;
  };
  /**
   * Adds the slide of an addSlide step at its place and readies its placeholders: the ones the step fills are kept
   * for `fills` (id, kind, box), the rest — and any given empty text — would show "click to add" prompts on a
   * generated page, and no later step can reach them, so they go now.
   */
  const build = async step => {
    const before = await order();
    await run(ctx => { ctx.presentation.slides.add(step.layoutId ? { layoutId: step.layoutId } : undefined); });
    const created = (await order()).find(id => !before.includes(id));
    if (created && step.index !== (await order()).indexOf(created)) await run(ctx => { ctx.presentation.slides.getItem(created).moveTo(step.index); });
    step.createdId = created;
    step.slideId = created;
    if (step.ref && created) { slideRefs.set(step.ref, created); refs[step.ref] = created; }
    const work = [];
    fills.set(step, work);
    if (!created || !step.layoutId) return;
    let targets = [];
    await call(async ctx => {
      const shapes = ctx.presentation.slides.getItem(created).shapes;
      shapes.load('items/id,items/type,items/left,items/top,items/width,items/height');
      await ctx.sync();
      const holders = shapes.items.filter(shape => shape.type === 'Placeholder');
      const formats = holders.map(shape => { const format = shape.placeholderFormat; format.load('type'); return format; });
      await ctx.sync();
      targets = holders.map((shape, index) => ({ id: shape.id, key: placeholderKey(formats[index].type), geometry: { left: shape.left, top: shape.top, width: shape.width, height: shape.height } }));
    });
    // A new page's placeholders inherit their place from the layout, and the editor reports them as 0,0,0,0:
    // the outline takes the layout's placeholder of the same kind instead.
    if (targets.some(target => !sized(target.geometry))) {
      await call(async ctx => {
        const shapes = ctx.presentation.slides.getItem(created).layout.shapes;
        shapes.load('items/type,items/left,items/top,items/width,items/height');
        await ctx.sync();
        const holders = shapes.items.filter(shape => shape.type === 'Placeholder');
        const formats = holders.map(shape => { const format = shape.placeholderFormat; format.load('type'); return format; });
        await ctx.sync();
        // Matched by kind; a kind the layout has twice (two body columns) cannot say which box is whose.
        const keys = formats.map(format => placeholderKey(format.type));
        const inherited = new Map(holders.flatMap((shape, index) => keys.indexOf(keys[index]) === keys.lastIndexOf(keys[index]) ? [[keys[index], { left: shape.left, top: shape.top, width: shape.width, height: shape.height }]] : []));
        for (const target of targets) if (!sized(target.geometry) && sized(inherited.get(target.key))) target.geometry = inherited.get(target.key);
      }).catch(error => { if (signal?.aborted || halted) throw error; });
    }
    // The check before anything ran read the layout; this one reads the page the editor made (a backstop).
    for (const [key, text] of Object.entries(step.placeholders ?? {})) {
      if (!text) continue;
      const target = targets.find(item => item.key === key);
      if (!target) throw Object.assign(new Error(`Layout ${step.layoutName ?? step.layoutId} has no ${key} placeholder.`), { code: 'invalid-edit' });
      targets = targets.filter(item => item !== target);
      work.push({ target, text });
    }
    if (targets.length) await run(ctx => { for (const target of targets) ctx.presentation.slides.getItem(created).shapes.getItem(target.id).delete(); });
  };
  const runStep = async (step, lane) => {
    check();
    if (step.slideRef !== undefined) step.slideId = slideRefs.get(step.slideRef);
    if (step.shapeRef !== undefined) step.shapeId = shapeRefs.get(step.shapeRef);
    await focus(step.slideId, lane);
    const geometry = step.shapeId ? geometryOf(step.slideId, step.shapeId) ?? born.get(`${step.slideId}/${step.shapeId}`) : undefined;
    if (planned) {
      await planned.selectModule(step, geometry ?? step.geometry, statusOf(step, geometry), animate, wait);
      check();
    }
    switch (step.op) {
      case 'setText':
        await frame(step, geometry, 'edit', marks.edit, lane);
        await typeInto(step.slideId, step.shapeId, step.text, geometry, lane);
        await run(ctx => { shapeOf(ctx, step).textFrame.textRange.text = step.text; });
        clearCaption(lane);
        break;
      case 'setFont':
        await frame(step, geometry, 'edit', marks.edit, lane);
        await run(ctx => { Object.assign(shapeOf(ctx, step).textFrame.textRange.font, step.font); });
        break;
      case 'setAlignment':
        await frame(step, geometry, 'edit', marks.edit, lane);
        await run(ctx => { shapeOf(ctx, step).textFrame.textRange.paragraphFormat.horizontalAlignment = step.align; });
        break;
      case 'setFill':
        await frame(step, geometry, 'edit', marks.edit, lane);
        await run(ctx => { shapeOf(ctx, step).fill.setSolidColor(step.color); });
        break;
      case 'setGeometry':
        await frame(step, geometry, 'edit', marks.edit, lane);
        await run(ctx => { Object.assign(shapeOf(ctx, step), step.geometry); });
        await frame(step, { ...geometry, ...step.geometry }, 'edit', marks.edit, lane);
        break;
      case 'addTextBox': {
        await frame(step, step.geometry, 'edit', marks.add, lane);
        const typedFirst = typing === 'overlay';
        if (typedFirst) await typeInto(step.slideId, undefined, step.text, step.geometry, lane);
        const id = await create(step, shapes => {
          const shape = shapes.addTextBox(animate && !typedFirst ? '' : step.text, step.geometry);
          // A text box is text on the slide: no theme fill or outline.
          shape.fill.clear();
          shape.lineFormat.visible = false;
          styleText(shape, step);
          return shape;
        });
        if (animate && !typedFirst) {
          await typeInto(step.slideId, id, step.text, undefined, lane);
          await run(ctx => { ctx.presentation.slides.getItem(step.slideId).shapes.getItem(id).textFrame.textRange.text = step.text; });
        }
        clearCaption(lane);
        break;
      }
      case 'addShape': {
        await frame(step, step.geometry, 'edit', marks.add, lane);
        const id = await create(step, shapes => {
          const shape = shapes.addGeometricShape(step.shape, { ...step.geometry, ...(step.fill && { fillColor: step.fill }), ...(step.line === 'none' && { lineVisible: false }) });
          if (step.fill) shape.fill.setSolidColor(step.fill);
          if (step.line && step.line !== 'none') shape.lineFormat.color = step.line;
          return shape;
        });
        if (step.text) {
          if (typing === 'overlay') await typeInto(step.slideId, id, step.text, step.geometry, lane);
          await fillText(step.slideId, id, step.text, step.geometry, lane);
        }
        if (step.font || step.align) await run(ctx => { styleText(ctx.presentation.slides.getItem(step.slideId).shapes.getItem(id), step); });
        break;
      }
      case 'addImage':
        await frame(step, step.geometry, 'edit', marks.add, lane);
        await create(step, shapes => shapes.addImage(step.imageBase64, step.geometry));
        break;
      case 'addTable':
        await frame(step, step.geometry, 'edit', marks.add, lane);
        await create(step, shapes => shapes.addTable(step.values.length, step.values[0].length, { ...step.geometry, values: step.values }));
        break;
      case 'addChart':
        await frame(step, step.geometry, 'edit', marks.add, lane, 'edge');
        await create(step, shapes => shapes.addChart({ ...step.geometry, chartType: step.chartType, data: step.data, ...(step.title !== undefined && { title: step.title }),
          ...(step.seriesColors && { seriesColors: step.seriesColors }), ...(step.legend && (step.legend === 'none' ? { legendVisible: false } : { legendVisible: true, legendPosition: step.legend })) }));
        break;
      case 'setBackground':
        await frameSlide(step.slideId, 'edit', marks.edit, lane);
        await run(ctx => { ctx.presentation.slides.getItem(step.slideId).background.fill.setSolidFill({ color: step.color }); });
        break;
      case 'deleteShape':
        await frame(step, geometry, 'delete', marks.delete, lane);
        await run(ctx => { shapeOf(ctx, step).delete(); });
        break;
      case 'addSlide': {
        // Built here step by step; a batch played together built it before its tracks started.
        if (!fills.has(step)) { await build(step); await focus(step.createdId, lane); }
        for (const { target, text } of fills.get(step)) {
          check();
          // Its status is the placeholder's: the title is 写入标题, the rest the slide's content (not 创建页面).
          if (locale) await frame({ op: 'setText', slideId: step.slideId }, target.geometry && { ...target.geometry, placeholder: target.key === 'title' ? 'Title' : target.key }, 'edit', marks.add, lane);
          else await frame(step, target.geometry, 'edit', marks.add, lane);
          if (typing === 'overlay') await typeInto(step.slideId, target.id, text, target.geometry, lane);
          await fillText(step.slideId, target.id, text, target.geometry, lane);
        }
        break;
      }
      case 'deleteSlide':
        await frameSlide(step.slideId, 'delete', marks.delete, lane);
        await run(ctx => { ctx.presentation.slides.getItem(step.slideId).delete(); });
        // Not left over the slide the editor shows next.
        overlay?.frame(undefined, lane.key !== undefined ? { key: lane.key } : undefined);
        break;
      case 'moveSlide':
        await run(ctx => { ctx.presentation.slides.getItem(step.slideId).moveTo(step.index); });
        await focus(step.slideId, lane);
        break;
    }
    if (animate) await wait(planned ? 280 : 160);
  };
  const setView = viewSwitch(api);
  const viewNow = () => { try { return viewOf(document); } catch { return undefined; } };
  /** Whether the main view's area is wide enough for the sorter's cards; unknown widths pass. */
  const roomy = () => {
    try {
      const width = typeof document?.querySelectorAll === 'function' ? locateSlide(document)?.area?.clientWidth : undefined;
      return !(width > 0) || width >= SORTER_MIN_WIDTH;
    } catch { return true; }
  };
  /**
   * Back to the normal view after the sorter was asked for (also when that seemed not to happen: the editor may
   * reveal a view late). Not through `call`: after a stop it would throw; it waits for the editor call in flight
   * instead. Tried twice; a failure to switch never hides the error that ended the batch.
   */
  const backToNormal = async () => {
    await queue;
    for (let attempt = 0; attempt < 2; attempt++) {
      if (!await setView(NORMAL).then(() => true, () => false)) continue;
      await wait(settleMs);
      if (viewNow() !== SORTER) return;
    }
    report.view = 'stuck';
  };
  /**
   * v1.1 (handoff: PPT): the blank slides are created one by one in the slide sorter, each card outlined with the
   * agent's cursor, then the editor goes back to the normal view to fill them. When the switch fails or the view does
   * not change, they are built in the normal view as before; when the editor leaves the sorter meanwhile (the user or
   * the editor itself), the rest are built there without outlines. Whatever happens, it ends in the normal view.
   */
  const buildInSorter = async added => {
    // Not asked for at all once the batch is stopped: no flash of the sorter before the stop is reported.
    check();
    let drawing = false;
    try {
      const switched = await setView(SORTER).then(() => true, () => false);
      if (switched) await wait(settleMs);
      drawing = switched && viewNow() !== NORMAL;
      report.view = drawing ? 'sorter' : 'failed';
      if (drawing) { overlay?.view?.(SORTER); overlay?.say(labels.sorter ?? LABELS.zh.sorter); }
      for (const step of added) {
        check();
        await build(step);
        if (drawing && viewNow() === NORMAL) {
          drawing = false;
          report.view = 'left';
          overlay?.frame(undefined, { key: SORTER_TRACK });
          overlay?.view?.(NORMAL);
        }
        if (drawing) await overlay?.frame({ slideId: step.slideId, whole: true }, { text: agentLabel(locale, 'page'), kind: 'edge', select: animate, selecting, key: SORTER_TRACK });
      }
    } finally {
      overlay?.frame(undefined, { key: SORTER_TRACK });
      overlay?.mark?.([]);
      overlay?.view?.(NORMAL);
      await backToNormal();
    }
  };
  /** Builds the new slides in order, then writes every slide of the batch at the same time, one track each. */
  const together = async () => {
    const added = steps.filter(step => step.op === 'addSlide');
    // Only the v1.1 flow, only for at least two new slides, from the normal view (a user already in the sorter
    // keeps it), and only where the cards are readable.
    if (locale) report.view = !setView ? 'unsupported' : added.length >= 2 && viewNow() !== SORTER && roomy() ? 'pending' : 'skipped';
    if (report.view === 'pending') await buildInSorter(added);
    else for (const step of added) { check(); await build(step); }
    const lanes = new Map();
    for (const step of steps) {
      if (step.slideRef !== undefined) step.slideId = slideRefs.get(step.slideRef);
      if (!lanes.has(step.slideId)) lanes.set(step.slideId, []);
      lanes.get(step.slideId).push(step);
    }
    // In deck order: the first is shown, and its track starts first.
    const deck = await order();
    const slideIds = [...lanes.keys()].sort((left, right) => deck.indexOf(left) - deck.indexOf(right));
    if (locale) {
      if (report.view === 'sorter') {
        // Back from the sorter: the batch's first slide is opened, and the caption names it.
        focused = slideIds[0];
        await run(ctx => ctx.presentation.setSelectedSlides([slideIds[0]])).catch(() => {});
        overlay?.say((labels.toNormal ?? LABELS.zh.toNormal)(deck.indexOf(slideIds[0]) + 1));
        if (animate) await wait(300);
        overlay?.say(labels.slide(deck.indexOf(slideIds[0]) + 1));
      }
      // v1.1: the blank slides are all there (the thumbnails show them); now page by page, the current page done
      // before the next, its targets on up to three agent tracks in turn (the editor takes one write at a time).
      for (const slideId of slideIds) {
        check();
        overlay?.mark?.([slideId]);
        let turn = 0;
        for (const step of lanes.get(slideId)) {
          const lane = { key: `${slideId}:${turn++ % 3}`, select: true, same: true };
          await runStep(step, lane);
        }
        for (let track = 0; track < 3; track++) { overlay?.frame(undefined, { key: `${slideId}:${track}` }); overlay?.write(undefined, undefined, `${slideId}:${track}`); }
      }
      return;
    }
    overlay?.mark?.(slideIds);
    await run(ctx => ctx.presentation.setSelectedSlides([slideIds[0]])).catch(() => {});
    overlay?.say(labels.slides(slideIds.map(id => deck.indexOf(id) + 1)));
    if (animate) await wait(180);
    let failure;
    await Promise.allSettled(slideIds.map(async slideId => {
      const lane = { key: slideId, select: false };
      try { for (const step of lanes.get(slideId)) await runStep(step, lane); }
      catch (error) { if (!failure) { failure = error; halted = Object.assign(new Error('Another slide of this edit failed.'), { code: 'halted' }); } }
      finally { overlay?.frame(undefined, { key: slideId }); overlay?.write(undefined, undefined, slideId); }
    }));
    if (failure) throw failure;
  };
  // The 完善结论 text of each slide: the lowest non-title text the batch writes there (spec §7.6 7).
  if (locale) for (const slideId of new Set(steps.map(step => step.slideId ?? step.slideRef).filter(Boolean))) {
    const at = conclusionIndex(steps.map(step => ({ ...step, slideId: step.slideId ?? step.slideRef })), slideId, step => step.shapeId ? geometryOf(step.slideId, step.shapeId) : undefined);
    if (at >= 0) conclusions.add(steps[at]);
  }
  report.together = playsTogether(steps, { animate, typing, outline });
  if (planned) {
    // One lane per returned page. Kit transactions share the serial API queue, animations can overlap.
    const pages = new Map();
    for (const step of steps) {
      if (!pages.has(step.slideId)) pages.set(step.slideId, []);
      pages.get(step.slideId).push(step);
    }
    let failure;
    await Promise.allSettled([...pages].map(async ([id, work]) => {
      try { for (const step of work) await runStep(step, { key: id, select: false }); }
      catch (error) { if (!failure) failure = halted = error; }
    }));
    if (failure) throw failure;
  } else if (report.together) await together();
  else if (locale) {
    // v1.1: one slide's targets on up to three tracks in turn; the slide is selected when the batch moves to it.
    let turn = 0;
    for (const step of steps) await runStep(step, { key: `k${turn++ % 3}`, select: true, same: true });
    for (let track = 0; track < 3; track++) { overlay?.frame(undefined, { key: `k${track}` }); overlay?.write(undefined, undefined, `k${track}`); }
  } else for (const step of steps) await runStep(step, MAIN);
  overlay?.frame(undefined);
  overlay?.mark?.([]);
  if (attended) await attend(null);
  return refs;
}

const sized = geometry => Boolean(geometry && geometry.width > 0 && geometry.height > 0);
/** A chart, as the editor reports a shape's type. */

const same = (left, right) => left.length === right.length && left.every((value, index) => value === right[index]);

/**
 * Walks the editor's own history back to the bytes it had before the edit started. `undo` defaults to the
 * store's; a Kit page passes the editor's undo, the one that also emits each step as a local change.
 */
export function rollback(store, baseline, maxSteps = 2000, undo = () => store.undo()) {
  for (let step = 0; step < maxSteps; step++) {
    if (same(store.encodeBinary(), baseline)) return true;
    if (!store.canUndo) break;
    undo();
  }
  return same(store.encodeBinary(), baseline);
}

async function post(fetch, url, body, type) {
  const response = await fetch(url, { method: 'POST', body, headers: { 'content-type': type }, credentials: 'same-origin' });
  const result = await response.json().catch(() => ({}));
  if (!response.ok) throw agentError(result.error ?? 'editor-failed', result.message ?? `Request failed: ${response.status}`);
  return result;
}

const POLL_MS = 150;
const MIN_GAP_MS = 400;
const RETRY_MS = [1000, 2000, 5000, 10000];

/**
 * Local offline editing: the user's changes stay in the editor and the Host session, never compose/pull.
 * Polls the editor revision while not `paused()` (the AI applying) and stages a change right away, at
 * most one stage per `minGapMs`, so little is still in the editor alone when its tab goes away. A failed
 * stage stays pending, is reported through `onError` and retried with backoff. `flush()` stages a pending
 * change now and resolves true once the Host holds the latest bytes; every stage runs after the previous one.
 */
export function createEditSync({ revision, stage, onManual = () => {}, paused = () => false, onError = () => {}, timers = globalThis, now = () => Date.now(), pollMs = POLL_MS, minGapMs = MIN_GAP_MS, retryMs = RETRY_MS }) {
  let seen = revision(), pending = false, timer, lastStage = -Infinity, failures = 0, queue = Promise.resolve();
  // The queue itself never rejects: a failed stage or edit must not fail the next one.
  const serial = task => { const run = queue.then(task); queue = run.catch(() => {}); return run; };
  const detect = (force = false) => {
    if (!force && paused()) return false;
    const current = revision();
    if (current === seen) return false;
    seen = current; pending = true; onManual();
    return true;
  };
  const schedule = delay => {
    if (timer !== undefined) return;
    timer = timers.setTimeout(run, delay ?? Math.max(0, lastStage + minGapMs - now()));
  };
  // While paused only the AI writes: its revisions are not hand edits, and a stage waits until it is done.
  function run(force = false) {
    timers.clearTimeout(timer); timer = undefined;
    if (!force && paused()) { schedule(minGapMs); return queue.then(() => !pending && failures === 0); }
    if (!pending) return queue.then(() => !pending && failures === 0);
    pending = false;
    return serial(async () => {
      lastStage = now();
      try { await stage(); failures = 0; return true; }
      catch (error) {
        pending = true; failures++;
        onError(error);
        schedule(retryMs[Math.min(failures, retryMs.length) - 1]);
        return false;
      }
    });
  }
  const flush = (force = false) => { detect(force); return run(force); };
  const poll = timers.setInterval(() => { if (detect()) schedule(); }, pollMs);
  return {
    flush,
    /**
     * Runs `task` (an AI edit or undo) after pending stages; its own revisions are not the user's.
     * A task that resolves has staged the whole deck, hand edits included.
     */
    async exclusive(task) {
      await flush(true);
      try {
        const result = await serial(task);
        pending = false; failures = 0; timers.clearTimeout(timer); timer = undefined;
        return result;
      } finally { seen = revision(); }
    },
    /** True when the user changed the deck since the last stage or exclusive task. */
    changed() { const hit = detect(true); if (hit) schedule(); return hit || pending; },
    dispose() { timers.clearInterval(poll); timers.clearTimeout(timer); timer = undefined; },
  };
}

const MAX_UNDO = 20;
const EDIT_ID = /^[A-Za-z0-9_-]{1,64}$/;

/**
 * Frame-side agent for one presentation session: `inspect` reads, `apply` validates against the live
 * deck, animates and stages the result in the Host session (the file is written only by a save),
 * `undo` returns the most recent AI edit to the bytes it started from. Between AI edits the user edits
 * the deck directly; those changes are staged too, and they end the AI undo history.
 */
export function createPresentationAgent({ sdk, document, meta, fetch, channel, animate = () => true, locale = 'zh', wait, signal, timers }) {
  const editor = sdk.getEditor();
  const store = () => editor.getDocument().store;
  const permissions = editor.permissions;
  const labels = LABELS[locale] ?? LABELS.zh;
  let editRevision = meta.editRevision ?? 0;
  const overlay = createOverlay(document);
  let busy = false;
  const undoStack = [];
  const editable = meta.mode === 'edit' && !!meta.stageUrl;
  permissions?.setEditingSuspended?.(!editable);
  const stage = async ({ manual = false } = {}) => {
    const staged = await post(fetch, new URL(`snapshot?base=${editRevision}`, meta.stageUrl).href, store().encodeBinary(), 'application/octet-stream');
    editRevision = staged.editRevision;
    channel?.notify('dirty', { dirty: true, editRevision, ...manual && { manual: true } });
  };
  const restore = baseline => rollback(store(), baseline);
  // A hand edit after an AI edit makes the editor's own undo the only safe way back.
  let lastError;
  const sync = editable ? createEditSync({ revision: () => store().revision, stage: () => stage({ manual: true }), paused: () => busy,
    onManual: () => { undoStack.length = 0; },
    // The panel shows the edit as unsaved and says it has not reached DSH yet; a later stage clears it.
    onError: error => { lastError = error; channel?.notify('dirty', { dirty: true, editRevision, manual: true, unstaged: error?.message || 'The change could not be handed to DSH.' }); },
    timers }) : undefined;
  const flushed = async () => {
    if (!sync || await sync.flush()) return { editRevision };
    throw agentError('stage-failed', `The latest changes could not be handed to DSH${lastError?.message ? `: ${lastError.message}` : ''}.`);
  };
  const view = document.defaultView;
  const flushOnLeave = () => { if (document.visibilityState !== 'visible' || !document.hasFocus?.()) void sync?.flush(); };
  const saveShortcut = event => {
    if (event.type !== 'keydown' || !(event.metaKey || event.ctrlKey) || String(event.key).toLowerCase() !== 's') return;
    event.preventDefault(); event.stopPropagation();
    // Only bytes the Host holds can be saved; a failed hand-over is already shown in the panel.
    if (!busy) void flushed().then(() => channel?.notify('save-request', { editRevision }), () => {});
  };
  if (sync) {
    document.addEventListener('keydown', saveShortcut, true);
    document.addEventListener('visibilitychange', flushOnLeave);
    view?.addEventListener('blur', flushOnLeave);
    signal?.addEventListener('abort', () => {
      sync.dispose(); overlay.dispose();
      document.removeEventListener('keydown', saveShortcut, true);
      document.removeEventListener('visibilitychange', flushOnLeave);
      view?.removeEventListener('blur', flushOnLeave);
    }, { once: true });
  }
  const exclusive = task => sync ? sync.exclusive(task) : task();
  /** One AI edit, after any pending hand edits were staged; the baseline includes them. */
  const applyNow = async ({ ops, editId }, signal) => {
    const api = sdk.getJsApi();
    const baseline = store().encodeBinary();
    const refs = {};
    try {
      const deck = await inspectDeck(api);
      const { steps } = validatePptxOps(ops, deck);
      const geometry = new Map(deck.slides.flatMap(slide => slide.shapes.map(shape => [`${slide.id}/${shape.id}`, shape])));
      overlay.busy(true, labels.editing);
      try {
        await runSteps(api, steps, { overlay, document, signal, labels, wait, animate: animate(), refs, geometryOf: (slideId, shapeId) => geometry.get(`${slideId}/${shapeId}`) });
        // The typing effect writes intermediate text; the result must match the plan exactly.
        const after = await inspectDeck(api);
        const texts = new Map(after.slides.flatMap(slide => slide.shapes.map(shape => [`${slide.id}/${shape.id}`, shape.text])));
        for (const step of steps) if (step.op === 'setText' && texts.get(`${step.slideId}/${step.shapeId}`) !== step.text) throw agentError('editor-failed', 'The presentation did not take the planned text.');
      } catch (error) {
        if (!rollback(store(), baseline)) throw agentError('editor-failed', 'The edit failed and the editor could not return to the saved state; reload the document.');
        throw signal?.aborted ? agentError('cancelled', 'The edit was cancelled; nothing was changed.') : error;
      }
      overlay.say(labels.saving);
      // The editor and the Host session must hold the same deck: an edit the Host did not take is undone.
      try { await stage(); }
      catch (error) { restore(baseline); throw error; }
      if (EDIT_ID.test(editId ?? '')) { undoStack.push({ editId, baseline }); if (undoStack.length > MAX_UNDO) undoStack.shift(); }
      const touched = [...new Set(steps.map(step => step.createdId && step.op === 'addSlide' ? step.createdId : step.slideId).filter(Boolean))];
      const focus = touched.find(id => steps.every(step => !(step.op === 'deleteSlide' && step.slideId === id)));
      if (focus) await api.PowerPoint.run(async ctx => { ctx.presentation.setSelectedSlides([focus]); await ctx.sync(); }).catch(() => {});
      return { editRevision, applied: steps.length, slideIds: touched, ...(Object.keys(refs).length && { refs }), undoable: undoStack.at(-1)?.editId ?? null };
    } finally {
      overlay.frame(undefined);
    }
  };
  return {
    async inspect() { return inspectDeck(sdk.getJsApi()); },
    /** The panel saves or exports what the Host holds: hand edits still in the editor go first. */
    flush: flushed,
    async undo({ editId }) {
      if (meta.mode !== 'edit' || !meta.stageUrl) throw agentError('read-only', 'This presentation is open read only.');
      if (busy) throw agentError('busy', 'Another edit is still running.');
      if (sync?.changed()) undoStack.length = 0;
      const top = undoStack.at(-1);
      if (!top) throw agentError('undo-unavailable', 'There is no AI edit to undo in this editor (the user may have edited it by hand since; they can undo in the editor with Cmd/Ctrl+Z).');
      if (top.editId !== editId) throw agentError('undo-unavailable', undoStack.some(entry => entry.editId === editId) ? 'Only the most recent AI edit can be undone; undo the later edits first.' : 'That edit is no longer in this editor.');
      busy = true;
      overlay.busy(true, labels.saving);
      try {
        return await exclusive(async () => {
          if (!restore(top.baseline)) throw agentError('editor-failed', 'The editor could not return to the state before that edit; reload the document.');
          undoStack.pop();
          await stage();
          return { editRevision, undone: editId, undoable: undoStack.at(-1)?.editId ?? null };
        });
      } finally { busy = false; overlay.busy(false); }
    },
    async apply({ ops, editId }, signal) {
      if (meta.mode !== 'edit' || !meta.stageUrl) throw agentError('read-only', 'This presentation is open read only.');
      if (busy) throw agentError('busy', 'Another edit is still running.');
      busy = true;
      // Input stops before pending hand edits are staged, so nothing lands between the baseline and the edit.
      overlay.busy(true, labels.editing);
      try {
        return await exclusive(() => applyNow({ ops, editId }, signal));
      } finally { busy = false; overlay.busy(false); }
    },
  };
}
