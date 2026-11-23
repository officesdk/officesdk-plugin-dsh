import { AttentionOverlay } from './attention-overlay.mjs';
import { SELECTION_TIMING, selectionAt, selectionMs } from './selection.mjs';

/**
 * The agent's layer for flowing documents (Word, spreadsheets): up to a few tracks, each the v1.1 selection outline
 * (AttentionOverlay), the agent cursor and its status (first-generation interaction handoff v1.1; spec 261105 §7),
 * plus one status bar. A track follows what its `find()` locates every frame (`{ left, top, width, height, clip? }` in
 * the page's CSS pixels, or undefined to hide) and goes through: cursor at the top-left → the outline pulled open to
 * the bottom-right ("Agent · 框选中") → full outline with the task's status while the caller writes → a short tail →
 * gone. Everything is the plugin's own elements on the page body with inline styles: no CSS is injected and the SDK's
 * DOM is only read. While busy, typing and presses are held back; the wheel still scrolls.
 */
const SVG_NS = 'http://www.w3.org/2000/svg';
const GAP = 9;
const RADIUS = 10;
const GLIDE_MS = 250;
const CURSOR_ROOM = 38;
const ARROW_PX = 20;
const PILL_LEFT = 10;
const MAIN = 'main';
const KEYS = ['keydown', 'keypress', 'beforeinput', 'paste', 'cut', 'drop', 'compositionstart'];
const PRESSES = ['pointerdown', 'mousedown', 'touchstart', 'dblclick', 'auxclick', 'contextmenu', 'dragstart'];
/** Keys that only move the view: let them through so the user can still look around while the AI writes. */
const VIEW_KEYS = new Set(['PageUp', 'PageDown', 'Home', 'End']);

export const DOC_LABELS = {
  zh: { editing: 'AI 编辑中', step: (n, total) => `AI 正在写 第 ${n}/${total} 步`, marks: { add: 'Agent · 写入中', edit: 'Agent · 修改中', delete: 'Agent · 删除中' }, finishing: '正在完成…' },
  en: { editing: 'AI editing', step: (n, total) => `AI is writing step ${n} of ${total}`, marks: { add: 'Agent · Writing', edit: 'Agent · Editing', delete: 'Agent · Deleting' }, finishing: 'Finishing…' },
};

export function createDocOverlay(document, { animate = () => true, wait = ms => new Promise(resolve => setTimeout(resolve, ms)) } = {}) {
  const view = document.defaultView;
  const now = () => view?.performance?.now?.() ?? Date.now();
  const tracks = new Map();
  const makeTrack = () => {
    const box = document.createElement('div');
    box.dataset.shimoAgent = 'frame';
    Object.assign(box.style, { position: 'fixed', zIndex: '2147483647', display: 'none', pointerEvents: 'none', boxSizing: 'border-box' });
    const cursor = document.createElement('div');
    cursor.dataset.shimoAgent = 'cursor';
    Object.assign(cursor.style, { position: 'fixed', zIndex: '2147483647', display: 'none', width: '0', height: '0', pointerEvents: 'none' });
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
    document.body.append(box, cursor);
    return { box, cursor, arrow, pill, outline: undefined, edge: false, sized: undefined, locate: undefined, label: '', selecting: '', shown: '', pillWidth: 0, above: false,
      start: undefined, animate: true, drawn: undefined, glide: undefined, settle: undefined, tail: undefined };
  };
  const trackOf = key => { let track = tracks.get(key); if (!track) tracks.set(key, track = makeTrack()); return track; };
  const status = document.createElement('div');
  status.setAttribute('role', 'status');
  Object.assign(status.style, { position: 'fixed', left: '50%', bottom: '16px', transform: 'translateX(-50%)', zIndex: '2147483647', display: 'none', padding: '6px 12px',
    font: '12px/18px system-ui, sans-serif', color: '#fff', background: 'rgba(17,24,39,.86)', borderRadius: '14px', pointerEvents: 'none' });
  document.body.append(status);

  let busy = false, frameId;
  const stop = event => { event.preventDefault?.(); event.stopImmediatePropagation?.(); event.stopPropagation?.(); };
  const guardKeys = event => { if (busy && !(event.type === 'keydown' && VIEW_KEYS.has(event.key) && !event.shiftKey)) stop(event); };
  const guardPress = event => { if (busy) stop(event); };
  for (const type of KEYS) document.addEventListener(type, guardKeys, true);
  for (const type of PRESSES) (view ?? document).addEventListener?.(type, guardPress, true);

  const clipTo = (element, rect, clip) => {
    element.style.clipPath = clip ? `inset(${clip.top - rect.top}px ${rect.left + rect.width - clip.right}px ${rect.top + rect.height - clip.bottom}px ${clip.left - rect.left}px)` : '';
  };
  const setLabel = (track, text) => { if (text !== track.shown) { track.shown = text; track.pill.textContent = text; track.pillWidth = 0; } };
  /** The cursor at `point` (the dragged corner) while selecting; otherwise at the bottom-right, or top-right without room. */
  const placeCursor = (track, rect, clip, dragging) => {
    const { cursor, arrow, pill } = track;
    if (!rect || !track.shown) { cursor.style.display = 'none'; return; }
    const bounds = clip ?? { left: 0, top: 0, right: view?.innerWidth ?? Infinity, bottom: view?.innerHeight ?? Infinity };
    const right = rect.left + rect.width, bottom = rect.top + rect.height;
    const flip = !dragging && bottom + CURSOR_ROOM > bounds.bottom;
    const y = flip ? rect.top : bottom;
    if (flip !== track.above) {
      track.above = flip;
      arrow.style.transform = flip ? 'scaleY(-1)' : '';
      Object.assign(pill.style, flip ? { top: '-8px', transform: 'translateY(-100%)' } : { top: '10px', transform: '' });
    }
    if (!track.pillWidth) track.pillWidth = pill.getBoundingClientRect?.().width || 0;
    const width = Math.max(ARROW_PX, PILL_LEFT + track.pillWidth);
    Object.assign(cursor.style, { display: 'block', left: `${right}px`, top: `${y}px` });
    // Hidden as a whole when the corner or the status would leave the view, rather than cut off (handoff 02).
    const top = flip ? y - CURSOR_ROOM + 2 : y, end = flip ? y : y + CURSOR_ROOM;
    cursor.style.visibility = right >= bounds.left && right + width <= bounds.right && top >= bounds.top && end <= bounds.bottom ? 'visible' : 'hidden';
  };
  const place = track => {
    let found;
    try { found = track.locate?.(); } catch { found = undefined; }
    const gap = track.edge ? 0 : GAP;
    let rect = found && found.width > 0 && found.height > 0 ? { left: found.left - gap, top: found.top - gap, width: found.width + 2 * gap, height: found.height + 2 * gap } : undefined;
    if (rect && track.glide) {
      const t = Math.min(1, (now() - track.glide.start) / GLIDE_MS), ease = 1 - (1 - t) ** 3;
      rect = Object.fromEntries(['left', 'top', 'width', 'height'].map(key => [key, track.glide.from[key] + (rect[key] - track.glide.from[key]) * ease]));
      if (t >= 1) track.glide = undefined;
    }
    track.drawn = rect;
    const stage = selectionAt(track.start, now(), { animate: track.animate });
    if (!rect || stage.phase === 'cursor') {
      track.box.style.display = 'none';
      track.outline?.hide(true);
      track.sized = undefined;
      setLabel(track, track.selecting);
      // The cursor waits at the top-left corner before it pulls the selection open.
      placeCursor(track, rect && { left: rect.left, top: rect.top, width: 0, height: 0 }, found?.clip, true);
      return;
    }
    const { box } = track;
    Object.assign(box.style, { display: 'block', left: `${rect.left}px`, top: `${rect.top}px`, width: `${rect.width}px`, height: `${rect.height}px` });
    clipTo(box, rect, found.clip);
    if (!track.outline) { track.outline = new AttentionOverlay(box, { edge: track.edge }); track.outline.svg.style.opacity = '0.7'; }
    const radiusPx = track.edge ? 0 : RADIUS;
    if (stage.phase === 'selecting') {
      track.sized = undefined;
      track.outline.focus({ whole: true }, { immediate: true, radiusPx, selectionProgress: stage.progress });
      setLabel(track, track.selecting);
      placeCursor(track, { left: rect.left, top: rect.top, width: rect.width * stage.progress, height: rect.height * stage.progress }, found.clip, true);
      return;
    }
    const size = `${Math.round(rect.width)}x${Math.round(rect.height)}`;
    if (size !== track.sized) { track.sized = size; track.outline.focus({ whole: true }, { immediate: true, radiusPx }); }
    setLabel(track, track.label);
    placeCursor(track, rect, found.clip, false);
  };
  const active = () => [...tracks.values()].some(track => track.locate);
  const loop = () => {
    frameId = undefined;
    if (!active()) return;
    for (const track of tracks.values()) if (track.locate) place(track);
    frameId = view?.requestAnimationFrame?.(loop);
  };
  const follow = () => {
    for (const track of tracks.values()) place(track);
    if (frameId === undefined && active()) frameId = view?.requestAnimationFrame?.(loop);
  };
  const drop = key => {
    const track = tracks.get(key);
    if (!track) return;
    clearTimeout(track.tail);
    track.settle?.();
    track.outline?.dispose();
    track.box.remove(); track.cursor.remove();
    tracks.delete(key);
  };
  const clear = () => {
    for (const key of [...tracks.keys()]) drop(key);
    if (frameId !== undefined) view?.cancelAnimationFrame?.(frameId);
    frameId = undefined;
  };
  return {
    get holding() { return busy; },
    busy(on, message) { busy = Boolean(on); status.style.display = on && message ? 'block' : 'none'; if (message) status.textContent = message; if (!on) clear(); },
    say(message) { status.textContent = message ?? ''; status.style.display = message ? 'block' : 'none'; },
    /**
     * Starts track `key` on what `find()` locates: cursor, then the selection pulled open (`selecting` as its status),
     * then `label`. Resolves once the selection is complete, which is when the caller may write (timed, so a target that
     * cannot be located yet never holds the writing up). `edge`: outline on the target's edge, square (cells, pages).
     */
    select(key, find, { label = '', selecting = '', edge = false } = {}) {
      const track = trackOf(key ?? MAIN);
      clearTimeout(track.tail);
      track.settle?.();
      if (edge !== track.edge) { track.outline?.dispose(); track.outline = undefined; track.edge = edge; }
      Object.assign(track, { locate: find, label, selecting: selecting || label, animate: animate(), start: now(), glide: undefined, drawn: undefined, sized: undefined });
      follow();
      const ms = selectionMs({ animate: track.animate });
      return new Promise(resolve => {
        let settled = false;
        track.settle = () => { if (!settled) { settled = true; track.settle = undefined; resolve(); } };
        if (!ms) track.settle(); else void wait(ms).then(() => track.settle?.());
      });
    },
    /** Moves track `key` to what `find()` locates now (the written target), gliding there; the selection stays done. */
    point(key, find, { label } = {}) {
      const track = tracks.get(key ?? MAIN);
      if (!track) return;
      track.glide = track.drawn ? { from: { ...track.drawn }, start: now() } : undefined;
      track.locate = find;
      if (label !== undefined) track.label = label;
      follow();
    },
    status(key, label) { const track = tracks.get(key ?? MAIN); if (track) { track.label = label; follow(); } },
    /** Ends track `key` after the short tail of the handoff. */
    release(key) {
      const track = tracks.get(key ?? MAIN);
      if (!track) return;
      clearTimeout(track.tail);
      track.tail = setTimeout(() => drop(key ?? MAIN), animate() ? SELECTION_TIMING.tailMs : 0);
    },
    /** Compatibility: one track showing what `find()` locates, already selected, with `text` as its status. */
    frame(find, { text = '', key } = {}) {
      if (!find) { drop(key ?? MAIN); return; }
      const track = trackOf(key ?? MAIN);
      track.glide = track.drawn ? { from: { ...track.drawn }, start: now() } : undefined;
      Object.assign(track, { locate: find, label: text, selecting: text, start: undefined, animate: false });
      follow();
    },
    get tracks() { return [...tracks.keys()]; },
    dispose() {
      clear(); busy = false;
      for (const type of KEYS) document.removeEventListener(type, guardKeys, true);
      for (const type of PRESSES) (view ?? document).removeEventListener?.(type, guardPress, true);
      status.remove();
    },
  };
}
