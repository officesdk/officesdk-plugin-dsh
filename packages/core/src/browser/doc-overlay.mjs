import { writerMarquee } from './writer-marquee.mjs';
import { AttentionOverlay } from './attention-overlay.mjs';
import { AGENT_LAYER, SELECTION_TIMING, selectionAt, selectionMs } from './selection.mjs';

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
const RADIUS = 10;
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

export function createDocOverlay(document, { showStatus = true, animate = () => true, wait = ms => new Promise(resolve => setTimeout(resolve, ms)) } = {}) {
  const view = document.defaultView;
  const now = () => view?.performance?.now?.() ?? Date.now();
  const tracks = new Map();
  const makeTrack = () => {
    const box = document.createElement('div');
    box.dataset.shimoAgent = 'frame';
    Object.assign(box.style, { position: 'fixed', zIndex: AGENT_LAYER.selection, display: 'none', pointerEvents: 'none', boxSizing: 'border-box' });
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
    document.body.append(box, cursor);
    return { box, cursor, arrow, pill, fragments: [], outline: undefined, edge: false, sized: undefined, locate: undefined, label: '', selecting: '', shown: '', pillWidth: 0,
      start: undefined, animate: true, settle: undefined, tail: undefined };
  };
  const trackOf = key => { let track = tracks.get(key); if (!track) tracks.set(key, track = makeTrack()); return track; };
  const status = document.createElement('div');
  status.setAttribute('role', 'status');
  Object.assign(status.style, { position: 'fixed', left: '50%', bottom: '16px', transform: 'translateX(-50%)', zIndex: '2147483647', display: 'none', padding: '6px 12px',
    font: '12px/18px system-ui, sans-serif', color: '#fff', background: 'rgba(17,24,39,.86)', borderRadius: '14px', pointerEvents: 'none' });
  if (showStatus) document.body.append(status);

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
  /** Keep the tip on the measured corner. Turn the label inward at the viewport edge. */
  const placeCursor = (track, rect, clip) => {
    const { cursor, arrow, pill } = track;
    if (!rect || !track.shown) { cursor.style.display = 'none'; return; }
    const bounds = clip ?? { left: 0, top: 0, right: view?.innerWidth ?? Infinity, bottom: view?.innerHeight ?? Infinity };
    const right = rect.left + rect.width, bottom = rect.top + rect.height;
    if (right < bounds.left || right > bounds.right || bottom < bounds.top || bottom > bounds.bottom) {
      cursor.style.display = 'none';
      return;
    }
    Object.assign(cursor.style, { display: 'block', left: `${right}px`, top: `${bottom}px` });
    if (!track.pillWidth) track.pillWidth = pill.getBoundingClientRect?.().width || 0;
    let above = bottom + CURSOR_ROOM > bounds.bottom;
    let leftward = right + PILL_LEFT + Math.max(ARROW_PX, track.pillWidth) > bounds.right;
    if (track.writer && tracks.size>1 && track.pillWidth) {
      // Keep the tip attached to its own range. Put the label in the clearest adjacent quadrant.
      const obstacles=[];
      for(const other of tracks.values()){
        if(other.writer)obstacles.push(...(other.bands ?? []));
        if(other===track)continue;
        if(other.labelBox && other.cursor.style.display!=='none')obstacles.push(other.labelBox);
      }
      let best=Infinity;
      for(const [left,up] of [[leftward,above],[!leftward,above],[leftward,!above],[!leftward,!above]]){
        const x=left?right-PILL_LEFT-track.pillWidth:right+PILL_LEFT, y=up?bottom-36:bottom+10;
        if(x<bounds.left || y<bounds.top || x+track.pillWidth>bounds.right || y+28>bounds.bottom)continue;
        const area=obstacles.reduce((sum,r)=>sum+
          Math.max(0,Math.min(x+track.pillWidth,r.left+r.width)-Math.max(x,r.left))*
          Math.max(0,Math.min(y+28,r.top+r.height)-Math.max(y,r.top)),0);
        if(area<best){best=area;leftward=left;above=up;}
        if(!area)break;
      }
    }
    track.labelBox={left:leftward?right-PILL_LEFT-track.pillWidth:right+PILL_LEFT,top:above?bottom-36:bottom+10,width:track.pillWidth,height:28};
    arrow.style.transform = `scale(${leftward ? -1 : 1},${above ? -1 : 1})`;
    Object.assign(pill.style, { left: leftward ? 'auto' : '10px', right: leftward ? '10px' : 'auto',
      top: above ? '-8px' : '10px', transform: above ? 'translateY(-100%)' : '' });
  };
  const place = track => {
    track.previewFits=false;
    let found;
    try { found = track.locate?.(); } catch { found = undefined; }
    const rects = found?.rects ?? (found?.width > 0 && found.height > 0 ? [found] : []);
    track.bands=rects;
    const stage = selectionAt(track.start, now(), { animate: track.animate });
    if (found?.rects?.length && track.writer && track.box.ownerDocument) {
      track.marquee ??= writerMarquee(track.box);
      const progress=track.animate ? 1-Math.pow(1-Math.min(1,Math.max(0,(now()-track.start)/330)),3) : 1;
      const corner=track.marquee(found,progress,track.phase,now(),track.previewText);
      track.previewFits=corner.previewFits;
      clipTo(track.box,found,found.clip);
      setLabel(track,progress<1?track.selecting:track.label);
      placeCursor(track,corner,found.clip);
      return;
    }
    if (track.writer && track.marquee) { track.box.style.display='none'; track.cursor.style.display='none'; return; }
    let remaining = rects.reduce((sum, rect) => sum + rect.width, 0) * (stage.phase === 'done' ? 1 : stage.progress || 0);
    let corner;
    let count = 0;
    for (const rect of rects) {
      if (!remaining || stage.phase === 'cursor') break;
      const progress = Math.min(1, remaining / rect.width);
      const width = Math.min(rect.width, remaining);
      remaining = Math.max(0, remaining - rect.width);
      let part = count ? track.fragments[count - 1] : track;
      if (!part) {
        const box = track.box.cloneNode(false);
        document.body.append(box);
        part = { box };
        track.fragments.push(part);
      }
      count++;
      Object.assign(part.box.style, { display: 'block', left: `${rect.left}px`, top: `${rect.top}px`, width: `${width}px`, height: `${rect.height}px` });
      clipTo(part.box, { ...rect, width }, found.clip);
      if (!part.outline) {
        part.outline = new AttentionOverlay(part.box, { edge: track.edge });
        part.outline.svg.style.opacity = '0.7';
      }
      const size = `${width}x${rect.height}`;
      if (part.sized !== size) {
        part.outline.focus({ whole: true }, { immediate: true, radiusPx: track.edge ? 0 : RADIUS });
        part.sized = size;
      }
      corner = { ...rect, width, height: rect.height * progress };
    }
    for (const part of [track, ...track.fragments].slice(count)) {
      part.box.style.display = 'none';
      part.outline?.hide(true);
      part.sized = undefined;
    }
    setLabel(track, stage.phase === 'done' ? track.label : track.selecting);
    const first = rects[0];
    placeCursor(track, corner ?? (first && { ...first, width: 0, height: 0 }), found?.clip);
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
  let stopWriterObserver;
  function watchWriter() {
    if (stopWriterObserver) return;
    const root = document.querySelector?.('.writer-document-preview');
    if (!root || !view?.MutationObserver || !view.ResizeObserver) return;
    // Writer can finish zoom/re-layout after our animation callback. Re-measure from the
    // final glyphs before paint, including when only the containing pane changed size.
    const sync = () => {
      for (const track of tracks.values()) if (track.writer && track.locate) place(track);
    };
    const resize = new view.ResizeObserver(sync);
    resize.observe(root);
    const scroll = root.querySelector('.writer-document-preview__scroll');
    if (scroll) resize.observe(scroll);
    const paint = new view.MutationObserver(sync);
    paint.observe(root, { subtree:true, childList:true, characterData:true, attributes:true,
      attributeFilter:['style','class','transform','viewBox','x','y','dx','dy','width','height','font-size','textLength'] });
    root.addEventListener('scroll',sync,true);
    stopWriterObserver = () => {
      paint.disconnect(); resize.disconnect(); root.removeEventListener('scroll',sync,true);
      stopWriterObserver = undefined;
    };
  }
  const drop = key => {
    const track = tracks.get(key);
    if (!track) return;
    clearTimeout(track.tail);
    track.settle?.();
    for (const part of track.fragments) { part.outline?.dispose(); part.box.remove(); }
    track.outline?.dispose();
    track.box.remove(); track.cursor.remove();
    tracks.delete(key);
    if (![...tracks.values()].some(item => item.writer)) stopWriterObserver?.();
  };
  const clear = () => {
    for (const key of [...tracks.keys()]) drop(key);
    if (frameId !== undefined) view?.cancelAnimationFrame?.(frameId);
    frameId = undefined;
  };
  return {
    get holding() { return busy; },
    busy(on, message) {
      busy = Boolean(on); status.style.display = on && message ? 'block' : 'none';
      if (message) status.textContent = message;
      if (!on) for (const [key, track] of tracks) if (!track.persistent) drop(key);
    },
    say(message) { status.textContent = message ?? ''; status.style.display = message ? 'block' : 'none'; },
    /**
     * Starts track `key` on what `find()` locates: cursor, then the selection pulled open (`selecting` as its status),
     * then `label`. Resolves once the selection is complete, which is when the caller may write (timed, so a target that
     * cannot be located yet never holds the writing up). `edge`: outline on the target's edge, square (cells, pages).
     */
    select(key, find, { label = '', selecting = '', edge = false, writer = false, persistent = false } = {}) {
      const track = trackOf(key ?? MAIN);
      clearTimeout(track.tail);
      track.settle?.();
      if (edge !== track.edge) { track.outline?.dispose(); track.outline = undefined; track.edge = edge; }
      Object.assign(track, { locate: find, label, writer, persistent, phase:'edge', selecting: selecting || label, animate: animate(), start: now(), sized: undefined });
      if (writer) watchWriter();
      follow();
      const ms = writer && track.animate ? 330 : selectionMs({ animate: track.animate });
      return new Promise(resolve => {
        let settled = false;
        track.settle = () => { if (!settled) { settled = true; track.settle = undefined; resolve(); } };
        if (!ms) track.settle(); else void wait(ms).then(() => track.settle?.());
      });
    },
    async phase(key, phase, duration = 0) {
      const track=tracks.get(key);
      if (!track) return;
      track.phase=phase; follow();
      if (duration && track.animate) await wait(duration);
    },
    preview(key, text) {
      const track=tracks.get(key);
      if (!track) return false;
      track.previewText=text;
      place(track);
      return track.previewFits;
    },
    drop,
    /** Moves track `key` to what `find()` locates now (the written target), the selection stays done. */
    point(key, find, { label } = {}) {
      const track = tracks.get(key ?? MAIN);
      if (!track) return;
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
