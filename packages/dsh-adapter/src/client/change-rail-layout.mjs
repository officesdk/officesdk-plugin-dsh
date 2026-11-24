// Group edits by the lines they occupy, independently of the height of the cards.
export function groupChangeCards(records, anchors = []) {
  const positions = new Map(anchors.map(a => [a.id, a]));
  const newest = [...records].reverse();
  const located = newest.filter(r => positions.has(r.id)).sort((a, b) => positions.get(a.id).top - positions.get(b.id).top);
  const groups = [];
  for (const record of located) {
    const {top, bottom, lineHeight} = positions.get(record.id);
    const last = groups.at(-1);
    if (last && last.bottom - top >= Math.min(last.lineHeight, lineHeight) / 2) {
      last.records.push(record);
      if (bottom > last.bottom) Object.assign(last, {bottom, lineHeight});
    } else {
      groups.push({top, bottom, lineHeight, records:[record]});
    }
  }
  for (const record of newest) {
    if (!positions.has(record.id)) groups.push({records:[record]});
  }
  const order = new Map(newest.map((record, i) => [record.id, i]));
  for (const group of groups) {
    group.records.sort((a,b) => order.get(a.id) - order.get(b.id));
    group.key = group.records.at(-1).id;
  }
  return groups;
}

// Lay out from the first visible anchor in both directions. Cards above the viewport
// can still be reached by scrolling the rail without displacing the first visible edit.
export function positionChangeCards(list, groups, offset, scale = 1, follow = true) {
  const track = list.firstElementChild;
  if (!track || list.hidden || !list.clientHeight || track.children.length !== groups.length) return;
  const gap = 8;
  const heights = Array.from(track.children, el => el.offsetHeight);
  const targets = groups.map(g => g.top === undefined ? undefined : offset + g.top * scale);
  const firstVisible = targets.findIndex(y => y !== undefined && y >= 0);
  const pivot = firstVisible < 0 ? Math.max(0, targets.findLastIndex(Number.isFinite)) : firstVisible;
  const tops = [];
  for (let i = pivot; i < groups.length; ++i) {
    const next = i === pivot ? (targets[i] ?? gap) : tops[i-1] + heights[i-1] + gap;
    tops[i] = Math.max(next, targets[i] ?? next);
  }
  for (let i = pivot - 1; i >= 0; --i) {
    tops[i] = Math.min(targets[i], tops[i+1] - heights[i] - gap);
  }
  const origin = Math.max(0, gap - (tops[0] ?? gap));
  const relativeScroll = follow ? 0 : list.scrollTop - Number(track.dataset.origin || 0);
  for (let i = 0; i < tops.length; ++i) track.children[i].style.top = (origin + tops[i]) + 'px';
  const end = tops.length ? tops.at(-1) + heights.at(-1) + gap : 0;
  track.style.height = Math.ceil(origin + Math.max(list.clientHeight, end)) + 'px';
  track.dataset.origin = String(origin);
  list.scrollTop = origin + relativeScroll;
}

// The new centered card starts where its neighbour was just before the selection changed.
export function slideCardStrip(strip, offset) {
  if (!strip) return;
  for (const motion of strip.getAnimations()) motion.cancel();
  if (!offset || strip.ownerDocument.defaultView.matchMedia('(prefers-reduced-motion: reduce)').matches) return;
  strip.animate([{transform:`translateX(${offset}px)`},{transform:'translateX(0)'}],
    {duration:220,easing:'cubic-bezier(.2,.75,.25,1)'});
}

export function stripOffset(strip) {
  return strip.getBoundingClientRect().left-strip.parentElement.getBoundingClientRect().left;
}

// Let clicks, vertical touch scrolling and the card's editing controls behave normally.
// Capture only after a horizontal drag wins, so a swipe cannot also activate the card.
export function bindCardSwipe(element, step) {
  let press, dragged = false;
  const gesture = {last:-Infinity, distance:0, switched:false};
  function wheel(e) {
    if (e.ctrlKey || !e.target.closest('.office-stack-face')) return;
    const input = e.target.closest('textarea,input');
    if (input && input === element.ownerDocument.activeElement && input.scrollWidth > input.clientWidth) return;
    const dx = e.deltaX || (e.shiftKey ? e.deltaY : 0);
    if (!dx || (!e.shiftKey && Math.abs(e.deltaY) >= Math.abs(dx))) return;
    if (e.cancelable) e.preventDefault();
    if (e.timeStamp - gesture.last > 180 || Math.sign(dx) !== Math.sign(gesture.distance)) {
      gesture.distance = 0;
      gesture.switched = false;
    }
    gesture.last = e.timeStamp;
    const unit = e.deltaMode === 1 ? 16 : e.deltaMode === 2 ? element.clientWidth : 1;
    gesture.distance += dx * unit;
    if (!gesture.switched && Math.abs(gesture.distance) >= 28) {
      gesture.switched = true;
      step(Math.sign(gesture.distance));
    }
  }
  function down(e) {
    if (e.button !== 0 || e.isPrimary === false) return;
    dragged = false;
    if (e.target.closest('form,.office-change-actions,.office-undo-confirm,nav')) return;
    press = {id:e.pointerId, x:e.clientX, y:e.clientY};
  }
  function move(e) {
    if (press?.id !== e.pointerId) return;
    const dx = e.clientX-press.x, dy = e.clientY-press.y;
    if (!dragged) {
      if (Math.abs(dy)>8 && Math.abs(dy)>=Math.abs(dx)) { press=undefined; return; }
      if (Math.abs(dx)<8 || Math.abs(dx)<=Math.abs(dy)) return;
      dragged=true;
      const strip=element.querySelector('.office-stack-strip');
      press.offset=stripOffset(strip);
      for (const motion of strip.getAnimations()) motion.cancel();
      element.setPointerCapture(e.pointerId);
    }
    element.style.setProperty('--office-card-drag', Math.max(-element.clientWidth,Math.min(element.clientWidth,press.offset+dx))+'px');
  }
  function finish(e) {
    if (press?.id !== e.pointerId) return;
    const dx = e.clientX-press.x;
    press=undefined;
    if (!dragged) return;
    const strip=element.querySelector('.office-stack-strip'), offset=stripOffset(strip);
    if (e.type==='pointerup' && Math.abs(dx)>=28) step(dx<0 ? 1 : -1);
    element.style.removeProperty('--office-card-drag');
    if (e.type!=='pointerup' || Math.abs(dx)<28) slideCardStrip(strip,offset);
  }
  function click(e) {
    if (dragged && e.detail) { e.preventDefault(); e.stopPropagation(); }
    dragged=false;
  }
  function leave(e) { if (!dragged && press?.id===e.pointerId) press=undefined; }
  const listeners = [['pointerdown',down],['pointermove',move],['pointerup',finish],['pointercancel',finish],['lostpointercapture',finish],['pointerleave',leave],['click',click,true],['wheel',wheel,{passive:false}]];
  for (const [name,fn,capture] of listeners) element.addEventListener(name,fn,capture);
  return () => { for (const [name,fn,capture] of listeners) element.removeEventListener(name,fn,capture); };
}
