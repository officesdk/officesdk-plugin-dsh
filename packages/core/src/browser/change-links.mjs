import { createWriterChangeInk, paintWriterUnderline } from './writer-change-ink.mjs';
import { overlapsPreviousChange } from './change-history.mjs';

// Decorations leave pointer events to the editor for caret placement and drag selection.
export function createChangeLinks(document, window, locationsFor, reveal, onAnchors) {
  const layer=document.createElement('div');
  layer.className='office-change-links'; layer.setAttribute('aria-hidden','true');
  document.body.append(layer);
  const ink=createWriterChangeInk(document);
  const writerLines = [];
  let records, epoch=0, locations=new Map(), decorated=new Set(), hits=[], frame, press, disposed=false;
  let lastAnchors='', hadAnchors=false;
  let covered=new Map();
  const lastWriterAnchors = new Map();
  function draw() {
    frame=undefined;
    const next=[], anchors=[];
    for (const [id,locate] of locations) {
      const geometry=locate();
      if (!geometry) continue;
      const boxes=geometry.rects ?? [geometry];
      if (boxes.length && geometry.height>0) {
        // Send viewport CSS pixels, including offscreen rows. The rail is outside the editor iframe.
        const top=Math.round(geometry.top*2)/2, bottom=Math.round((geometry.top+geometry.height)*2)/2;
        anchors.push({id,top,bottom,lineHeight:Math.max(1,Math.round(boxes[0].height))});
      }
      if (!decorated.has(id)) continue;
      const clip=geometry.clip;
      for (const box of boxes) {
        if (box.inkNode) {
          next.push({ id, inkNode:box.inkNode, x:box.left, right:box.left+box.width });
          continue;
        }
        const x=Math.max(0,clip?.left ?? 0,box.left), y=Math.max(0,clip?.top ?? 0,box.top);
        const right=Math.min(window.innerWidth,clip?.right ?? Infinity,box.left+box.width);
        const bottom=box.top+box.height;
        if (right<=x || bottom<=y || bottom>Math.min(clip?.bottom ?? Infinity,window.innerHeight)) continue;
        next.push({id,x,y,right,bottom,inkNode:box.inkNode});
      }
    }
    let writerCount = 0, cellCount = 0;
    for (const box of next) {
      if (box.inkNode) {
        box.line = paintWriterUnderline(document, writerLines[writerCount], box);
        if (box.line) writerLines[writerCount++] = box.line;
        continue;
      }
      const {id,x,y,right,bottom}=box;
      const el=layer.children[cellCount++] ?? layer.appendChild(document.createElement('i'));
      el.dataset.change=id;
      el.style.cssText='left:'+x+'px;top:'+y+'px;width:'+(right-x)+'px;height:'+(bottom-y)+'px';
    }
    while (writerLines.length > writerCount) writerLines.pop().remove();
    while (layer.children.length>cellCount) layer.lastChild.remove();
    ink.paint(next);
    if (covered.size) {
      const direct=new Map(anchors.map(a=>[a.id,a]));
      for (const [id,source] of covered) {
        const anchor=direct.get(source);
        if (anchor) anchors.push({...anchor,id});
      }
    }
    // The Writer model and its SVG update in separate frames on a keystroke.
    // Keep each pending card on its last measured line until its new line is
    // painted. Only fresh geometry above is allowed to create clickable ink.
    if (lastWriterAnchors.size || records?.some(r => r.kind === 'docx')) {
      const scroll = document.querySelector('.writer-document-preview__scroll')?.scrollTop ?? 0;
      const measured = new Map(anchors.map(anchor => [anchor.id, anchor]));
      for (const record of records ?? []) {
        if (record.kind !== 'docx') continue;
        const anchor = measured.get(record.id);
        if (anchor) {
          lastWriterAnchors.set(record.id, {anchor, scroll});
        } else {
          const previous = lastWriterAnchors.get(record.id);
          if (previous) {
            const delta = previous.scroll - scroll;
            anchors.push({...previous.anchor, top:previous.anchor.top + delta, bottom:previous.anchor.bottom + delta});
          }
        }
      }
      anchors.sort((a,b) => a.top - b.top || a.id.localeCompare(b.id));
    }
    const signature=JSON.stringify([window.innerHeight,anchors]);
    if (signature!==lastAnchors) {
      lastAnchors=signature;
      hadAnchors=anchors.length>0;
      onAnchors?.({height:window.innerHeight,items:anchors});
    }
    hits=next;
    if(!next.length)document.documentElement.removeAttribute('data-office-over-change');
  }
  function schedule() {
    if (!disposed && (locations.size || hits.length || hadAnchors) && frame===undefined) frame=window.requestAnimationFrame(draw);
  }
  async function sync(list,force=false) {
    if (list===records && !force) return;
    records=list;
    if (lastWriterAnchors.size) {
      const retained = new Set(list.map(r => r.id));
      for (const id of lastWriterAnchors.keys()) if (!retained.has(id)) lastWriterAnchors.delete(id);
    }
    const generation=++epoch;
    try {
      const fresh=list.length?await locationsFor(list):new Map();
      if (generation===epoch && !disposed) {
        decorated=new Set(list.filter(r=>!r.accepted && r.revert?.original!=='').map(r=>r.id));
        covered=new Map();
        const recent=[...list].sort((a,b)=>(a.appliedAt??0)-(b.appliedAt??0));
        for (let i=recent.length-2;i>=0;i--) {
          const old=recent[i];
          if (fresh.has(old.id)) continue;
          const next=recent.slice(i+1).find(r=>(fresh.has(r.id)||covered.has(r.id)) && overlapsPreviousChange(old,r));
          if (next) covered.set(old.id,covered.get(next.id) ?? next.id);
        }
        locations=fresh;schedule();
      }
    } catch { if(generation===epoch) {locations.clear();covered.clear();schedule();} }
  }
  function reviewSurface(e) {
    return e.target?.closest?.('#editor-root') && !e.target.closest('.office-region,button,input,textarea,[role=toolbar],[role=tablist],[role=menu]');
  }
  function hit(e) {
    if (!reviewSurface(e)) return;
    return hits.findLast(b => {
      if (!b.inkNode) return e.clientX>=b.x && e.clientX<=b.right && e.clientY>=b.y && e.clientY<=b.bottom;
      if (!b.line?.isConnected || !b.inkNode.closest('.writer-document-preview__scroll')?.contains(e.target)) return false;
      const span = b.line.getBoundingClientRect(), glyph = b.inkNode.getBoundingClientRect();
      return e.clientX>=span.left && e.clientX<=span.right && e.clientY>=glyph.top && e.clientY<=glyph.bottom;
    })?.id;
  }
  const down=e=>{press=records?.length && e.button===0 && e.isPrimary!==false && reviewSurface(e)?{id:hit(e),x:e.clientX,y:e.clientY,pointer:e.pointerId}:null;};
  const move=e=>{
    if(press && Math.hypot(e.clientX-press.x,e.clientY-press.y)>4) press=null;
    document.documentElement.toggleAttribute('data-office-over-change',!e.buttons && !!hit(e));
  };
  const up=e=>{
    const start=press;press=null;
    if(start && start.pointer===e.pointerId && Math.hypot(e.clientX-start.x,e.clientY-start.y)<=4 && reviewSurface(e) && hit(e)===start.id) reveal(start.id ?? null);
  };
  const cancel=()=>{press=null;};
  const listeners=[['pointerdown',down],['pointermove',move],['pointerup',up],['pointercancel',cancel],['scroll',schedule]];
  for(const [name,fn] of listeners) document.addEventListener(name,fn,true);
  window.addEventListener('resize',schedule);
  return {sync,schedule,locationFor:id=>covered.get(id) ?? id,dispose(){
    disposed=true;++epoch;window.cancelAnimationFrame(frame);
    for(const [name,fn] of listeners) document.removeEventListener(name,fn,true);
    window.removeEventListener('resize',schedule);document.documentElement.removeAttribute('data-office-over-change');ink.dispose();writerLines.forEach(line=>line.remove());layer.remove();
  }};
}
