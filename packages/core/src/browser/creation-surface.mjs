// The waiting marquee has no document content. Its locator measures the live paper or SDK range on each paint.
export function creationBands(bounds, rows = 1) {
  if (!bounds || bounds.width <= 0 || bounds.height <= 0) return;
  const {left,top,width,height,clip} = bounds;
  const count = Math.max(1,Math.min(rows,32, Math.floor(height / 8)));
  const line = height / count;
  return {left,top,width,height,clip,rects:Array.from({length:count},(_,i)=>({left,top:top+i*line,width,height:line}))};
}

export function writerPaper(document) {
  return Array.from(document.querySelectorAll('.writer-document-preview__scroll svg[data-paint-key="page:1"]')).find(el=>el.getBoundingClientRect().width);
}

export function writerOpening(document, page = writerPaper(document)) {
  if (!page) return;
  const paper = page.getBoundingClientRect();
  const left = page.querySelector('path[data-paint-key$=":corner:top-left"]')?.getBoundingClientRect();
  const right = page.querySelector('path[data-paint-key$=":corner:top-right"]')?.getBoundingClientRect();
  if (!left || !right) return;
  const scale = paper.width / (page.viewBox?.baseVal?.width || paper.width);
  return {left:left.right,top:left.bottom,width:right.left-left.right,height:24*scale,
    clip:page.closest('.writer-document-preview__scroll')?.getBoundingClientRect()};
}

export function createGenerationSurface(overlay, locale) {
  const key = 'new-file-area';
  let active = false;
  return {
    get active() { return active; },
    async start(find) {
      active = true;
      await overlay.select(key,find,{edge:true,writer:true,persistent:true,
        label:locale==='en'?'Agent · Preparing':'Agent · 准备写入',selecting:locale==='en'?'Agent · Selecting':'Agent · 框选中'});
      if (active) await overlay.phase(key,'placeholder');
    },
    writing() { if (active) overlay.drop(key); },
    finish() {
      if (!active) return false;
      overlay.drop(key);
      active = false;
      return true;
    },
  };
}
