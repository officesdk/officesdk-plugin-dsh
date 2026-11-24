// Writer's SVG supplies the exact horizontal selection and the line's full leading.
// Use one vertical band per rendered line, independent of where the drag starts or ends.
const SELECTION = 'rect[data-paint-key^="overlay:selection:"]';
const TEXT = 'text[data-writer-document-content]';
const BAND_TO_FONT_HEIGHT = 1.65;

function selectionBand(top,bottom,lineHeight) {
  const height=Math.min(lineHeight,(bottom-top)*BAND_TO_FONT_HEIGHT);
  return {top:(top+bottom-height)/2,height};
}

// Keep the SDK's horizontal text extents. Prefer the actual manual band while it exists;
// once typing replaces it, reconstruct the same band from this line's rendered fonts and leading.
export function writerWorkBands(document,elements,rects) {
  const native=Array.from(document.querySelectorAll(SELECTION),r=>r.getBoundingClientRect()).filter(r=>r.height && r.width);
  let fonts;
  return rects.map(rect=>{
    const center=rect.top+rect.height/2;
    const band=native.find(r=>center>=r.top && center<r.bottom && r.left<rect.left+rect.width && r.right>rect.left);
    if (band) return {...rect,top:band.top,height:band.height};
    fonts ??= elements.map(el=>({box:el.getBoundingClientRect(),leading:Number(el.getAttribute('height'))*(el.getScreenCTM()?.d ?? 1)}));
    const row=fonts.filter(({box})=>Math.min(box.bottom,rect.top+rect.height)-Math.max(box.top,rect.top)>Math.min(box.height,rect.height)*.6);
    if (!row.length) return rect;
    const top=Math.min(...row.map(r=>r.box.top)), bottom=Math.max(...row.map(r=>r.box.bottom));
    const leading=Math.max(...row.map(r=>r.leading));
    return {...rect,...selectionBand(top,bottom,leading>0?leading:Infinity)};
  });
}

export function createWriterSelectionSizer(root) {
  return () => {
    const pages = new Map();
    for (const rect of root.querySelectorAll(SELECTION)) {
      const svg = rect.ownerSVGElement, m = rect.getScreenCTM();
      if (!svg || !m || m.a <= 0 || m.d <= 0 || Math.abs(m.b) + Math.abs(m.c) > .001) continue;
      // Read the SDK attributes so our CSS adjustment never feeds back into the next drag frame.
      const y = Number(rect.getAttribute('y')), h = Number(rect.getAttribute('height'));
      const lineTop = y*m.d+m.f, lineBottom = lineTop+h*m.d;
      // Use SVG transforms for both the band and the glyph. DOM client rectangles
      // can include an elastic scroll offset from a different compositor frame.
      if (!pages.has(svg)) {
        const rows = [];
        for (const run of svg.querySelectorAll(TEXT)) {
          const transform = run.getScreenCTM();
          if (!transform || transform.d <= 0 || transform.b || transform.c) continue;
          const glyph = run.getBBox();
          const top = glyph.y * transform.d + transform.f, height = glyph.height * transform.d;
          rows.push({ top, bottom:top+height, height });
        }
        pages.set(svg, rows);
      }
      let top = Infinity, bottom = -Infinity;
      for (const box of pages.get(svg)) {
        const center = (box.top+box.bottom)/2;
        if (!box.height || center < lineTop || center >= lineBottom) continue;
        top = Math.min(top,box.top);
        bottom = Math.max(bottom,box.bottom);
      }
      if (!(bottom>top)) {
        rect.style.removeProperty('--office-selection-y');
        rect.style.removeProperty('--office-selection-height');
        continue;
      }
      const band = selectionBand(top,bottom,h*m.d);
      const height = band.height;
      const localY = (band.top-m.f)/m.d;
      rect.style.setProperty('--office-selection-y',localY+'px');
      rect.style.setProperty('--office-selection-height',height/m.d+'px');
    }
  };
}

/** Keep the chip and its inline input outside the visible blue band, including the padding above the glyphs. */
export function selectionHighlightBounds(document, glyphs) {
  if (!glyphs) return;
  const bands = Array.from(document.querySelectorAll(SELECTION), r=>r.getBoundingClientRect())
    .filter(r => r.width && r.height && r.left < glyphs.left+glyphs.width && r.right > glyphs.left &&
      r.top < glyphs.top+glyphs.height && r.bottom > glyphs.top);
  if (!bands.length) return glyphs;
  const top = Math.min(...bands.map(r=>r.top)), bottom = Math.max(...bands.map(r=>r.bottom));
  return { ...glyphs, top, height:bottom-top };
}
