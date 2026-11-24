// Keep the original paragraph in Writer's flow while new glyphs still fit its reserved lines.
export function createWriterInlinePreview(svg) {
  const document = svg.ownerDocument, ns = svg.namespaceURI;
  const layer = document.createElementNS(ns, 'g');
  layer.setAttribute('data-office-write-preview', '');
  layer.style.pointerEvents = 'none';
  svg.append(layer);
  const measure = document.createElement('canvas').getContext('2d');
  const segmenter = new Intl.Segmenter(undefined, {granularity:'grapheme'});
  let previous;
  return (found, text) => {
    if (text === undefined) {
      layer.style.display = 'none';
      previous = undefined;
      return false;
    }
    const rows = [];
    for (const rect of found.rects) {
      const run = rect.run, transform = run?.getScreenCTM();
      if (!transform || !run.getNumberOfChars() || !measure) return false;
      const baseline = run.getStartPositionOfChar(0).matrixTransform(transform).y - found.top;
      const x = (rect.inkLeft ?? rect.left) - found.left;
      const right = rect.left + rect.width - found.left;
      const last = rows.at(-1);
      if (last && Math.abs(last.baseline - baseline) < 1) {
        last.right = Math.max(last.right, right);
        continue;
      }
      const css = document.defaultView.getComputedStyle(run);
      const font = css.fontStyle+' '+css.fontWeight+' '+(parseFloat(css.fontSize)*transform.d)+'px '+css.fontFamily;
      rows.push({x, right, baseline, font, color:css.fill});
    }
    const signature = JSON.stringify([text,rows]);
    if (previous?.signature === signature) return previous.fits;
    const chars = Array.from(segmenter.segment(text), item => item.segment);
    const pieces = [];
    let offset = 0, needsRow = false;
    for (const row of rows) {
      needsRow = false;
      measure.font = row.font;
      let line = '';
      while (offset < chars.length) {
        const char = chars[offset];
        if (char === '\n') { offset++; needsRow = true; break; }
        if (measure.measureText(line + char).width > row.right - row.x + .5) break;
        line += char;
        offset++;
      }
      pieces.push(line);
      if (offset === chars.length && !needsRow) break;
    }
    const fits = rows.length > 0 && offset === chars.length && !needsRow;
    previous = {signature, fits};
    layer.style.display = fits ? '' : 'none';
    if (!fits) return false;
    for (const [i, line] of pieces.entries()) {
      const row = rows[i], el = layer.children[i] ?? document.createElementNS(ns, 'text');
      el.setAttribute('x', row.x); el.setAttribute('y', row.baseline);
      el.setAttribute('fill',row.color); el.style.font = row.font; el.style.whiteSpace = 'pre';
      el.textContent = line;
      if (!el.parentNode) layer.append(el);
    }
    while (layer.children.length > pieces.length) layer.lastChild.remove();
    return true;
  };
}
