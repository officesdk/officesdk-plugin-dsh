// Writer paints text in SVG. Colour its existing glyphs, including just part of a run,
// without inserting text nodes or writing review formatting into the Word document.
const BLUE = 'var(--office-writer-selection-blue)';
const SVG = 'http://www.w3.org/2000/svg';

// The underline lives in the glyph's own SVG group. Composited scrolling and
// rubber-band movement then carry both in the same paint, even between JS frames.
export function paintWriterUnderline(document, line, { id, inkNode, x, right }) {
  const screen = inkNode.getBoundingClientRect(), local = inkNode.getBBox();
  if (!screen.width || !screen.height || !local.width) return;
  line ??= document.createElementNS(SVG, 'line');
  line.setAttribute('class', 'office-writer-change-line');
  line.setAttribute('aria-hidden', 'true');
  line.dataset.change = id;
  const toLocalX = value => local.x + (value - screen.left) * local.width / screen.width;
  const y = local.y + local.height - 0.5 * local.height / screen.height;
  line.setAttribute('x1', toLocalX(x));
  line.setAttribute('x2', toLocalX(right));
  line.setAttribute('y1', y);
  line.setAttribute('y2', y);
  const transform = inkNode.getAttribute('transform');
  if (line.getAttribute('transform') !== transform) {
    if (transform === null) line.removeAttribute('transform');
    else line.setAttribute('transform', transform);
  }
  if (line.parentNode !== inkNode.parentNode) inkNode.after(line);
  return line;
}

export function createWriterChangeInk(document) {
  const active = new Map();
  let defs, serial = 0;
  function clear(node, entry) {
    node.removeAttribute('data-office-change-ink');
    node.style.removeProperty('--office-change-paint');
    entry.gradient?.remove();
    active.delete(node);
  }
  function paint(boxes) {
    const groups = new Map();
    for (const {inkNode, x, right} of boxes) {
      if (!inkNode) continue;
      if (!groups.has(inkNode)) groups.set(inkNode, []);
      groups.get(inkNode).push([x,right]);
    }
    for (const [node,entry] of active) if (!groups.has(node)) clear(node,entry);
    for (const [node, ranges] of groups) {
      const bounds = node.getBoundingClientRect();
      if (!bounds.width) continue;
      let entry = active.get(node);
      if (!entry) {
        entry = {base:document.defaultView.getComputedStyle(node).fill};
        active.set(node,entry);
      }
      const spans = [];
      for (const [x,right] of ranges.sort((a,b) => a[0]-b[0])) {
        const start = Math.max(0,(x-bounds.left)/bounds.width);
        const end = Math.min(1,(right-bounds.left)/bounds.width);
        if (end <= start) continue;
        if (spans.length && start <= spans.at(-1)[1]+.00001) spans.at(-1)[1] = Math.max(spans.at(-1)[1],end);
        else spans.push([start,end]);
      }
      const base = node.getAttribute('fill') || entry.base;
      const signature = JSON.stringify([base,spans]);
      if (entry.signature === signature) continue;
      entry.signature = signature;
      let paint = BLUE;
      if (spans.length !== 1 || spans[0][0] > .00001 || spans[0][1] < .99999) {
        if (!defs) {
          defs = document.createElementNS(SVG,'svg');
          defs.setAttribute('width','0'); defs.setAttribute('height','0');
          defs.setAttribute('aria-hidden','true'); defs.style.position = 'absolute';
          document.body.append(defs);
        }
        if (!entry.gradient) {
          entry.gradient = document.createElementNS(SVG,'linearGradient');
          entry.gradient.id = 'officesdk-review-ink-'+(++serial);
          defs.append(entry.gradient);
        }
        const stops = [[0,base]];
        for (const [start,end] of spans) stops.push([start,base],[start,BLUE],[end,BLUE],[end,base]);
        stops.push([1,base]);
        entry.gradient.replaceChildren(...stops.map(([offset,color]) => {
          const stop = document.createElementNS(SVG,'stop');
          stop.setAttribute('offset',offset);
          stop.style.stopColor = color;
          return stop;
        }));
        paint = 'url(#'+entry.gradient.id+')';
      }
      node.style.setProperty('--office-change-paint',paint);
      node.setAttribute('data-office-change-ink','');
    }
  }
  return {paint,dispose() {
    for (const [node,entry] of active) clear(node,entry);
    defs?.remove();
  }};
}
