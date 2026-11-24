const ACTIVE = 'textarea[data-writer-composition-active="true"][data-writer-input-positioned="true"]';
const TEXT = 'text[data-writer-document-content]';
const properties = ['top', 'height', 'font', 'forward', 'span'];

// The browser owns marked text. Moving committed glyphs during composition changes
// the scrollable SVG bounds before Writer has a new paragraph layout.
export function createWriterComposition(root) {
  const view = root.ownerDocument.defaultView;
  let surface;
  function clear() {
    if (!surface) return;
    for (const name of properties) surface.style.removeProperty('--office-ime-' + name);
    surface = undefined;
  }
  function paint() {
    const input = root.querySelector(ACTIVE);
    if (!input?.value) { clear(); return; }
    if (surface !== input) { clear(); surface = input; }
    const style = view.getComputedStyle(input);
    const field = input.getBoundingClientRect();
    const origin = field.top - parseFloat(style.marginTop || 0);
    const lineHeight = parseFloat(style.getPropertyValue('--writer-input-caret-height'));
    // The visible field may extend backwards near the right margin. Its offset
    // parent and Writer's own transform still identify the real insertion point.
    const parent = input.offsetParent;
    if (!parent || !lineHeight) return;
    const viewport = input.closest('.writer-document-preview__scroll')?.getBoundingClientRect();
    const insertion = parent.getBoundingClientRect().left + input.offsetLeft + new view.DOMMatrixReadOnly(style.transform === 'none' ? undefined : style.transform).e;
    const middle = origin - lineHeight / 2;
    let left = Math.max(0, viewport?.left ?? 0), right = Math.min(view.innerWidth, viewport?.right ?? view.innerWidth);
    let nearest, best = Infinity;
    for (const page of root.querySelectorAll('svg[data-paint-key^="page:"]')) {
      const paper = page.getBoundingClientRect();
      if (middle < paper.top || middle > paper.bottom || insertion < paper.left - 1 || insertion > paper.right + 1) continue;
      const start = page.querySelector('[data-paint-key$=":corner:top-left"]')?.getBoundingClientRect();
      const end = page.querySelector('[data-paint-key$=":corner:top-right"]')?.getBoundingClientRect();
      left = Math.max(left, start && start.right <= insertion + 1 ? start.right : paper.left);
      right = Math.min(right, end && end.left >= insertion - 1 ? end.left : paper.right);
      for (const run of page.querySelectorAll(TEXT)) {
        const box = run.getBoundingClientRect();
        if (middle < box.top - 1 || middle > box.bottom + 1) continue;
        const dx = Math.max(box.left - insertion, insertion - box.right, 0);
        const score = dx + Math.abs(middle - (box.top + box.bottom) / 2);
        if (dx <= box.height && score < best) { best = score; nearest = {run, box}; }
      }
      break;
    }
    const set = (name, value) => {
      const prop = '--office-ime-' + name;
      if (surface.style.getPropertyValue(prop) !== value) surface.style.setProperty(prop, value);
    };
    set('forward', Math.max(0, right - insertion) + 'px');
    set('span', Math.max(2, right - left) + 'px');
    if (nearest) {
      set('top', nearest.box.top - origin + 'px');
      set('height', nearest.box.height + 'px');
      set('font', view.getComputedStyle(nearest.run).fontFamily);
    } else {
      for (const name of ['top', 'height', 'font']) surface.style.removeProperty('--office-ime-' + name);
    }
  }
  return {paint, clear};
}
