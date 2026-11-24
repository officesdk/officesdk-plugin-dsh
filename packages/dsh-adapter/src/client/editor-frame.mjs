/** Keep an open file's browsing context alive when DSH unmounts its inactive tab body. */
export function createEditorFrame(container, { printable, title }) {
  const doc = container.ownerDocument;
  const frame = doc.createElement('iframe');
  frame.title = title;
  frame.allowFullscreen = true;
  frame.referrerPolicy = 'no-referrer';
  frame.setAttribute('sandbox', 'allow-scripts allow-same-origin' + (printable ? ' allow-modals' : ''));
  frame.style.cssText = 'display:block;border:0;width:100%;height:100%';
  container.appendChild(frame);
  let parking;
  let lastSize = container.getBoundingClientRect();
  const sizes = new ResizeObserver(entries => {
    const size = entries[0].contentRect;
    if (size.width && size.height) lastSize = size;
  });
  sizes.observe(container);
  return {
    frame,
    update(view) {
      const allow = "clipboard-read 'none'; clipboard-write " + (view.mode === 'edit' ? "'self'" : "'none'");
      if (frame.allow !== allow) frame.allow = allow;
      if (view.url && frame.getAttribute('src') !== view.url) frame.src = view.url;
      if (!view.url && frame.hasAttribute('src')) frame.removeAttribute('src');
    },
    park() {
      if (!frame.isConnected || typeof doc.body.moveBefore !== 'function') return false;
      sizes.disconnect();
      parking = doc.createElement('div');
      parking.inert = true;
      parking.setAttribute('aria-hidden', 'true');
      parking.style.cssText = 'position:fixed;left:0;top:0;opacity:0;pointer-events:none;z-index:-1;overflow:hidden';
      parking.style.width = Math.max(1, lastSize.width) + 'px';
      parking.style.height = Math.max(1, lastSize.height) + 'px';
      doc.body.appendChild(parking);
      // appendChild would navigate the iframe again and discard pending SDK operations.
      parking.moveBefore(frame, null);
      return true;
    },
    attach(next) {
      next.moveBefore(frame, null);
      sizes.observe(next);
      parking?.remove();
      parking = undefined;
    },
    dispose() { sizes.disconnect(); frame.remove(); parking?.remove(); },
  };
}
