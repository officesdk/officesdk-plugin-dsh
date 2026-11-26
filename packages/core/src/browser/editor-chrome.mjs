import {officeIcon,officeWebsite} from './office-brand.mjs';
import { createWriterSelectionSizer } from './writer-selection-highlight.mjs';
import { createWriterComposition } from './writer-composition.mjs';
const STATUS = '.editor-statusbar,[data-role="statusbar"],.sm-sheet-footer,.ppt-shell__statusbar';
const LAYOUT_MENU = '[data-role="toolbar-ribbon-layout-switch-menu"]';
const LAYOUT_OPTION = '[data-role="toolbar-ribbon-layout-switch-menu-layout-option"]';
const NATIVE_MORE = '[data-role="toolbar-extra-menu"]';
const CSS = `
.office-website-link{display:inline-flex;align-items:center;gap:6px;flex:none;align-self:center;white-space:nowrap;font:normal 12px/20px -apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif;color:#50555b;text-decoration:none;border-radius:3px}
.office-website-link img{display:block;flex:none}
.office-website-link:hover,.office-website-link:focus-visible{color:#202428;text-decoration:underline;text-underline-offset:2px}
.office-website-link:focus-visible{outline:2px solid currentColor;outline-offset:2px}
.toolbar-ribbon-header:has(>.office-website-link) .toolbar-ribbon-header__tabs{min-width:0;flex:1}

#editor-root[data-office-v31-writer] .writer-document-preview__scroll{overscroll-behavior-y:none}
/* Meeting template: keep the specified PingFang for bold SVG text instead of the renderer’s Heiti substitution. */
#editor-root[data-meeting-font-fix] text[data-writer-document-content][font-family*="Heiti SC"][font-family*="PingFang SC"]{font-family:"PingFang SC",sans-serif!important}
#editor-root[data-office-v31-writer] rect[data-paint-key^="overlay:selection:"][style*="--office-selection-y"]{y:var(--office-selection-y);height:var(--office-selection-height)}
#editor-root[data-office-v31-writer] [data-paint-key^="overlay:caret:"]:is(rect,path),#editor-root[data-office-v31-writer] [data-paint-key^="overlay:caret:"] :is(rect,path){fill:#202124!important}
#editor-root[data-office-v31-writer] rect[data-paint-key^="overlay:caret:"][style*="--office-caret-top"]{y:var(--office-caret-top);height:var(--office-caret-height);width:var(--office-caret-width)}
#editor-root[data-office-v31-writer] textarea[data-writer-composition-active=true]{max-width:var(--office-ime-span,2px);translate:min(0px,calc(var(--office-ime-forward,2px) - 100%)) 0;white-space:pre;overflow:hidden}
#editor-root[data-office-v31-writer] textarea[data-writer-composition-active=true][style*="--office-ime-height"]{margin-top:var(--office-ime-top);height:var(--office-ime-height);line-height:var(--office-ime-height);font-family:var(--office-ime-font)!important}
#editor-root[data-office-v31-writer]:has(textarea[data-writer-composition-active=true]) [data-paint-key^="overlay:caret:"]{visibility:hidden}

#editor-root[data-office-v31-writer] :is(.toolbar-ribbon-header__mode-button,.toolbar-ribbon-header__comment-button,.writer-quick-comment-button){display:none!important}
.office-v31-layout-sync :is([data-role="toolbar-extra-menu-panel"],[data-role="toolbar-ribbon-layout-switch-menu"]){visibility:hidden!important}
`;

// In fit-whole-page mode Writer places page one flush with the viewport. Its
// scroll origin includes a leading gutter, which the native wheel can enter
// before Writer snaps back. Stop an upward gesture at the actual paper edge.
function installWriterScrollBoundary(root, document) {
  let wholePage = false;
  const changeZoom = event => {
    const control = event.target.closest?.('[data-menu-binding-token],button,.zoom-dropdown__input');
    if (!control) return;
    const token = control.getAttribute('data-menu-binding-token');
    if (token === 'root:fit-whole-page') wholePage = true;
    else if (token === 'root:fit-page-width' || token?.startsWith('root:preset-') ||
        control.matches('.zoom-dropdown__input') || control.closest('.zoom-dropdown__input') ||
        control.previousElementSibling?.matches('.editor-statusbar__zoom-value') ||
        control.nextElementSibling?.matches('.editor-statusbar__zoom-value')) wholePage = false;
  };
  const wheel = event => {
    if (event.ctrlKey || event.metaKey) { wholePage = false; return; }
    if (!wholePage || !event.cancelable || event.shiftKey || event.deltaY >= 0 || Math.abs(event.deltaX) > -event.deltaY) return;
    const scroll = event.target.closest?.('.writer-document-preview__scroll');
    const page = scroll?.querySelector('svg[data-paint-key="page:1"]');
    if (!page) return;
    const viewport = scroll.getBoundingClientRect(), paper = page.getBoundingClientRect();
    const distance = viewport.top - paper.top, edge = Math.round(scroll.scrollTop - distance);
    const unit = event.deltaMode === 1 ? (parseFloat(document.defaultView.getComputedStyle(scroll).lineHeight) || 16)
      : event.deltaMode === 2 ? scroll.clientHeight : 1;
    if (!paper.height || paper.height > viewport.height + 1 || edge <= 0 || distance > -event.deltaY * unit + 1) return;
    event.preventDefault();
    if (scroll.scrollTop > edge) scroll.scrollTop = edge;
    if (event.deltaX) scroll.scrollLeft += event.deltaX * unit;
  };
  document.addEventListener('click', changeZoom, true);
  document.addEventListener('input', changeZoom, true);
  root.addEventListener('wheel', wheel, { capture:true, passive:false });
  return () => {
    document.removeEventListener('click', changeZoom, true);
    document.removeEventListener('input', changeZoom, true);
    root.removeEventListener('wheel', wheel, true);
  };
}

export function installEditorChrome({ root, window, document, suite, locale, signal, onWebsite, onStatusInset, onGeometry }) {
  const style = document.createElement('style');
  style.textContent = CSS;
  document.head.append(style);
  if (suite === 'document') root.setAttribute('data-office-v31-writer', '');

  const removeScrollBoundary = suite === 'document' ? installWriterScrollBoundary(root, document) : undefined;
  const sizeSelection = suite === 'document' ? createWriterSelectionSizer(root) : undefined;
  const composition = suite === 'document' ? createWriterComposition(root) : undefined;
  const brand = document.createElement('a');
  brand.className = 'office-website-link';
  brand.href = officeWebsite;
  brand.target = '_blank';
  brand.rel = 'noopener noreferrer';
  brand.title = locale === 'en' ? 'Open the OfficeSDK website in a new tab' : '在新页签打开 OfficeSDK 官网';
  const icon = Object.assign(document.createElement('img'), {src:officeIcon, alt:'', width:14, height:14});
  brand.append(icon, 'officesdk.com');
  // The editor frame is sandboxed. Let the parent open only its own fixed homepage.
  function openWebsite(event) {
    if (!onWebsite || (event.type === 'auxclick' && event.button !== 1)) return;
    event.preventDefault();
    onWebsite();
  }
  brand.addEventListener('click', openWebsite);
  brand.addEventListener('auxclick', openWebsite);
  let status, lastInset;
  const updateBottom = () => {
    const box = status?.getBoundingClientRect();
    const inset = box?.height && box.top > window.innerHeight / 2 ? window.innerHeight - box.top : 0;
    const bottom = Math.ceil(inset + 14);
    if (lastInset !== bottom) { lastInset = bottom; onStatusInset?.(bottom); }
  };
  const resize = new window.ResizeObserver(onGeometryChange);
  resize.observe(root);
  function sizeCaret() {
    for (const caret of root.querySelectorAll('rect[data-paint-key^="overlay:caret:"]')) {
      const m = caret.getScreenCTM();
      if (!m || m.a <= 0 || m.d <= 0 || m.b || m.c) continue;
      // The SDK owns the insertion column. Read its unmodified line box and use the
      // adjacent text run's font box for the vertical position, independent of leading.
      const x = Number(caret.getAttribute('x')) * m.a + m.e;
      const top = Number(caret.getAttribute('y')) * m.d + m.f;
      const height = Number(caret.getAttribute('height')) * m.d;
      let nearest, distance = Infinity;
      for (const run of caret.ownerSVGElement.querySelectorAll('text[data-writer-document-content]')) {
        const box = run.getBoundingClientRect();
        const center = (box.top + box.bottom) / 2;
        if (!box.height || center < top || center > top + height) continue;
        const dx = Math.max(box.left - x, x - box.right, 0);
        if (dx > box.height) continue;
        const score = dx + Math.abs(center - (top + height / 2));
        if (score < distance) { nearest = box; distance = score; }
      }
      if (nearest) {
        caret.style.setProperty('--office-caret-top', (nearest.top - m.f) / m.d + 'px');
        caret.style.setProperty('--office-caret-height', nearest.height / m.d + 'px');
        caret.style.setProperty('--office-caret-width', 1 / m.a + 'px');
      } else {
        caret.style.removeProperty('--office-caret-top');
        caret.style.removeProperty('--office-caret-height');
        caret.style.removeProperty('--office-caret-width');
      }
    }
  }
  const storageKey = `officesdk:v3.1:ribbon:${suite}`;
  let layout = 'single-line';
  try { if (window.localStorage.getItem(storageKey) === 'double-line') layout = 'double-line'; } catch {}

  let disposed = false;
  let pending;
  let applied;
  let scheduled;
  let syncing;
  let initialTrigger;
  const frame = () => new Promise(resolve => window.requestAnimationFrame(resolve));
  let hostFocus = false;
  const fullscreen = () => hostFocus || document.fullscreenElement === document.documentElement;

  async function setRibbon(desired) {
    pending = desired;
    if (syncing || applied === desired) return;
    const trigger = root.querySelector(NATIVE_MORE);
    if (!trigger) return;
    syncing = true;
    document.documentElement.classList.add('office-v31-layout-sync');
    let openedHere = false;
    try {
      if (!document.querySelector(LAYOUT_MENU)) {
        trigger.click();
        openedHere = true;
      }
      for (let turn = 0; turn < 24 && !disposed; turn++) {
        const menu = document.querySelector(LAYOUT_MENU);
        const option = menu?.querySelector(`${LAYOUT_OPTION}[data-layout="${desired}"]`);
        if (option) {
          option.click();
          applied = desired;
          openedHere = false;
          break;
        }
        await frame();
      }
      if (openedHere && !disposed && trigger.getAttribute('aria-expanded') === 'true') trigger.click();
    } finally {
      syncing = false;
      document.documentElement.classList.remove('office-v31-layout-sync');
      if (!disposed && pending !== desired) void setRibbon(pending);
    }
  }

  function mount() {
    scheduled = undefined;
    const header = root.querySelector('.toolbar-ribbon-header');
    if (header && brand.parentNode !== header) header.appendChild(brand);
    const current = root.querySelector(STATUS);
    if (current !== status) {
      if (status) resize.unobserve(status);
      status = current;
      if (status) resize.observe(status);
    }
    updateBottom();
    if (suite === 'document') sizeCaret();
    composition?.paint();
    const trigger = root.querySelector(NATIVE_MORE);
    if (trigger && trigger !== initialTrigger && !syncing) {
      initialTrigger = trigger;
      applied = undefined;
      void setRibbon(fullscreen() ? 'double-line' : layout);
    }
  }

  function onLayoutClick(event) {
    const option = event.target.closest?.(LAYOUT_OPTION);
    if (syncing || !option || !['single-line', 'double-line'].includes(option.dataset.layout)) return;
    applied = option.dataset.layout;
    if (fullscreen()) return;
    layout = applied;
    try { window.localStorage.setItem(storageKey, layout); } catch {}
  }
  function onFullscreen() {
    void setRibbon(fullscreen() ? 'double-line' : layout);
  }
  function scheduleMount() {
    if (!disposed && scheduled === undefined) scheduled = window.requestAnimationFrame(mount);
  }
  // Writer can replace its SVG selection inside an animation frame. Apply the band
  // in the mutation microtask, before that frame paints at the native line height.
  function onGeometryChange() {
    onGeometry?.();
    sizeSelection?.();
    scheduleMount();
  }
  const observer = new window.MutationObserver(onGeometryChange);
  function dispose() {
    if (disposed) return;
    disposed = true;
    observer.disconnect();
    brand.remove();
    composition?.clear();
    removeScrollBoundary?.();
    resize.disconnect();
    document.fonts?.removeEventListener('loadingdone',onGeometryChange);
    document.documentElement.style.removeProperty('--office-dock-bottom');
    if (scheduled !== undefined) window.cancelAnimationFrame(scheduled);
    signal?.removeEventListener('abort', dispose);
    document.removeEventListener('click', onLayoutClick, true);
    document.removeEventListener('fullscreenchange', onFullscreen);
    document.documentElement.classList.remove('office-v31-layout-sync');
    root.removeAttribute('data-office-v31-writer');
    root.removeAttribute('data-meeting-font-fix');
    root.removeEventListener('input', scheduleMount, true);
    root.removeEventListener('compositionupdate', scheduleMount, true);
    root.removeEventListener('compositionend', scheduleMount, true);
    style.remove();
  }
  signal?.addEventListener('abort', dispose, { once: true });
  document.addEventListener('click', onLayoutClick, true);
  document.addEventListener('fullscreenchange', onFullscreen);
  document.fonts?.addEventListener('loadingdone',onGeometryChange);
  if (composition) {
    root.addEventListener('input', scheduleMount, true);
    root.addEventListener('compositionupdate', scheduleMount, true);
    root.addEventListener('compositionend', scheduleMount, true);
  }
  observer.observe(root, { childList: true, subtree: true, attributes: true, characterData: true, attributeFilter: ['x','y','width','height','transform','viewBox','d','font-size','font-family','font-weight','data-writer-composition-active','data-writer-input-positioned'] });
  sizeSelection?.();
  mount();
  return { dispose, focusMode(active) { hostFocus = active; onFullscreen(); } };
}
