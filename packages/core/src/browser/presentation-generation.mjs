// The placeholders live on the sorter's thumbnail elements, outside the presentation model.
import { SORTER, viewSwitch } from './presentation-agent.mjs';
import { agentLabel } from './agent-status.mjs';
import { selectionMs } from './selection.mjs';

const css = `
[data-thumbnail-slide-id]:has(> .office-slide-generation){position:relative}
.office-slide-generation{position:absolute;inset:0;pointer-events:none;z-index:2;isolation:isolate;container-type:inline-size}
.office-slide-generation::after{content:"";position:absolute;inset:0;border:2px solid transparent;border-radius:3px;background:linear-gradient(120deg,#65a4ff,#b485f2,#74d8df,#65a4ff) border-box;background-size:220% 100%;mask:linear-gradient(#000 0 0) padding-box,linear-gradient(#000 0 0);mask-composite:exclude;animation:office-slide-glow 2.8s ease-in-out infinite alternate;animation-delay:var(--phase);filter:hue-rotate(var(--hue))}
.office-slide-generation [data-placeholder]{position:absolute;inset:2px;background:#fff;overflow:hidden}
.office-slide-generation i{position:absolute;background:#e6eaf0;border-radius:2px;animation:office-slide-breathe 1.6s ease-in-out infinite alternate;animation-delay:var(--phase)}
.office-slide-generation [data-agent]{position:absolute;right:3%;bottom:4%;z-index:1;display:flex;align-items:center;gap:2px;color:#fff;font:500 clamp(8px,4cqw,12px)/1.3 system-ui;white-space:nowrap}
.office-slide-generation [data-agent]::before{content:"";align-self:flex-start;width:1em;height:1.15em;background:#292d34;clip-path:polygon(0 0,100% 35%,53% 55%,35% 100%);filter:drop-shadow(0 0 1px white)}
.office-slide-generation span{padding:.2em .55em;background:#292d34;border-radius:99px}
.office-slide-generation[data-phase=creating] [data-placeholder] i,.office-slide-generation[data-phase=writing] [data-placeholder],.office-slide-generation[data-phase=partial] [data-placeholder]{display:none}
.office-slide-generation[data-phase=creating]::after{animation:office-slide-appear .28s ease-out both}
.office-slide-generation[data-phase=creating] [data-agent]{animation:office-slide-cursor .28s cubic-bezier(.16,.74,.27,1) both}
.office-slide-generation[data-phase=writing]::after{display:none}
[data-thumbnail-slide-id]:has(.office-slide-generation[data-phase=writing]) :is([data-mop-svg-placeholder-prompt],[data-mop-svg-placeholder-border]){visibility:hidden}
.office-slide-generation [data-module]{position:absolute;pointer-events:none}
.office-slide-generation [data-marquee]{position:absolute;top:0;left:0;width:100%;height:100%;box-sizing:border-box;border:1.5px solid transparent;border-radius:2px;background:linear-gradient(135deg,#5799ff,#bc82e8,#57cad1) border-box;mask:linear-gradient(#000 0 0) padding-box,linear-gradient(#000 0 0);mask-composite:exclude}
.office-slide-generation [data-module] [data-agent]{right:auto;bottom:auto;left:100%;top:100%}
.office-slide-generation [data-module][data-flip-x] span{position:absolute;right:100%;margin-right:2px}
.office-slide-generation [data-module][data-flip-y] [data-agent]{transform:translateY(-100%)}
.office-slide-generation [data-module][data-flip-y] [data-agent]::before{transform:scaleY(-1)}
.office-slide-generation [data-module][data-pulling] [data-agent]{animation:office-module-cursor var(--pull-ms) cubic-bezier(.16,.74,.27,1) var(--elapsed) both}
.office-slide-generation [data-module][data-pulling] [data-marquee]{animation:office-module-select var(--pull-ms) cubic-bezier(.16,.74,.27,1) var(--elapsed) both}
@keyframes office-module-cursor{from{left:0;top:0}}
@keyframes office-module-select{from{width:0;height:0}}
@keyframes office-slide-cursor{from{right:100%;bottom:100%;transform:translate(100%,100%)}}
@keyframes office-slide-appear{from{clip-path:inset(0 96% 96% 0)}}
@keyframes office-slide-glow{to{background-position:100% 50%}}
@keyframes office-slide-breathe{to{opacity:.43}}
@media(prefers-reduced-motion:reduce){.office-slide-generation,.office-slide-generation::after,.office-slide-generation i,.office-slide-generation [data-agent],.office-slide-generation [data-marquee]{animation:none!important}}
`;

const layouts = [
  [[9,30,58,12],[9,49,42,5],[9,60,29,5]],
  [[8,11,46,8],[8,33,40,44],[54,33,37,5],[54,45,30,5],[54,57,35,5]],
  [[8,12,48,8],[8,32,24,39],[38,32,24,39],[68,32,24,39]],
  [[8,13,40,8],[8,35,33,5],[8,49,29,5],[8,63,34,5],[50,29,42,45]],
  [[8,12,48,8],[10,62,13,14],[30,49,13,27],[50,40,13,36],[70,31,13,45]],
  [[9,16,64,10],[9,39,80,5],[9,52,68,5],[9,65,74,5]],
];

export function showViewTab(document) {
  const tab = document.querySelector('[role=tab][data-tab-key="View"]');
  if (tab && tab.getAttribute('aria-selected') !== 'true') tab.click();
}

export async function switchDeckView(api, document, mode) {
  // Some Kit releases expose setViewMode before their product actions are bound.
  try { const set = viewSwitch(api); if (!set) throw new Error(); await set(mode); return; } catch {}
  const item = mode === SORTER ? 'view-mode-slide-sorter' : 'view-mode-normal';
  const find = () => [...document.querySelectorAll('[data-item-id="'+item+'"]')].find(el => el.getBoundingClientRect().width && !el.disabled);
  let button = find();
  if (!button) {
    const tab = [...document.querySelectorAll('[role=tab],button')].find(el => /^(视图|View)$/i.test(el.textContent.trim()) && el.getBoundingClientRect().width);
    tab?.click();
    await new Promise(resolve => document.defaultView.requestAnimationFrame(resolve));
    button = find();
  }
  if (!button) throw new Error('The presentation view control is unavailable.');
  button.click();
}

// Store thumbnail-relative coordinates so the outline stays on its module when the pane resizes.
// Template placeholders often report zero dimensions in the API. Their thumbnail SVG has the
// actual shape outline, including the position inherited from the slide layout.
function moduleBounds(thumb, target, slideSize) {
  if (target.whole) return [0,0,1,1];
  if (target.shapeId) {
    const shape = thumb.querySelector('[data-mop-move-preview-id="' + CSS.escape(target.shapeId) + '"]');
    const box = shape?.getBoundingClientRect();
    const page = thumb.getBoundingClientRect();
    if (box?.width && box.height && page.width && page.height) {
      return [(box.left-page.left)/page.width, (box.top-page.top)/page.height, box.width/page.width, box.height/page.height];
    }
  }
  const g = target.geometry;
  if (g?.width > 0 && g.height > 0 && slideSize?.width && slideSize.height) {
    return [g.left/slideSize.width,g.top/slideSize.height,g.width/slideSize.width,g.height/slideSize.height];
  }
}

export function deckProgress(document, slideIds, locale, slideSize) {
  const view = document.defaultView;
  const pages = new Map(slideIds.map((id, index) => [id, { index, phase: 'creating' }]));
  const drawnTargets = new WeakMap();
  const style = document.createElement('style');
  style.textContent = css;
  document.head.append(style);
  let frame, mounted, disposed = false;
  let following = true, activeId, pendingId, resume;
  const visibleThumb = id => [...document.querySelectorAll('[data-thumbnail-slide-id="' + CSS.escape(id) + '"]')]
    .find(el => !el.closest('[inert]') && el.getBoundingClientRect().width > 0);
  function reveal() {
    if (!following || !pendingId) return;
    const cover = pendingId === slideIds[0];
    const thumb = visibleThumb(pendingId);
    if (cover) {
      // The first row may not be mounted until its scroll container is reset.
      const mounted = thumb ?? [...document.querySelectorAll('[data-thumbnail-slide-id]')]
        .find(el => !el.closest('[inert]') && el.getBoundingClientRect().width > 0);
      for (let parent = mounted?.parentElement;parent;parent = parent.parentElement) {
        if (parent.scrollHeight > parent.clientHeight && /auto|scroll/.test(view.getComputedStyle(parent).overflowY)) {
          parent.scrollTop = 0;
          if (thumb) pendingId = undefined;
          return;
        }
      }
    }
    if (!thumb) return;
    pendingId = undefined;
    thumb.scrollIntoView({block: 'nearest', inline: 'nearest', behavior: 'instant'});
  }
  function track(id) {
    if (!id || id === activeId) return;
    activeId = pendingId = id;
    reveal();
  }
  function pauseFollowing(event) {
    if (!activeId || !event.target?.closest?.('[data-thumbnail-slide-id]') && !visibleThumb(activeId)?.parentElement?.contains(event.target)) return;
    following = false;
    if (resume) return;
    resume = document.createElement('button');
    resume.textContent = locale === 'zh' ? '回到正在生成的页' : 'Follow current slide';
    resume.style.cssText = 'position:fixed;bottom:40px;right:16px;z-index:5;padding:5px 10px;border:1px solid #ccc;border-radius:6px;background:white;color:#222';
    resume.onclick = () => { following = true; pendingId = activeId; reveal(); resume.remove(); resume = undefined; };
    document.body.append(resume);
  }
  for (const event of ['wheel','touchmove']) document.addEventListener(event, pauseFollowing, {passive:true});
  const firstThumbnail = new Promise(resolve => { mounted = resolve; });
  function draw(slideId) {
    if (disposed) return;
    reveal();
    const selector = slideId ? '[data-thumbnail-slide-id="' + CSS.escape(slideId) + '"]' : '[data-thumbnail-slide-id]';
    for (const thumb of document.querySelectorAll(selector)) {
      const id = thumb.getAttribute('data-thumbnail-slide-id');
      const page = pages.get(id);
      let layer = thumb.querySelector(':scope > .office-slide-generation');
      if (!page || page.phase === 'done') { layer?.remove(); continue; }
      if (!layer) {
        layer = document.createElement('div');
        layer.className = 'office-slide-generation';
        layer.setAttribute('aria-hidden','true');
        layer.style.setProperty('--hue', ((page.index * 11) % 41 - 20) + 'deg');
        layer.style.setProperty('--phase', -(page.index % 17) * .17 + 's');
        const placeholder = document.createElement('div');
        placeholder.dataset.placeholder = '';
        for (const [x,y,w,h] of layouts[page.index % layouts.length]) {
          const bar = document.createElement('i');
          Object.assign(bar.style, { left:x+'%', top:y+'%', width:(w * (1 - .12 * ((page.index * .618034) % 1)))+'%', height:h+'%' });
          placeholder.append(bar);
        }
        const cursor = document.createElement('div');
        cursor.dataset.agent = '';
        cursor.append(document.createElement('span'));
        layer.append(placeholder, cursor);
        thumb.append(layer);
      }
      const cursor = layer.querySelector('[data-agent]');
      if (layer.dataset.phase !== page.phase) {
        layer.dataset.phase = page.phase;
        mounted();
        cursor.firstElementChild.textContent = (locale === 'en'
        ? { creating:'Agent · Creating', waiting:'Agent · Preparing', partial:'Agent · Preparing', writing:'Agent · Writing' }
        : { creating:'Agent · 建页中', waiting:'Agent · 准备中', partial:'Agent · 准备中', writing:'Agent · 生成中' })[page.phase];
      }
      const target = page.target;
      if (drawnTargets.get(layer) === target) continue;
      const old = layer.querySelector('[data-module]');
      if (old) { layer.append(cursor); old.remove(); }
      if (!target) { drawnTargets.delete(layer); continue; }
      const bounds = moduleBounds(thumb, target, slideSize);
      if (!bounds) continue;
      drawnTargets.set(layer, target);
      const [x,y,w,h] = bounds;
      const module = document.createElement('div');
      module.dataset.module = target.op;
      Object.assign(module.style,{left:x*100+'%',top:y*100+'%',width:w*100+'%',height:h*100+'%'});
      if (x+w > .62) module.dataset.flipX = '';
      if (y+h > .84) module.dataset.flipY = '';
      const remaining = target.started + target.duration - view.performance.now();
      if (remaining > 0) {
        module.dataset.pulling = '';
        module.style.setProperty('--pull-ms', target.duration+'ms');
        // A virtualized thumbnail can remount halfway through a selection.
        module.style.setProperty('--elapsed', (remaining-target.duration)+'ms');
      }
      const marquee = document.createElement('div');
      marquee.dataset.marquee = '';
      cursor.firstElementChild.textContent = target.pulling ? agentLabel(locale,'selecting') : target.label;
      module.append(marquee,cursor);
      layer.append(module);
    }
  }
  const observer = new view.MutationObserver(() => { if (frame === undefined) frame = view.requestAnimationFrame(() => { frame = undefined; draw(); }); });
  observer.observe(document.body,{childList:true,subtree:true});
  draw();
  return {
    pages, firstThumbnail, track,
    phase(ids, phase) {
      for (const id of ids) {
        const page = pages.get(id);
        if (page) { page.phase = phase; page.target = undefined; }
      }
      draw();
    },
    async selectModule(step, geometry, label, moving, pause) {
      const page = pages.get(step.slideId);
      if (!page || disposed) return;
      track(step.slideId);
      // Styling a just-written object keeps the agent on that same object.
      const key = step.shapeRef ?? step.ref ?? step.shapeId;
      const previous = page.target;
      const duration = moving && !(key && previous?.key === key) ? selectionMs() : 0;
      const target = { key, op:step.op, shapeId:step.shapeId, geometry, whole:step.op==='setBackground',
        label, duration, started:view.performance.now(), pulling:duration > 0 };
      page.phase = 'writing';
      page.target = target;
      draw(step.slideId);
      if (!duration) return;
      await pause(duration);
      if (disposed || page.target !== target) return;
      target.pulling = false;
      for (const layer of document.querySelectorAll('.office-slide-generation[data-phase=writing]')) {
        if (drawnTargets.get(layer) !== target) continue;
        const module = layer.querySelector('[data-module]');
        module.removeAttribute('data-pulling');
        module.querySelector('span').textContent = label;
      }
    },
    dispose() { disposed = true; for (const event of ['wheel','touchmove']) document.removeEventListener(event,pauseFollowing); resume?.remove(); observer.disconnect(); view.cancelAnimationFrame(frame); style.remove(); document.querySelectorAll('.office-slide-generation').forEach(el => el.remove()); },
  };
}

const roleKeys = { cover:'title,subtitle', content:'title,body', title:'title', closing:'body', blank:'' };
export function plannedLayout(layouts, role) {
  return layouts?.find(layout => layout.placeholders?.join(',') === roleKeys[role])?.id
    ?? layouts?.find(layout => layout.placeholders?.length === 0)?.id;
}
