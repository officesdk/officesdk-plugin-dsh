import test from 'node:test';
import assert from 'node:assert/strict';
import { createOverlay, keyAllowedWhileBusy, pointerTarget, locateSlide, screenRect, readView, locateCard, SORTER, NORMAL } from '../src/browser/presentation-agent.mjs';

/**
 * Just enough DOM for the overlay: elements with style and attributes, one main slide surface whose
 * position, size and slide the test changes, a thumbnail, a clipping editor area, a ribbon button,
 * listeners, and animation frames and time run by hand.
 */
function fakePage({ left = 100, top = 50, width = 480, slide = 'slide-1', area = { left: 0, top: 0, right: 2000, bottom: 2000 }, clipped = true, named = true } = {}) {
  const frames = new Map(); let nextFrame = 1;
  const listeners = [];
  const element = tag => {
    const attributes = new Map(), children = [];
    const node = { tag, style: {}, dataset: {}, textContent: '', parentElement: undefined, children,
      setAttribute: (name, value) => attributes.set(name, String(value)), getAttribute: name => attributes.get(name) ?? null,
      append: (...items) => { for (const item of items) { item.parentElement = node; children.push(item); } },
      remove: () => { const list = node.parentElement?.children; if (list) list.splice(list.indexOf(node), 1); node.parentElement = undefined; },
      closest: selector => { for (let at = node; at; at = at.parentElement) if (selector === 'svg.ppt-slide-surface' ? at.surface : selector.split(', ').includes(at.tag)) return at; return null; },
      contains: other => { for (let at = other; at; at = at.parentElement) if (at === node) return true; return false; },
      appendChild: item => { node.append(item); return item; },
      // Laid out where its style puts it; a status pill is as wide as its text (7 px a character).
      getBoundingClientRect: () => { const left = px(node.style.left) || 0, top = px(node.style.top) || 0, width = tag === 'span' ? node.textContent.length * 7 : px(node.style.width) || 0, height = px(node.style.height) || 0; return { left, top, width, height, right: left + width, bottom: top + height }; },
      get ownerDocument() { return node.owner ?? document; }, set ownerDocument(value) { node.owner = value; } };
    return node;
  };
  const surfaceOf = rect => { const node = element('svg'); node.surface = true; node.viewBox = { baseVal: { width: 9144000, height: 5143500 } }; node.getBoundingClientRect = rect; return node; };
  const body = element('body');
  const shell = element('div');
  const editor = element('div');
  editor.overflow = clipped ? 'auto' : 'visible';
  editor.getBoundingClientRect = () => ({ ...page.area, width: page.area.right - page.area.left, height: page.area.bottom - page.area.top });
  // 10 px scroll bars on the right and at the bottom.
  Object.defineProperties(editor, { clientWidth: { get: () => page.area.right - page.area.left - 10 }, clientHeight: { get: () => page.area.bottom - page.area.top - 10 } });
  const surface = surfaceOf(() => ({ left: page.left, top: page.top, width: page.width, height: page.width * 9 / 16 }));
  surface.setAttribute('data-mop-svg-slide', slide);
  if (named) surface.setAttribute('data-mop-svg-definition-namespace', `mop-main-${slide}`);
  const shape = element('g');
  surface.append(shape);
  editor.append(surface);
  const list = element('div');
  list.overflow = 'auto';
  list.getBoundingClientRect = () => ({ left: 0, top: 560, right: 190, bottom: 960, width: 190, height: 400 });
  const thumbnail = surfaceOf(() => ({ left: 0, top: page.thumbnailTop, width: 150, height: 84 }));
  thumbnail.setAttribute('data-mop-svg-slide', 'slide-9');
  thumbnail.setAttribute('data-mop-svg-definition-namespace', 'mop-thumbnail-slide-9');
  const thumbnailShape = element('g');
  thumbnail.append(thumbnailShape);
  // As in the editor: each thumbnail sits in a wrapper of its own size that hides overflow, inside the scrolling list.
  const card = element('div');
  card.overflow = 'hidden';
  card.getBoundingClientRect = thumbnail.getBoundingClientRect;
  card.append(thumbnail);
  list.append(card);
  const ribbon = element('button');
  shell.append(editor, list, ribbon);
  body.append(shell);
  const listen = scope => ({
    addEventListener: (type, listener, capture) => listeners.push({ scope, type, listener, capture }),
    removeEventListener: (type, listener) => { const index = listeners.findIndex(item => item.scope === scope && item.type === type && item.listener === listener); if (index >= 0) listeners.splice(index, 1); },
  });
  const view = {
    ...listen('window'),
    performance: { now: () => page.time },
    requestAnimationFrame: callback => { const id = nextFrame++; frames.set(id, callback); return id; },
    cancelAnimationFrame: id => frames.delete(id),
    getComputedStyle: node => ({ overflow: node.overflow ?? 'visible', overflowX: 'visible', overflowY: 'visible' }),
  };
  const document = { body, defaultView: view, createElement: element, createElementNS: (namespace, tag) => element(tag), ...listen('document'),
    querySelectorAll: selector => selector === 'svg.ppt-slide-surface' ? [thumbnail, surface] : [],
    // What lies under a point: the list over thumbnails (the editor's cover there is the list itself) unless a menu floats there.
    elementFromPoint: (x, y) => page.floating ?? (x < 190 && y >= 560 ? list : body) };
  surface.ownerDocument = thumbnail.ownerDocument = document;
  const page = { left, top, width, area, thumbnailTop: 600, floating: undefined, list, time: 0, document, surface, thumbnail, frames, listeners, editor, shape, thumbnailShape, ribbon,
    /** Runs the frames requested so far, as the browser would before the next paint. */
    tick(ms = 16) { page.time += ms; const due = [...frames]; frames.clear(); for (const [, callback] of due) callback(); },
    show(id) { surface.setAttribute('data-mop-svg-slide', id); },
    /** Dispatches one event to the capture listeners of `scope`; true when the overlay stopped it. */
    fire(scope, type, target, extra = {}) {
      let stopped = false;
      const event = { type, target, clientX: 0, clientY: 0, ...extra, preventDefault() { stopped = true; }, stopPropagation() { stopped = true; }, stopImmediatePropagation() { stopped = true; } };
      for (const item of listeners.filter(entry => entry.scope === scope && entry.type === type && entry.capture)) item.listener(event);
      return stopped;
    },
    get box() { return body.children.find(node => node.dataset.shimoAgent === 'frame'); },
    get cursor() { return body.children.find(node => node.dataset.shimoAgent === 'cursor'); },
    /** The selection drawn in the box: the AttentionOverlay's svg. */
    get outline() { return page.box.children.find(node => node.getAttribute('data-attention-overlay') === 'true'); },
    get caption() { return body.children.find(node => node.dataset.shimoAgent === 'caption'); } };
  return page;
}
const px = value => Number.parseFloat(value);
const placed = node => node.style.display === 'none' ? null : [px(node.style.left), px(node.style.top), px(node.style.width), px(node.style.height)];
const title = { slideId: 'slide-1', geometry: { left: 30, top: 60, width: 300, height: 60 } };

test('screenRect measures the main slide, not a thumbnail, and only for the slide it shows', () => {
  const page = fakePage();
  // 720 pt wide slide drawn 480 px wide: 2/3 px per point.
  assert.deepEqual(screenRect(page.document, { left: 30, top: 60, width: 300, height: 60 }, 'slide-1'), { left: 120, top: 90, width: 200, height: 40, clip: { left: 0, top: 0, right: 2000, bottom: 2000, width: 2000, height: 2000 } });
  assert.equal(screenRect(page.document, { left: 30, top: 60, width: 300, height: 60 }, 'slide-2'), undefined, 'another slide is shown');
  assert.ok(screenRect(page.document, { left: 0, top: 0, width: 1, height: 1 }), 'without a slide id any shown slide counts');
  assert.equal(locateSlide(page.document).surface, page.surface);
  // Zoomed far out the main slide is smaller than a thumbnail; it is still the one measured.
  page.width = 120;
  assert.equal(locateSlide(page.document).surface, page.surface);
  assert.equal(screenRect(page.document, { left: 0, top: 0, width: 720, height: 405 }, 'slide-1').width, 120);
});

test('the outline sits on its shape and follows the slide as it is scrolled, zoomed and resized', () => {
  const page = fakePage();
  const overlay = createOverlay(page.document);
  overlay.frame(title, { text: 'Agent · 修改中' });
  assert.deepEqual(placed(page.box), [111, 81, 218, 58], 'the shape plus the 9 px gap of a text outline');
  // The user scrolls the canvas up and left, then zooms in twice as far.
  page.left = 40; page.top = 10; page.width = 960;
  page.tick();
  assert.deepEqual(placed(page.box), [71, 81, 418, 98], 'moved and scaled with the slide');
  // A window or panel resize shrinks the slide again.
  page.left = 200; page.top = 120; page.width = 240;
  page.tick();
  assert.deepEqual(placed(page.box), [201, 131, 118, 38]);
  overlay.dispose();
  assert.equal(page.frames.size, 0, 'no frame keeps running after dispose');
  assert.equal(page.listeners.length, 0, 'and no listener stays behind');
});

test('a new target on the same slide glides there and keeps following the view meanwhile', () => {
  const page = fakePage();
  const overlay = createOverlay(page.document);
  overlay.frame(title);
  overlay.frame({ slideId: 'slide-1', geometry: { left: 30, top: 240, width: 300, height: 60 } });
  assert.deepEqual(placed(page.box), [111, 81, 218, 58], 'starts where the last target was');
  page.tick(125);
  const halfway = placed(page.box)[1];
  assert.ok(halfway > 81 && halfway < 201, `between the two targets (${halfway})`);
  page.top = 0; // the user scrolls during the glide
  page.tick(125);
  assert.deepEqual(placed(page.box), [111, 151, 218, 58], 'ends on the new target where the slide is now');
  // Another slide's target does not glide across slides.
  overlay.frame({ slideId: 'slide-2', geometry: title.geometry });
  assert.equal(placed(page.box), null, 'slide 2 is not shown yet');
  overlay.dispose();
});

test('the outline hides while another slide is shown and comes back with its slide; it stops when the edit ends', () => {
  const page = fakePage();
  const overlay = createOverlay(page.document);
  overlay.frame(title);
  overlay.write(title, '云帆');
  assert.ok(placed(page.box) && placed(page.caption));
  assert.equal(page.caption.textContent, '云帆');
  page.show('slide-3');
  page.tick();
  assert.equal(placed(page.box), null, 'hidden on another slide');
  assert.equal(placed(page.caption), null);
  page.show('slide-1');
  page.tick();
  assert.deepEqual(placed(page.box), [111, 81, 218, 58], 'back on its slide');
  assert.deepEqual(placed(page.caption), [120, 90, 200, 40]);
  // The caption alone keeps the frames running; with nothing left they stop.
  overlay.frame(undefined);
  page.tick();
  assert.equal(page.frames.size, 1);
  overlay.write(undefined);
  page.tick();
  assert.equal(page.frames.size, 0);
  overlay.frame(title);
  overlay.busy(false);
  assert.equal(placed(page.box), null);
  assert.equal(page.frames.size, 0, 'no frame keeps running after the edit');
  overlay.dispose();
});

test('the outline is cut off where the editor area cuts off the slide', () => {
  const page = fakePage({ left: 300, top: 50, width: 480, area: { left: 0, top: 0, right: 500, bottom: 2000 } });
  const overlay = createOverlay(page.document);
  overlay.frame({ slideId: 'slide-1', geometry: { left: 270, top: 30, width: 300, height: 60 } });
  // Box 471..689 wide, 61..119 high; the area ends at 500, so 189 px on the right are cut off. The other
  // sides reach past the box, which keeps the glow around it.
  assert.equal(page.box.style.clipPath, 'inset(-61px 189px -1881px -471px)');
  overlay.dispose();
  const open = fakePage({ clipped: false });
  const plain = createOverlay(open.document);
  plain.frame(title);
  assert.equal(open.box.style.clipPath, '', 'nothing cuts the slide off: no clip');
  plain.dispose();
});

test('during an AI edit the slide and the ribbon take no input; scrolling, zooming and picking a slide still work', () => {
  const page = fakePage({ area: { left: 0, top: 0, right: 2000, bottom: 2000 } });
  const overlay = createOverlay(page.document);
  const onSlide = { clientX: 150, clientY: 100 }, onThumbnail = { clientX: 10, clientY: 610 }, besideSlide = { clientX: 1000, clientY: 1500 }, onScrollbar = { clientX: 1995, clientY: 500 }, onRibbon = { clientX: 2500, clientY: 10 };
  assert.equal(page.fire('window', 'pointerdown', page.shape, onSlide), false, 'between edits the user edits directly');
  overlay.busy(true, 'AI 编辑中');
  assert.ok(!page.listeners.some(item => item.type === 'wheel'), 'wheel and pinch zoom are never stopped');
  assert.equal(page.fire('window', 'pointerdown', page.shape, onSlide), true, 'a press on the slide is stopped');
  assert.equal(page.fire('window', 'click', page.shape, onSlide), true);
  assert.equal(page.fire('window', 'pointerdown', page.editor, besideSlide), true, 'beside the slide (a selected shape\'s handles) is stopped');
  assert.equal(page.fire('window', 'pointerdown', page.ribbon, onRibbon), true, 'the ribbon is stopped');
  assert.equal(page.fire('window', 'contextmenu', page.ribbon, onThumbnail), true, 'menus are stopped');
  assert.equal(page.fire('document', 'keydown', page.shape, { key: 'a' }), true, 'typing is stopped');
  // A click on a thumbnail selects that slide, whatever element the editor put over it.
  assert.equal(page.fire('window', 'pointerdown', page.ribbon, onThumbnail), false);
  assert.equal(page.fire('window', 'pointermove', page.ribbon, { clientX: 12, clientY: 611, buttons: 1 }), false);
  assert.equal(page.fire('window', 'pointerup', page.ribbon, { clientX: 12, clientY: 611 }), false);
  assert.equal(page.fire('window', 'click', page.ribbon, { clientX: 12, clientY: 611 }), false);
  // Dragging a thumbnail would move the slide: the moves are stopped, the release still reaches the editor.
  page.fire('window', 'pointerdown', page.thumbnailShape, onThumbnail);
  assert.equal(page.fire('window', 'pointermove', page.thumbnailShape, { clientX: 10, clientY: 640, buttons: 1 }), true);
  assert.equal(page.fire('window', 'pointerup', page.thumbnailShape, { clientX: 10, clientY: 640 }), false, 'the editor is not left pressed');
  page.fire('window', 'click', page.thumbnailShape, { clientX: 10, clientY: 640 });
  assert.equal(page.fire('window', 'dragstart', page.thumbnailShape, onThumbnail), true);
  // A drag begun before the edit moves nothing while it runs; plain hovering passes.
  assert.equal(page.fire('window', 'pointermove', page.shape, { clientX: 160, clientY: 110, buttons: 1 }), true);
  assert.equal(page.fire('window', 'pointerup', page.shape, { clientX: 160, clientY: 110 }), true);
  assert.equal(page.fire('window', 'pointermove', page.shape, { clientX: 160, clientY: 110, buttons: 0 }), false);
  // The main view's scroll bars pass.
  assert.equal(page.fire('window', 'pointerdown', page.editor, onScrollbar), false);
  assert.equal(pointerTarget(page.document, 150, 100), 'other');
  assert.equal(pointerTarget(page.document, 10, 610), 'thumbnail');
  assert.equal(pointerTarget(page.document, 1995, 500), 'scrollbar');
  overlay.busy(false);
  assert.equal(page.fire('window', 'pointerdown', page.shape, onSlide), false, 'after the edit input reaches the editor again');
  overlay.dispose();
});

test('a thumbnail counts only where its list shows it, with nothing over it, and only when the main view is known by name', () => {
  const page = fakePage();
  page.thumbnailTop = 520; // scrolled half out of the list, which starts at 560: its top half lies over the ribbon
  assert.equal(pointerTarget(page.document, 10, 540), 'other');
  assert.equal(pointerTarget(page.document, 10, 580), 'thumbnail');
  page.floating = page.ribbon; // a ribbon menu open over the list
  assert.equal(pointerTarget(page.document, 10, 580), 'other');
  page.floating = undefined;
  const overlay = createOverlay(page.document);
  overlay.busy(true);
  assert.equal(page.fire('window', 'pointerdown', page.ribbon, { clientX: 10, clientY: 580, button: 2 }), true, 'a right press (a menu) is stopped');
  assert.equal(page.fire('window', 'mousedown', page.ribbon, { clientX: 10, clientY: 580, button: 0, ctrlKey: true }), true, 'so is Ctrl-click');
  overlay.dispose();
  const unnamed = fakePage({ named: false });
  assert.equal(pointerTarget(unnamed.document, 10, 610), 'other', 'without the name a thumbnail never passes');
  // A control the list holds over a thumbnail (say a hover button) is not the slide.
  const button = page.document.createElement('button');
  page.list.append(button);
  assert.equal(pointerTarget(page.document, 10, 580), 'thumbnail');
  page.floating = button;
  assert.equal(pointerTarget(page.document, 10, 580), 'other', 'a button in the list over a thumbnail is stopped');
  page.floating = undefined;
  // Only surfaces named as thumbnails count: another surface over the same spot (a layout gallery) does not.
  page.thumbnail.setAttribute('data-mop-svg-definition-namespace', 'mop-layout-1');
  assert.equal(pointerTarget(page.document, 10, 580), 'other', 'a surface not named as a thumbnail never passes');
});

test('during an AI edit only navigation in the slide list, copy and Escape reach the editor', () => {
  const page = fakePage();
  const key = (key, extra = {}) => keyAllowedWhileBusy({ type: 'keydown', key, target: page.thumbnailShape, ...extra }, page.document);
  assert.equal(key('ArrowDown'), true, 'in the slide list an arrow picks a slide');
  assert.equal(key('ArrowDown', { target: page.list }), true, 'so it does with the focus on the list itself');
  page.thumbnail.setAttribute('data-mop-svg-definition-namespace', 'mop-layout-1');
  assert.equal(key('ArrowDown', { target: page.list }), false, 'a container of surfaces not named as thumbnails is not the slide list');
  page.thumbnail.setAttribute('data-mop-svg-definition-namespace', 'mop-thumbnail-slide-9');
  assert.equal(key('ArrowDown', { target: page.shape }), false, 'on the slide it would nudge a shape');
  assert.equal(key('ArrowDown', { target: page.document.body }), false, 'so it does wherever else the focus is');
  assert.equal(key('ArrowUp', { ctrlKey: true, shiftKey: true }), false, 'moving a slide is stopped');
  assert.equal(key('c', { metaKey: true }), true);
  assert.equal(key('Escape'), true);
  for (const blocked of ['=', '-', '0', 'g', 'l', ']', 'd', 'v', 'z']) assert.equal(key(blocked, { metaKey: true }), false, blocked);
  assert.equal(key('Delete'), false);
  assert.equal(key('a'), false);
  assert.equal(keyAllowedWhileBusy({ type: 'paste', target: page.ribbon }, page.document), false);
  assert.equal(keyAllowedWhileBusy({ type: 'keydown', key: 'ArrowDown', target: page.thumbnailShape }, fakePage({ named: false }).document), false, 'without the main view by name nothing counts as the list');
});

test('a new target during a glide starts from where the box is drawn', () => {
  const page = fakePage();
  const overlay = createOverlay(page.document);
  overlay.frame(title);
  overlay.frame({ slideId: 'slide-1', geometry: { left: 30, top: 240, width: 300, height: 60 } });
  page.tick(125);
  const middle = placed(page.box)[1];
  overlay.frame({ slideId: 'slide-1', geometry: { left: 30, top: 0, width: 300, height: 60 } });
  assert.equal(placed(page.box)[1], middle, 'no jump');
  page.tick(250);
  assert.deepEqual(placed(page.box), [111, 41, 218, 58]);
  overlay.dispose();
});

test('several slides at once: each track outlines its own slide and only the slide in the main view is drawn', () => {
  const page = fakePage();
  const overlay = createOverlay(page.document);
  overlay.frame(title, { text: 'Agent · 新增中', key: 'slide-1' });
  overlay.frame({ slideId: 'slide-2', geometry: { left: 30, top: 120, width: 300, height: 60 } }, { text: 'Agent · 新增中', key: 'slide-2' });
  overlay.write({ slideId: 'slide-2', geometry: { left: 30, top: 120, width: 300, height: 60 } }, '第二页▍', 'slide-2');
  const boxes = () => page.document.body.children.filter(node => node.dataset.shimoAgent === 'frame');
  const captions = () => page.document.body.children.filter(node => node.dataset.shimoAgent === 'caption');
  assert.equal(boxes().length, 3, 'the main track plus one per slide');
  assert.deepEqual(boxes().map(placed), [null, [111, 81, 218, 58], null], 'slide 1 is shown: only its box');
  assert.equal(captions().filter(node => placed(node)).length, 0, 'the typing on slide 2 is not drawn over slide 1');
  // The user picks slide 2 in the list: its box and typing show, slide 1's go.
  page.show('slide-2');
  page.tick();
  assert.deepEqual(boxes().map(placed), [null, null, [111, 121, 218, 58]]);
  assert.equal(captions()[2].textContent, '第二页▍');
  // Clearing one track leaves the other; the end of the edit removes the extra tracks.
  overlay.frame(undefined, { key: 'slide-2' });
  page.tick();
  assert.equal(placed(boxes()[2]), null);
  // Without the main view's slide id (K9) the tracks of several slides cannot be told apart: none is drawn.
  overlay.frame(title, { key: 'slide-1' });
  overlay.frame({ slideId: 'slide-2', geometry: title.geometry }, { key: 'slide-2' });
  const read = page.surface.getAttribute;
  page.surface.getAttribute = name => name === 'data-mop-svg-slide' ? null : read(name);
  page.tick();
  assert.deepEqual(boxes().map(placed), [null, null, null]);
  overlay.frame(undefined, { key: 'slide-2' });
  overlay.write(undefined, undefined, 'slide-2');
  page.tick();
  assert.deepEqual(boxes().map(placed), [null, null, null], 'not even the last one left: it may not be the slide shown');
  page.surface.getAttribute = read;
  page.show('slide-1');
  page.tick();
  assert.deepEqual(boxes().map(placed), [null, [111, 81, 218, 58], null], 'with the slide id back it is drawn again');
  overlay.busy(true, 'AI 正在生成第 1–2 页');
  overlay.busy(false);
  assert.equal(boxes().length, 1);
  assert.equal(page.frames.size, 0);
  overlay.dispose();
});

test('the thumbnails of the slides being written are ringed where the list shows them, and follow the list', () => {
  const page = fakePage();
  const overlay = createOverlay(page.document);
  const ring = () => page.document.body.children.find(node => node.dataset.shimoAgent === 'mark');
  overlay.mark(['slide-9', 'slide-404']);
  assert.deepEqual(placed(ring()), [-4, 596, 158, 92], 'around the thumbnail with a 4 px margin');
  assert.equal(ring().style.clipPath, 'inset(-36px -36px -272px 4px)', 'cut off where the list cuts off the thumbnail, not by the thumbnail\'s own wrapper');
  assert.equal(page.document.body.children.filter(node => node.dataset.shimoAgent === 'mark').length, 2);
  assert.equal(placed(page.document.body.children.filter(node => node.dataset.shimoAgent === 'mark')[1]), null, 'no ring for a slide the list does not show');
  assert.ok(page.frames.size > 0, 'rings follow on every frame');
  // The list scrolls.
  page.thumbnailTop = 700;
  page.tick();
  assert.deepEqual(placed(ring()), [-4, 696, 158, 92]);
  // The main slide itself is never ringed, even when it shows a marked slide.
  overlay.mark(['slide-1']);
  assert.equal(placed(ring()), null);
  // Nor is a surface not named as a thumbnail (a layout gallery) that shows the same slide.
  page.thumbnail.setAttribute('data-mop-svg-definition-namespace', 'mop-layout-1');
  overlay.mark(['slide-9']);
  assert.equal(placed(ring()), null, 'only thumbnails are ringed');
  page.thumbnail.setAttribute('data-mop-svg-definition-namespace', 'mop-thumbnail-slide-9');
  overlay.mark([]);
  page.tick();
  assert.equal(page.frames.size, 0, 'nothing left to follow');
  overlay.mark(['slide-9']);
  overlay.busy(false);
  assert.equal(placed(ring()), null, 'the end of the edit clears the rings');
  overlay.dispose();
  assert.equal(page.document.body.children.filter(node => node.dataset.shimoAgent === 'mark').length, 0);
});

const rects = node => [...(node.tag === 'rect' ? [node] : []), ...node.children.flatMap(rects)];
const at = cursor => cursor.style.display === 'none' ? null : [px(cursor.style.left), px(cursor.style.top)];
const nodesOf = (page, kind) => page.document.body.children.filter(node => node.dataset.shimoAgent === kind);
const watchers = page => page.listeners.filter(item => item.type === 'visibilitychange').length;

test('a text target is outlined 9 px around it with rounded corners, an edge target on its edge with square ones; the selection is drawn at 70%', () => {
  const page = fakePage();
  const overlay = createOverlay(page.document);
  overlay.frame(title, { text: 'Agent · 修改中' });
  assert.deepEqual(placed(page.box), [111, 81, 218, 58]);
  assert.equal(page.outline.style.opacity, '0.7', 'stroke and light together');
  assert.equal(page.outline.style.overflow, 'visible', 'the glow reaches past the box');
  assert.ok(rects(page.outline).length > 0 && rects(page.outline).every(rect => rect.getAttribute('rx') === '10'));
  assert.deepEqual([rects(page.outline)[0].getAttribute('x'), rects(page.outline)[0].getAttribute('width')], ['0', '218']);
  assert.equal(page.cursor.style.opacity, undefined, 'the cursor is not faded');
  overlay.frame(title, { text: 'Agent · 修改中', kind: 'edge' });
  assert.deepEqual(placed(page.box), [120, 90, 200, 40], 'on the shape itself');
  assert.equal(page.outline.style.opacity, '0.7');
  assert.equal(page.outline.style.overflow, 'hidden');
  assert.ok(rects(page.outline).every(rect => rect.getAttribute('rx') === '0'), 'square corners');
  // The 1.6 px line is drawn half a line inside, so its outer edge is the target's.
  assert.deepEqual([rects(page.outline)[0].getAttribute('x'), rects(page.outline)[0].getAttribute('width')], ['0.8', '198.4']);
  overlay.dispose();
});

test('a whole slide is outlined on its edge as the main view draws it, only while it is shown', () => {
  const page = fakePage();
  const overlay = createOverlay(page.document);
  overlay.frame({ slideId: 'slide-1', whole: true }, { text: 'Agent · 修改中' });
  assert.deepEqual(placed(page.box), [100, 50, 480, 270]);
  assert.equal(page.outline.style.overflow, 'hidden', 'drawn as an edge');
  page.left = 60;
  page.tick();
  assert.deepEqual(placed(page.box), [60, 50, 480, 270], 'follows the slide');
  page.show('slide-3');
  page.tick();
  assert.equal(placed(page.box), null, 'hidden while another slide is shown');
  assert.equal(at(page.cursor), null);
  // Without the main view's slide id (K9) the main track still draws it, as it does a shape; slide tracks do not.
  page.show('slide-1');
  const read = page.surface.getAttribute;
  page.surface.getAttribute = name => name === 'data-mop-svg-slide' ? null : read(name);
  page.tick();
  assert.deepEqual(placed(page.box), [60, 50, 480, 270]);
  overlay.frame({ slideId: 'slide-1', whole: true }, { key: 'slide-1' });
  overlay.frame({ slideId: 'slide-2', whole: true }, { key: 'slide-2' });
  page.tick();
  assert.deepEqual(nodesOf(page, 'frame').slice(1).map(placed), [null, null]);
  overlay.dispose();
});

test('switching between a shape, a whole slide and a chart jumps there, rebuilds the outline and leaves nothing behind', () => {
  const page = fakePage();
  const overlay = createOverlay(page.document);
  const chart = { slideId: 'slide-1', geometry: { left: 30, top: 240, width: 300, height: 60 } };
  overlay.frame(title, { text: 'Agent · 修改中' });
  const first = page.outline;
  overlay.frame({ slideId: 'slide-1', whole: true }, { text: 'Agent · 修改中' });
  assert.deepEqual(placed(page.box), [100, 50, 480, 270], 'no glide into a whole slide');
  assert.notEqual(page.outline, first);
  assert.equal(first.parentElement, undefined, 'the old outline is gone');
  overlay.frame(title, { text: 'Agent · 修改中' });
  assert.deepEqual(placed(page.box), [111, 81, 218, 58], 'nor out of one');
  overlay.frame(chart, { text: 'Agent · 修改中', kind: 'edge' });
  assert.deepEqual(placed(page.box), [120, 210, 200, 40], 'another look on the same slide: no glide either');
  overlay.frame(title, { text: 'Agent · 修改中' });
  page.tick(125);
  assert.ok(placed(page.box).every(Number.isFinite));
  assert.deepEqual(placed(page.box), [111, 81, 218, 58]);
  assert.equal(page.box.children.length, 1, 'one outline in the box');
  assert.equal(watchers(page), 1, 'and one listener for it');
  overlay.dispose();
  assert.equal(watchers(page), 0);
});

test('the agent cursor points at the bottom-right corner of the drawn box, or the top-right one without room below, and moves with it', () => {
  const page = fakePage();
  const overlay = createOverlay(page.document);
  overlay.frame(title, { text: 'Agent · 修改中' });
  const [arrow, pill] = page.cursor.children;
  assert.equal(pill.textContent, 'Agent · 修改中');
  assert.deepEqual(at(page.cursor), [329, 139], 'the bottom-right corner of the box, gap included');
  assert.equal(page.cursor.style.visibility, 'visible');
  assert.equal(arrow.style.transform ?? '', '');
  assert.equal(pill.style.top, '10px');
  // No room under the box for the status (38 px): it moves to the top-right corner, the arrow turned up.
  page.area = { left: 0, top: 0, right: 2000, bottom: 170 };
  page.tick();
  assert.deepEqual(at(page.cursor), [329, 81]);
  assert.equal(arrow.style.transform, 'scaleY(-1)');
  assert.deepEqual([pill.style.top, pill.style.transform], ['-8px', 'translateY(-100%)']);
  assert.equal(page.cursor.style.visibility, 'visible');
  page.area = { left: 0, top: 0, right: 2000, bottom: 2000 };
  page.tick();
  assert.deepEqual([at(page.cursor), arrow.style.transform], [[329, 139], '']);
  // During a glide it stays on the box as drawn now, not on the target.
  overlay.frame({ slideId: 'slide-1', geometry: { left: 30, top: 240, width: 300, height: 60 } }, { text: 'Agent · 新增中' });
  page.tick(125);
  const [left, top, width, height] = placed(page.box);
  assert.ok(top > 81 && top < 201);
  assert.deepEqual(at(page.cursor), [left + width, top + height]);
  assert.equal(pill.textContent, 'Agent · 新增中');
  overlay.frame(title);
  assert.equal(at(page.cursor), null, 'no status, no cursor');
  overlay.dispose();
});

test('the cursor hides while its corner or its status would leave the editor area, and comes back', () => {
  const page = fakePage({ area: { left: 0, top: 0, right: 500, bottom: 2000 } });
  const overlay = createOverlay(page.document);
  // Corner at 329; the status (11 characters, 77 px) starts 10 px to its right and ends at 416.
  overlay.frame(title, { text: 'Agent · 修改中' });
  assert.equal(page.cursor.style.visibility, 'visible');
  page.left = 190; // corner at 419: the status would end at 506
  page.tick();
  assert.equal(page.cursor.style.visibility, 'hidden', 'not cut in half');
  page.left = 300; // the corner itself is outside
  page.tick();
  assert.equal(page.cursor.style.visibility, 'hidden');
  page.left = 100;
  page.tick();
  assert.equal(page.cursor.style.visibility, 'visible');
  // No room below, nor above.
  page.area = { left: 0, top: 100, right: 500, bottom: 160 };
  page.tick();
  assert.equal(page.cursor.style.visibility, 'hidden');
  overlay.dispose();
});

test('several agents at once: each track has its own box and cursor, and only those of the slide shown are drawn', () => {
  const page = fakePage();
  const overlay = createOverlay(page.document);
  overlay.frame(title, { text: 'Agent · 新增中', key: 'slide-1' });
  overlay.frame({ slideId: 'slide-1', geometry: { left: 30, top: 240, width: 300, height: 60 } }, { text: 'Agent · 修改中', key: 'slide-1b' });
  overlay.frame({ slideId: 'slide-2', geometry: title.geometry }, { text: 'Agent · 新增中', key: 'slide-2' });
  const cursors = () => nodesOf(page, 'cursor');
  assert.equal(cursors().length, 4, 'the main track plus one per agent');
  assert.deepEqual(cursors().map(at), [null, [329, 139], [329, 259], null]);
  assert.deepEqual(cursors().map(node => node.children[1].textContent), ['', 'Agent · 新增中', 'Agent · 修改中', 'Agent · 新增中']);
  page.show('slide-2');
  page.tick();
  assert.deepEqual(cursors().map(at), [null, null, null, [329, 139]]);
  overlay.busy(false);
  assert.equal(cursors().length, 1, 'the end of the edit removes the agents\' tracks');
  assert.equal(at(cursors()[0]), null);
  overlay.dispose();
  assert.equal(cursors().length, 0);
});

test('the end of an edit and dispose leave no outline, cursor, frame or listener behind', () => {
  const page = fakePage();
  const overlay = createOverlay(page.document);
  overlay.busy(true, 'AI 编辑中');
  overlay.frame(title, { text: 'Agent · 修改中' });
  overlay.frame({ slideId: 'slide-1', whole: true }, { text: 'Agent · 删除中', key: 'slide-1' });
  overlay.write(title, '云帆');
  page.tick();
  assert.equal(watchers(page), 2);
  overlay.busy(false);
  assert.deepEqual([placed(page.box), at(page.cursor), placed(page.caption)], [null, null, null]);
  assert.equal(page.frames.size, 0, 'no frame keeps running, the outline\'s light included');
  assert.equal(watchers(page), 1, 'the slide track\'s outline went with it');
  overlay.dispose();
  assert.deepEqual(['frame', 'cursor', 'caption', 'mark'].flatMap(kind => nodesOf(page, kind)), []);
  assert.equal(page.frames.size, 0);
  assert.equal(page.listeners.length, 0);
});

test('while its box is hidden the outline stops its light; it only redraws when the box changes size', () => {
  const page = fakePage();
  const overlay = createOverlay(page.document);
  overlay.frame(title, { text: 'Agent · 修改中' });
  page.tick();
  assert.equal(page.frames.size, 2, 'the overlay follows the slide and the outline moves its light');
  const width = () => rects(page.outline)[0].getAttribute('width');
  const viewBox = page.outline.getAttribute('viewBox');
  page.left = 140; // scrolled: the box moves, its size does not
  page.tick();
  assert.equal(page.outline.getAttribute('viewBox'), viewBox);
  page.width = 960; // zoomed: a new size is drawn
  page.tick();
  assert.equal(width(), '418');
  page.show('slide-3');
  page.tick();
  page.tick();
  assert.equal(page.frames.size, 1, 'only the overlay waits for its slide; the hidden outline runs nothing');
  page.show('slide-1');
  page.tick();
  assert.deepEqual([placed(page.box)?.length, page.outline.style.display, page.frames.size], [4, 'block', 2], 'back with its slide, drawn and lit again');
  overlay.dispose();
});

test('a new status is measured again: a longer one that would leave the area hides the cursor', () => {
  const page = fakePage({ area: { left: 0, top: 0, right: 500, bottom: 2000 } });
  const overlay = createOverlay(page.document);
  overlay.frame(title, { text: 'Agent · 修改中' }); // ends at 329 + 10 + 77 = 416
  assert.equal(page.cursor.style.visibility, 'visible');
  overlay.frame(title, { text: 'Agent · Deleting this long text' }); // 30 characters: ends at 549
  assert.equal(page.cursor.style.visibility, 'hidden');
  overlay.frame(title, { text: 'Agent · 删除中' });
  assert.equal(page.cursor.style.visibility, 'visible');
  overlay.dispose();
});

// Slide sorter (spec 261105-sorter-view-live B2–B5): cards instead of the main view.
test('the view is read from the page: normal while the main slide is drawn, slide sorter while only cards are', () => {
  const page = fakePage();
  assert.equal(readView(page.document), NORMAL);
  page.width = 0;
  assert.equal(readView(page.document), SORTER, 'the main slide is hidden, a card is shown');
  page.thumbnailTop = 600;
  assert.equal(readView({}), undefined, 'no page to read');
  const card = locateCard(page.document, 'slide-9');
  assert.deepEqual([card.box.left, card.box.top, card.box.width, card.box.height], [0, 600, 150, 84]);
  assert.equal(card.area, page.list, 'cut off by the scrolling list, not by the card\'s own wrapper');
  assert.equal(locateCard(page.document, 'slide-1'), undefined, 'the main slide is not a card');
});

test('in the slide sorter a whole-slide frame goes on the card, nothing is projected on the main view, and no rings', () => {
  const page = fakePage();
  const overlay = createOverlay(page.document, { animate: () => false });
  overlay.busy(true);
  overlay.mark(['slide-9']);
  overlay.view(SORTER);
  overlay.frame({ slideId: 'slide-9', whole: true }, { text: 'Agent · 创建页面', kind: 'edge' });
  overlay.frame(title, { key: 'other', text: 'Agent · 写入标题' });
  page.tick();
  assert.deepEqual(placed(page.box), [0, 600, 150, 84], 'the frame is on the card');
  const shapeBoxes = page.document.body.children.filter(node => node.dataset.shimoAgent === 'frame' && node !== page.box);
  assert.ok(shapeBoxes.every(node => node.style.display === 'none'), 'a shape frame is not projected while there is no main view');
  assert.ok(page.document.body.children.filter(node => node.dataset.shimoAgent === 'mark').every(node => node.style.display === 'none'), 'no ring: the frame on the card is the mark');
  overlay.view(NORMAL);
  page.tick();
  assert.equal(placed(page.box), null, 'back in the normal view, slide-9 is not the slide shown: no frame on its thumbnail');
  overlay.view(SORTER);
  overlay.busy(false);
  overlay.frame({ slideId: 'slide-9', whole: true }, { text: 'x', kind: 'edge' });
  page.tick();
  assert.equal(placed(page.box), null, 'the edit ended: the overlay is back in the normal view');
  overlay.dispose();
});
