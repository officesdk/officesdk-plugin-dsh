#!/usr/bin/env node
/**
 * Measure the slides of an HTML style (Genspark layout, e.g. Noir) in headless Chrome (style-distill tooling;
 * developer machines only; Node 22+, Google Chrome).
 *
 *   node tooling/style-distill/measure_html.mjs --root <folder to copy> --layouts <layouts.json> --slides <dir under root>
 *        [--font-css <css under root>] [--pages 01,02] --out measure.json
 *
 * The source is untrusted: <root> is copied into a new temporary folder and the copy is opened there, never the
 * original. Chrome runs headless with its own empty profile, page JavaScript disabled (Emulation.setScriptExecution
 * Disabled) and every request outside the copy refused (Fetch: only file:// inside the copy and data: pass). Nothing
 * is read back from the page with JavaScript: boxes and computed styles come from DOMSnapshot.captureSnapshot.
 *
 * Placeholders are filled the way the source's generator does: {{TEXT:id}} with the slot's `example` from
 * layouts.json, {{ASSET:id}} with a grey 1×1 image (photos are not needed to measure), {{ALT:…}}, {{VISUAL:…}} and
 * {{PAGE_TITLE}} with nothing, @@FONT_CSS@@ with a link to --font-css. Output per page (px on the 1920×1080 canvas):
 * regions (data-region: box, kind, background, opacity, border, filter, blend, transform, z), text slots (data-slot or
 * data-text-slots: element box, union of its line boxes, line count, font size/family/weight/style, colour, letter
 * spacing, line height, alignment, transform, opacity, text) and visual slots (data-visual-slot: box, object-fit,
 * filter).
 */
import { spawn } from 'node:child_process';
import { cpSync, existsSync, lstatSync, mkdtempSync, readdirSync, readFileSync, realpathSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join, relative, resolve, sep } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';

const CHROME = process.env.CHROME_PATH ?? '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome';
const STYLES = ['display', 'position', 'z-index', 'font-size', 'font-family', 'font-weight', 'font-style', 'color', 'letter-spacing', 'line-height',
  'text-align', 'text-transform', 'opacity', 'background-color', 'background-image', 'border-top-width', 'border-top-color', 'border-left-width',
  'border-left-color', 'border-bottom-width', 'border-bottom-color', 'border-right-width', 'border-right-color', 'border-radius', 'filter',
  'mix-blend-mode', 'transform', 'object-fit', 'writing-mode', 'white-space'];
const GREY = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mN4/w8AAn8B/ZJ8b9gAAAAASUVORK5CYII=';
const escapeHtml = text => String(text ?? '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

function args(argv) {
  const out = { pages: null };
  for (let i = 0; i < argv.length; i += 1) {
    const key = argv[i].replace(/^--/, '');
    if (!['root', 'layouts', 'slides', 'font-css', 'pages', 'out'].includes(key)) throw new Error(`未知参数 ${argv[i]}`);
    out[key] = argv[++i];
  }
  for (const key of ['root', 'layouts', 'slides', 'out']) if (!out[key]) throw new Error(`缺少 --${key}`);
  return out;
}

/** Refuse a source with symbolic links: a link copied into the sandbox would let Chrome read files outside it. */
function assertNoLinks(dir) {
  for (const entry of readdirSync(dir, { withFileTypes: true })) {
    const path = join(dir, entry.name);
    if (lstatSync(path).isSymbolicLink()) throw new Error(`源包里有符号链接 ${path}，不复制也不打开（先确认它指向哪里）`);
    if (entry.isDirectory()) assertNoLinks(path);
  }
}

class Cdp {
  constructor(url) { this.ws = new WebSocket(url); this.id = 0; this.waiting = new Map(); this.handlers = []; }
  open() {
    return new Promise((ok, fail) => {
      this.ws.onopen = ok; this.ws.onerror = fail;
      this.ws.onmessage = event => {
        const msg = JSON.parse(event.data);
        if (msg.id && this.waiting.has(msg.id)) {
          const { resolve: done, reject } = this.waiting.get(msg.id); this.waiting.delete(msg.id);
          if (msg.error) reject(new Error(`${msg.error.message} (${msg.error.code})`)); else done(msg.result);
        } else for (const handler of this.handlers) handler(msg);
      };
    });
  }
  send(method, params = {}, sessionId) {
    const id = ++this.id;
    this.ws.send(JSON.stringify({ id, method, params, ...(sessionId ? { sessionId } : {}) }));
    return new Promise((done, reject) => this.waiting.set(id, { resolve: done, reject }));
  }
  once(method, sessionId, timeout = 20000) {
    return new Promise((done, reject) => {
      const timer = setTimeout(() => reject(new Error(`等 ${method} 超时`)), timeout);
      const handler = msg => { if (msg.method === method && msg.sessionId === sessionId) { clearTimeout(timer); this.handlers.splice(this.handlers.indexOf(handler), 1); done(msg.params); } };
      this.handlers.push(handler);
    });
  }
}

async function launch(profile) {
  const child = spawn(CHROME, ['--headless=new', '--remote-debugging-port=0', `--user-data-dir=${profile}`, '--no-first-run', '--no-default-browser-check',
    '--disable-extensions', '--disable-background-networking', '--disable-sync', '--disable-component-update', '--mute-audio', '--hide-scrollbars', 'about:blank'],
  { stdio: ['ignore', 'ignore', 'pipe'] });
  const portFile = join(profile, 'DevToolsActivePort');
  for (let k = 0; k < 200 && !existsSync(portFile); k += 1) await new Promise(r => setTimeout(r, 100));
  if (!existsSync(portFile)) { child.kill(); throw new Error('Chrome 没有启动（找不到 DevToolsActivePort）'); }
  const [port, path] = readFileSync(portFile, 'utf8').trim().split('\n');
  return { child, url: `ws://127.0.0.1:${port}${path}` };
}

/** Elements, their attributes, boxes, computed styles and text from one DOMSnapshot document. */
function readSnapshot(snap) {
  const doc = snap.documents[0], s = snap.strings, str = i => (i < 0 ? '' : s[i]);
  const n = doc.nodes, lay = doc.layout;
  const layoutOf = new Map(lay.nodeIndex.map((node, k) => [node, k]));
  const children = new Map();
  n.parentIndex.forEach((p, i) => { if (!children.has(p)) children.set(p, []); children.get(p).push(i); });
  const attrs = i => { const a = n.attributes[i] ?? [], o = {}; for (let k = 0; k < a.length; k += 2) o[str(a[k])] = str(a[k + 1]); return o; };
  const style = k => Object.fromEntries(STYLES.map((name, j) => [name, str(lay.styles[k][j])]));
  // Line boxes of text nodes (textBoxes → layoutIndex), grouped by the text node.
  const lines = new Map();
  doc.textBoxes.layoutIndex.forEach((li, t) => { const node = lay.nodeIndex[li]; if (!lines.has(node)) lines.set(node, []); lines.get(node).push(doc.textBoxes.bounds[t]); });
  const textOf = i => (n.nodeType[i] === 3 ? str(n.nodeValue[i]) : (children.get(i) ?? []).map(textOf).join(n.nodeName[i] && str(n.nodeName[i]) === 'BR' ? '\n' : ''));
  const linesUnder = i => [...(lines.get(i) ?? []), ...(children.get(i) ?? []).flatMap(linesUnder)];
  return { count: n.parentIndex.length, attrs, box: i => (layoutOf.has(i) ? lay.bounds[layoutOf.get(i)] : null), style: i => (layoutOf.has(i) ? style(layoutOf.get(i)) : null), textOf, linesUnder, name: i => str(n.nodeName[i]) };
}

function union(boxes) {
  if (!boxes.length) return null;
  const l = Math.min(...boxes.map(b => b[0])), t = Math.min(...boxes.map(b => b[1]));
  return [l, t, Math.max(...boxes.map(b => b[0] + b[2])) - l, Math.max(...boxes.map(b => b[1] + b[3])) - t];
}
const round = box => box && box.map(v => Math.round(v * 10) / 10);
/** Visual lines: text boxes whose tops differ by more than half a line. */
function lineCount(boxes) {
  const tops = [...boxes].sort((a, b) => a[1] - b[1]);
  let count = 0, last = -Infinity;
  for (const b of tops) if (b[1] - last > b[3] * 0.5) { count += 1; last = b[1]; }
  return count;
}

async function measurePage(cdp, sessionId, file) {
  const loaded = cdp.once('Page.loadEventFired', sessionId);
  await cdp.send('Page.navigate', { url: pathToFileURL(file).href }, sessionId);
  await loaded;
  await new Promise(r => setTimeout(r, 400));   // web fonts settle (no page JavaScript to wait on document.fonts)
  const snap = await cdp.send('DOMSnapshot.captureSnapshot', { computedStyles: STYLES, includeDOMRects: true }, sessionId);
  const d = readSnapshot(snap);
  const regions = {}, texts = {}, visuals = {};
  let background = null;
  for (let i = 0; i < d.count; i += 1) {
    const a = d.attrs(i);
    if (!background && /(^|\s)slide-container(\s|$)/.test(a.class ?? '')) background = d.style(i)?.['background-color'] ?? null;
    if (a['data-region']) {
      const st = d.style(i) ?? {};
      regions[a['data-region']] = { kind: a['data-object-type'] ?? '', box: round(d.box(i)), style: st };
    }
    const slotIds = a['data-slot'] ? [a['data-slot']] : a['data-text-slots'] ? a['data-text-slots'].split(/\s+/) : [];
    if (slotIds.length) {
      const boxes = d.linesUnder(i);
      const entry = { box: round(d.box(i)), ink: round(union(boxes)), lines: lineCount(boxes), tag: d.name(i), style: d.style(i), text: d.textOf(i).replace(/[ \t\n]+/g, ' ').trim(), slots: slotIds };
      for (const id of slotIds) if (!texts[id]) texts[id] = entry;
    }
    if (a['data-visual-slot']) visuals[a['data-visual-slot']] = { box: round(d.box(i)), style: d.style(i) };
  }
  return { background, regions, texts, visuals };
}

async function main() {
  const a = args(process.argv.slice(2));
  const root = resolve(a.root);
  const layouts = JSON.parse(readFileSync(resolve(a.layouts), 'utf8'));
  const pages = Object.entries(layouts.layouts).filter(([pid]) => !a.pages || a.pages.split(',').includes(pid));
  const work = mkdtempSync(join(tmpdir(), 'measure-html-'));
  const copy = join(work, 'src'), profile = join(work, 'profile');
  let chrome;
  try {
    assertNoLinks(root);
    cpSync(root, copy, { recursive: true, dereference: false, filter: path => !/\.(py|mjs|js|sh)$/.test(path) || path.endsWith('.html') });
    const fill = (html, page) => html
      .replace(/\{\{TEXT:([^}]+)\}\}/g, (_, id) => escapeHtml(page.texts?.[id]?.example ?? ''))
      .replace(/\{\{ASSET:[^}]+\}\}/g, GREY)
      .replace(/\{\{(?:ALT|VISUAL):[^}]+\}\}|\{\{PAGE_TITLE\}\}/g, '')
      .replace('@@FONT_CSS@@', a['font-css'] ? `</style><link rel="stylesheet" href="${relative(join(copy, a.slides), join(copy, a['font-css'])).split(sep).join('/')}"><style>` : '');
    const files = [];
    for (const [pid, page] of pages) {
      const src = join(copy, a.slides, page.file);
      const out = src.replace(/\.html$/, '.measure.html');
      writeFileSync(out, fill(readFileSync(src, 'utf8'), page));
      files.push([pid, out]);
    }
    chrome = await launch(profile);
    const cdp = new Cdp(chrome.url);
    await cdp.open();
    const { targetId } = await cdp.send('Target.createTarget', { url: 'about:blank' });
    const { sessionId } = await cdp.send('Target.attachToTarget', { targetId, flatten: true });
    const inside = pathToFileURL(realpathSync(copy)).href;
    cdp.handlers.push(msg => {
      if (msg.method !== 'Fetch.requestPaused' || msg.sessionId !== sessionId) return;
      const url = msg.params.request.url;
      let ok = url.startsWith('data:');
      if (!ok && url.startsWith('file://')) {
        // Compare real paths, so nothing that resolves outside the copy is served.
        try { ok = pathToFileURL(realpathSync(fileURLToPath(url.split(/[?#]/)[0]))).href.startsWith(inside + '/'); } catch { ok = false; }
      }
      cdp.send(ok ? 'Fetch.continueRequest' : 'Fetch.failRequest', ok ? { requestId: msg.params.requestId } : { requestId: msg.params.requestId, errorReason: 'BlockedByClient' }, sessionId).catch(() => {});
    });
    await cdp.send('Fetch.enable', { patterns: [{ urlPattern: '*' }] }, sessionId);
    await cdp.send('Page.enable', {}, sessionId);
    await cdp.send('Emulation.setScriptExecutionDisabled', { value: true }, sessionId);
    await cdp.send('Emulation.setDeviceMetricsOverride', { width: layouts.width ?? 1920, height: layouts.height ?? 1080, deviceScaleFactor: 1, mobile: false }, sessionId);
    const result = { canvas: { width: layouts.width ?? 1920, height: layouts.height ?? 1080 }, pages: {} };
    for (const [pid, file] of files) result.pages[pid] = { name: layouts.layouts[pid].name, ...(await measurePage(cdp, sessionId, file)) };
    writeFileSync(resolve(a.out), JSON.stringify(result, null, 1));
    console.log(JSON.stringify({ pages: files.length, out: resolve(a.out) }));
    await cdp.send('Browser.close').catch(() => {});
  } finally {
    if (chrome && chrome.child.exitCode === null) {
      const exited = new Promise(done => chrome.child.once('exit', done));
      chrome.child.kill();
      await Promise.race([exited, new Promise(done => setTimeout(done, 5000))]);
    }
    rmSync(work, { recursive: true, force: true, maxRetries: 5, retryDelay: 200 });
  }
}

main().catch(error => { console.error(error.message); process.exitCode = 1; });
