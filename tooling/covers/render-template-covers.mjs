#!/usr/bin/env node
/**
 * Covers for the inspiration library's built-in Word and Excel templates, rendered by the real editor.
 *
 *   node tooling/covers/render-template-covers.mjs --resources <package resources> [--python <python with python-docx
 *        and openpyxl>] [--out packages/dsh-adapter/src/client/covers] [--keep <dir>] [--headed]
 *
 * 1. template_samples.py fills each template (skills/shimo-docx/templates) with sample
 *    content in its own styles.
 * 2. Each sample opens through kit-host and office-kit's Kit in Chromium the way the DSH tab opens it (read-only view),
 *    and the first page is screenshotted: Word the whole first A4 page (595:842), Excel the top-left of the sheet at 4:3
 *    (column and row headers included, the selection moved off the shot).
 * 3. Each shot is scaled with ImageMagick (`magick`) to the card size, 2× for sharp screens (Word 420×594, Excel
 *    480×360), reduced to a 128-colour palette (they ship inside client.js; test/template-covers.test.mjs caps the
 *    total), and written to --out as <kind>-<id>.png; then src/client/template-covers.mjs is regenerated from the
 *    PNGs (data URIs the client bundle carries; test/template-covers.test.mjs checks the two agree).
 *
 * --resources: a built package's resources (kit, kit-page, editor, office-kit, host-fonts): packages/dsh-adapter/resources
 * after a build, or the released package's (`git archive origin/dist resources | tar -x -C <dir>`, then <dir>/resources).
 * Rerun whenever a template, the sample text or the editor's rendering changes.
 */
import { createServer } from 'node:http';
import { mkdirSync, mkdtempSync, readdirSync, realpathSync, rmSync, statSync } from 'node:fs';
import { execFileSync } from 'node:child_process';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { parseArgs } from 'node:util';
import { chromium } from 'playwright';
import { COVERS_DIR, writeCoversModule } from './covers-module.mjs';

const repo = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '../..');
const adapter = path.join(repo, 'packages/dsh-adapter');
const { values } = parseArgs({ options: {
  resources: { type: 'string' }, python: { type: 'string', default: 'python3' }, out: { type: 'string', default: path.join(adapter, 'src/client/covers') },
  keep: { type: 'string' }, headed: { type: 'boolean', default: false },
} });
if (!values.resources) { console.error('--resources <built package resources> is required'); process.exit(2); }
try { execFileSync('magick', ['-version'], { stdio: 'ignore' }); } catch { console.error('ImageMagick 7 (`magick`) is required to scale the shots: brew install imagemagick'); process.exit(2); }
const resources = path.resolve(values.resources);
const out = path.resolve(values.out);
const { KitDocuments, KitProcess, createKitHttpHandler } = await import(path.join(repo, 'packages/kit-host/src/index.mjs'));

/** Output size per kind (2× the largest card width the library draws). */
const SIZE = { word: [420, 594], excel: [480, 360] };
/** Excel: CSS pixels of the sheet shot, from the grid's top-left corner (4:3). */
const SHEET_SHOT = { width: 560, height: 420 };

const PARENT = `<!doctype html><meta charset="utf-8"><style>html,body{margin:0;height:100%}iframe{border:0;width:100%;height:100%}</style>
<iframe id="doc"></iframe><script type="module">
const frame = document.getElementById('doc');
let view;
addEventListener('message', event => {
  const m = event.data;
  if (event.source !== frame.contentWindow || !m || m.v !== 1 || m.id !== view || typeof m.type !== 'string') return;
  const name = m.type.replace('officesdk:', '');
  if (name === 'ready') window.ready = true;
  if (name === 'failed') window.failed = m;
});
window.openView = (id, entry) => { view = id; window.ready = false; window.failed = undefined; frame.src = entry + '&revision=1&parentOrigin=' + encodeURIComponent(location.origin); };
</script>`;

const ID = /^(word|excel)-([a-z0-9-]+)\.(docx|xlsx)$/;
const scale = (input, output, [width, height]) => execFileSync('magick', [input, '-resize', `${width}x${height}!`, '-strip', '-colors', '128', '-define', 'png:compression-level=9', `PNG8:${output}`]);

// Everything from here owns something to clean up (temporary folder, Kit, server, browser): one finally releases it.
const base = realpathSync(mkdtempSync(path.join(tmpdir(), 'template-covers-')));
let docs, server, browser;
try {
  const workspace = path.join(base, 'ws');
  const shots = values.keep ? path.resolve(values.keep) : path.join(base, 'shots');
  mkdirSync(workspace);
  mkdirSync(shots, { recursive: true });
  mkdirSync(out, { recursive: true });
  console.log(execFileSync(values.python, [path.join(repo, 'tooling/covers/template_samples.py'), '--out', workspace], { encoding: 'utf8' }).trim());

  const kit = new KitProcess({ binary: path.join(resources, 'kit', 'offline-kit-runtime'), work: path.join(base, 'kit') });
  kit.on('log', text => { if (/error|panic/i.test(text)) console.log('[kit]', text.trim().slice(0, 300)); });
  docs = new KitDocuments({ kit, exportRoot: path.join(base, 'exports'),
    authorize: async () => {}, stat: async (_scope, file) => { const absolutePath = path.resolve(workspace, file); const info = statSync(absolutePath); return { absolutePath, version: `${info.mtimeMs}:${info.size}:${info.ino}` }; },
    writeAccess: async () => 'outside' });
  const scope = { sessionId: 's', workspaceRoot: workspace };
  const peer = { id: 'tab', principal: 'owner', signal: new AbortController().signal };
  let origin;
  const handler = createKitHttpHandler({ documents: docs, admit: () => ({ peer }), allowedOrigin: value => value === origin,
    staticRoots: { page: path.join(resources, 'kit-page'), editor: path.join(resources, 'editor'), resources: path.join(resources, 'office-kit'), fonts: path.join(resources, 'host-fonts') },
    meta: view => ({ v: 1, id: view.id, suite: view.suite, mode: 'view', name: path.basename(view.path), diskChanged: false, user: { uid: 'local:covers', name: 'covers' } }) });
  server = createServer(async (req, res) => {
    if (req.url === '/') { res.writeHead(200, { 'content-type': 'text/html' }); res.end(PARENT); return; }
    if (!(await handler(req, res))) { res.statusCode = 404; res.end(); }
  });
  await new Promise(resolve => server.listen(0, '127.0.0.1', resolve));
  origin = `http://127.0.0.1:${server.address().port}`;

  browser = await chromium.launch({ headless: !values.headed });
  // Tall enough for the whole first Word page at the editor's fit-width zoom.
  const page = await browser.newPage({ viewport: { width: 1400, height: 1900 }, deviceScaleFactor: 2, colorScheme: 'light' });
  await page.goto(`${origin}/`);
  const frame = () => page.frames().find(f => f.url().includes('/kit/static/page/'));

  const open = async file => {
    const view = await docs.open(scope, file, peer);
    await page.evaluate(({ id, entry }) => window.openView(id, entry), { id: view.id, entry: `/shimo/kit/static/page/index.html?view=${view.id}` });
    // @ts-ignore set by PARENT
    await page.waitForFunction(() => window.ready || window.failed, null, { timeout: 120_000 });
    // @ts-ignore set by PARENT
    const failed = await page.evaluate(() => window.failed);
    if (failed) throw new Error(`${file} failed to open: ${JSON.stringify(failed).slice(0, 300)}`);
    // Fonts and the first layout pass settle after "ready".
    await page.waitForTimeout(5000);
    return view;
  };
  const close = async view => {
    await page.evaluate(() => { document.getElementById('doc').src = 'about:blank'; });
    await page.waitForTimeout(300);
    await docs.close(view.id, peer).catch(() => {});
  };

  /** The first page's rectangle (CSS pixels; the frame fills the page): the topmost A4-shaped SVG of the print preview. */
  const wordPage = async () => {
    const rect = await frame().evaluate(() => {
      const preview = document.querySelector('.writer-document-preview');
      const pages = [...(preview?.querySelectorAll('svg') ?? [])].map(el => el.getBoundingClientRect())
        .filter(r => r.width > 200 && Math.abs(r.width / r.height - 595 / 842) < 0.01);
      const first = pages.sort((a, b) => a.y - b.y)[0];
      return first ? { x: first.x, y: first.y, width: first.width, height: first.height } : null;
    });
    if (!rect) throw new Error('no A4 page found in the Writer preview (.writer-document-preview svg)');
    if (rect.y < 0 || rect.y + rect.height > page.viewportSize().height) throw new Error(`the first page is not fully visible: ${JSON.stringify(rect)}`);
    return rect;
  };

  /** The sheet grid's top-left (the select-all corner), after moving the active cell (A1 on open) off the shot. */
  const sheetArea = async () => {
    const grid = await frame().evaluate(() => {
      const canvases = [...document.querySelectorAll('canvas')].map(el => el.getBoundingClientRect()).filter(r => r.width > 600 && r.height > 400);
      const largest = canvases.sort((a, b) => b.width * b.height - a.width * a.height)[0];
      return largest ? { x: largest.x, y: largest.y, width: largest.width, height: largest.height } : null;
    });
    if (!grid) throw new Error('no sheet canvas found');
    // A cell well outside the shot but inside the visible grid.
    await page.mouse.click(grid.x + SHEET_SHOT.width + 300, grid.y + SHEET_SHOT.height + 300);
    await page.waitForTimeout(800);
    return { x: grid.x, y: grid.y, ...SHEET_SHOT };
  };

  for (const file of readdirSync(workspace).filter(name => ID.test(name)).sort()) {
    const [, kind, id] = file.match(ID);
    const view = await open(file);
    const clip = kind === 'word' ? await wordPage() : await sheetArea();
    const raw = path.join(shots, `${kind}-${id}.raw.png`);
    await page.screenshot({ path: raw, clip });
    scale(raw, path.join(out, `${kind}-${id}.png`), SIZE[kind]);
    console.log(`${kind}-${id}.png`, JSON.stringify(clip));
    await close(view);
  }
  // The client module is regenerated only from the shipped covers directory.
  if (out === COVERS_DIR) console.log(writeCoversModule(out));
  else console.log(`covers written to ${out}; src/client/template-covers.mjs left unchanged`);
} finally {
  await browser?.close().catch(() => {});
  server?.close();
  await docs?.dispose().catch(() => {});
  rmSync(base, { recursive: true, force: true });
}
