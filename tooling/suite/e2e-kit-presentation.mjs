#!/usr/bin/env node
/**
 * PPTX through office-kit's Kit in a real browser (spec §12.7 #5): the built package's Kit, editor bundle,
 * SDK resources and page, served by kit-host's HTTP handler; the parent page plays the DSH tab.
 *
 *   node tooling/suite/e2e-kit-presentation.mjs [--resources packages/dsh-adapter/resources] [--deck test/fixtures/sample.pptx] [--shots <dir>] [--headed]
 */
import { createServer } from 'node:http';
import { copyFileSync, mkdirSync, mkdtempSync, realpathSync, rmSync, writeFileSync } from 'node:fs';
import { execFileSync } from 'node:child_process';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { isDeepStrictEqual, parseArgs } from 'node:util';
import { chromium } from 'playwright';

const repo = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '../..');
const { values } = parseArgs({ options: { large: { type: 'string' }, resources: { type: 'string', default: path.join(repo, 'packages/dsh-adapter/resources') }, deck: { type: 'string', default: path.join(repo, 'test/fixtures/sample.pptx') }, shots: { type: 'string' }, headed: { type: 'boolean', default: false } } });
const { KitDocuments, KitProcess, createKitHttpHandler } = await import(path.join(repo, 'packages/kit-host/src/index.mjs'));
const resources = path.resolve(values.resources);

const base = realpathSync(mkdtempSync(path.join(tmpdir(), 'kit-e2e-')));
const workspace = path.join(base, 'ws');
mkdirSync(workspace);
copyFileSync(path.resolve(values.deck), path.join(workspace, 'deck.pptx'));
const kit = new KitProcess({ binary: path.join(resources, 'kit', 'offline-kit-runtime'), work: path.join(base, 'kit') });
kit.on('log', text => { if (/error|panic/i.test(text)) console.log('[kit]', text.trim().slice(0, 300)); });
const docs = new KitDocuments({ kit, exportRoot: path.join(base, 'exports'),
  authorize: async () => {}, stat: async (scope, file) => ({ absolutePath: path.resolve(workspace, file) }),
  writeAccess: async (scope, real) => real.startsWith(workspace + path.sep) ? 'allowed' : 'outside' });
const scope = { sessionId: 's', workspaceRoot: workspace };
const peer = { id: 'tab', principal: 'owner', signal: new AbortController().signal };

const PARENT = `<!doctype html><meta charset="utf-8"><style>html,body{margin:0;height:100%}iframe{border:0;width:100%;height:100%}</style>
<iframe id="doc"></iframe><script type="module">
const frame = document.getElementById('doc');
const pending = new Map(); let view, revision = 1, n = 0;
window.events = [];
addEventListener('message', event => {
  const m = event.data;
  if (event.source !== frame.contentWindow || event.origin !== location.origin || !m || m.v !== 1 || m.id !== view || typeof m.type !== 'string') return;
  const name = m.type.replace('officesdk:', '');
  window.events.push({ name, ...m });
  if (name === 'ready') window.ready = true;
  if (name === 'failed') window.failed = m.message;
  if (name === 'agent-result') { pending.get(m.requestId)?.(m); pending.delete(m.requestId); }
});
window.openView = (id, entry) => { view = id; window.ready = false; window.failed = undefined; frame.src = entry + '&revision=1&parentOrigin=' + encodeURIComponent(location.origin); };
window.agent = (op, args) => new Promise(resolve => { const requestId = 'r' + (++n); pending.set(requestId, resolve); frame.contentWindow.postMessage({ type: 'officesdk:command', v: 1, id: view, revision, command: 'agent', requestId, op, args: args ?? {} }, location.origin); });
</script>`;
let origin;
const handler = createKitHttpHandler({ documents: docs, admit: () => ({ peer }), allowedOrigin: value => value === origin,
  staticRoots: { page: path.join(resources, 'kit-page'), editor: path.join(resources, 'editor'), resources: path.join(resources, 'office-kit') },
  meta: view => ({ v: 1, id: view.id, suite: view.suite, mode: view.writer ? 'edit' : 'view', name: path.basename(view.path), diskChanged: view.diskChanged }) });
const server = createServer(async (req, res) => {
  if (req.url === '/') { res.writeHead(200, { 'content-type': 'text/html' }); res.end(PARENT); return; }
  if (!(await handler(req, res))) { res.statusCode = 404; res.end(); }
});
await new Promise(resolve => server.listen(0, '127.0.0.1', resolve));
origin = `http://127.0.0.1:${server.address().port}`;

const results = [];
const within = (promise, ms, what) => Promise.race([promise, new Promise((_, reject) => setTimeout(() => reject(new Error(`${what} timed out after ${ms} ms`)), ms))]);
const check = (name, ok, detail) => { results.push({ name, ok }); console.log(`${ok ? 'ok  ' : 'FAIL'} ${name}${detail !== undefined ? ` ${JSON.stringify(detail).slice(0, 300)}` : ''}`); };
const savedText = () => { const file = path.join(base, 'saved.pptx'); copyFileSync(path.join(workspace, 'deck.pptx'), file); return execFileSync('sh', ['-c', `unzip -p "${file}" 'ppt/slides/*.xml'`], { encoding: 'utf8', maxBuffer: 64 * 1024 * 1024 }); };
const browser = await chromium.launch({ headless: !values.headed });
const errors = [];
try {
  const page = await browser.newPage({ viewport: { width: 1280, height: 800 } });
  page.on('pageerror', error => errors.push(error.message));
  page.on('console', message => { if (message.type() === 'error' && !/Failed to load resource/.test(message.text())) errors.push(message.text().slice(0, 200)); });
  // A view the test already released or replaced may still finish loading images: only the live view counts.
  let liveView;
  page.on('response', response => { const where = new URL(response.url()).pathname; if (response.status() >= 400 && response.status() !== 416 && (!where.includes('/kit/v/') || where.includes(`/kit/v/${liveView}/`))) errors.push(`${response.status()} ${where}`); });
  await page.goto(`${origin}/`);
  const open = async () => {
    const view = await docs.open(scope, 'deck.pptx', peer);
    liveView = view.id;
    await page.evaluate(({ id, entry }) => window.openView(id, entry), { id: view.id, entry: `/shimo/kit/static/page/index.html?view=${view.id}` });
    await page.waitForFunction(() => window.ready || window.failed, null, { timeout: 120_000 });
    return { view, failed: await page.evaluate(() => window.failed) };
  };
  // Like the DSH tab: the frame goes away before its view is released.
  const closeView = async (view, options) => {
    // Requests already queued by the old iframe may finish after it is detached. They are stale once this view
    // closes, so the response probe must stop treating its expected 409s as page errors.
    if (liveView === view.id) liveView = undefined;
    await page.evaluate(() => { document.getElementById('doc').src = 'about:blank'; });
    await page.waitForTimeout(300);
    await docs.close(view.id, peer, options).catch(() => {});
  };
  const started = Date.now();
  const first = await open();
  check('the editor page mounts the presentation from Kit', !first.failed, { failed: first.failed, ms: Date.now() - started });
  if (values.shots) { mkdirSync(values.shots, { recursive: true }); await page.screenshot({ path: path.join(values.shots, 'mounted.png') }); }
  const frame = page.frameLocator('iframe');
  const surfaces = await frame.locator('svg.ppt-slide-surface').evaluateAll(nodes => nodes.map(node => { const r = node.getBoundingClientRect(); return { x: r.x, y: r.y, width: r.width, height: r.height }; }));
  check('a slide is drawn', surfaces.length > 0, surfaces.length);
  if (!surfaces.length) throw new Error('nothing to edit: the slide was not drawn');

  // A hand edit: type into the first text the page shows.
  const typeOnce = async (find, text) => {
    // The title as drawn on the slide being edited (the thumbnails reuse the same text).
    let target;
    for (let i = 0; i < 50 && !target; i++) {
      const boxes = await frame.locator('svg.ppt-slide-surface text').evaluateAll((nodes, wanted) => nodes.filter(node => node.textContent.includes(wanted)).map(node => { const r = node.getBoundingClientRect(); return { x: r.x + r.width / 2, y: r.y + r.height / 2, area: r.width * r.height }; }), find);
      target = boxes.sort((a, b) => b.area - a.area)[0];
      if (!target) await page.waitForTimeout(200);
    }
    if (!target) throw new Error(`no text "${find}" on the slide`);
    await page.mouse.dblclick(target.x, target.y);
    await page.keyboard.press('Meta+A');
    await page.keyboard.insertText(text);
    await page.keyboard.press('Escape'); await page.keyboard.press('Escape');
  };  // A double click right after a save can land before the editor takes it; retry until the editor reports a change.
  const typeAt = async (find, text) => {
    const before = (await within(page.evaluate(() => window.agent('flush')), 20_000, 'flush')).result?.editRevision ?? 0;
    for (let attempt = 0; attempt < 3; attempt++) {
      await typeOnce(find, text);
      const after = (await within(page.evaluate(() => window.agent('flush')), 20_000, 'flush')).result?.editRevision ?? 0;
      if (after > before) return;
      await page.waitForTimeout(500);
    }
    throw new Error(`typing "${text}" never reached the editor`);
  };

  await typeAt('Roadmap', 'KIT-HAND');
  await page.waitForFunction(() => window.events.some(event => event.name === 'dirty'), null, { timeout: 15_000 }).catch(() => {});
  const flushed = await within(page.evaluate(() => window.agent('flush')), 20_000, 'flush').catch(error => ({ ok: false, message: error.message }));
  check('a hand edit reaches Kit as numbered changes and the page reports it', flushed.ok && flushed.result.editRevision > 0 && docs.views.get(first.view.id).unsaved, flushed.ok ? flushed.result : flushed);
  check('the file is untouched until a save', !savedText().includes('KIT-HAND'));
  const kitContent = await (await kit.ready()).request('content', { documentId: docs.views.get(first.view.id).documentId });
  const pageSnapshot = await within(page.evaluate(() => window.agent('snapshot')), 20_000, 'snapshot').catch(error => ({ ok: false, message: error.message }));
  check('Kit\'s composed content equals the page\'s own model', kitContent.ok && pageSnapshot.ok && isDeepStrictEqual(JSON.parse(kitContent.content), JSON.parse(pageSnapshot.result.content)), { kit: kitContent.ok, page: pageSnapshot.ok ? 'ok' : pageSnapshot.message });
  const other = { id: 'tab-2', principal: 'owner', signal: new AbortController().signal };
  const readOnly = await docs.open({ ...scope, sessionId: 's2' }, 'deck.pptx', other);
  const refused = await docs.change(readOnly.id, other, { seq: 0, body: '{}' }).then(() => 'accepted', error => error.code);
  check('a second conversation on the same file gets a read-only view', readOnly.mode === 'view' && refused === 'read-only', { mode: readOnly.mode, refused });
  await docs.close(readOnly.id, other);
  await docs.save(first.view.id, peer);
  check('save writes the edit into the PPTX', savedText().includes('KIT-HAND'));

  await closeView(first.view, { autosave: false });
  const second = await open();
  const content = await (await fetch(`${origin}/shimo/kit/v/${second.view.id}/content`)).text().catch(() => '');
  check('reopening shows the saved edit', !second.failed && (content.includes('KIT-HAND') || savedText().includes('KIT-HAND')));

  if (values.shots) await page.screenshot({ path: path.join(values.shots, 'reopened.png') });
  // AI editing on the Kit page (spec §12.5): the typing is drawn on the overlay, the document takes final values.
  const kitText = async viewId => (await (await kit.ready()).request('content', { documentId: docs.views.get(viewId).documentId })).content;
  const deck = (await within(page.evaluate(() => window.agent('inspect')), 30_000, 'inspect')).result;
  const slideId = deck.slides[0].id;
  const title = deck.slides[0].shapes.find(shape => shape.text?.includes('KIT-HAND'));
  check('the AI reads the deck through the page', Boolean(title), { slides: deck.slides.length, title: title?.text });
  const seqBefore = docs.views.get(second.view.id).nextSeq;
  const ops = [{ op: 'setText', slideId, shapeId: title.id, text: 'AI 改写的标题' }, { op: 'addTextBox', slideId, text: 'AI 新增的说明', left: 60, top: 470, width: 420, height: 36 }];
  const applying = page.evaluate(ops => window.agent('apply', { ops, editId: 'ai1' }), ops);
  let captionSeen = false;
  for (let i = 0; i < 40 && !captionSeen; i++) { captionSeen = await frame.locator('[data-shimo-agent="caption"]').isVisible().catch(() => false); if (!captionSeen) await page.waitForTimeout(50); }
  if (values.shots) await page.screenshot({ path: path.join(values.shots, 'ai-typing.png') });
  const applied = await within(applying, 60_000, 'apply');
  const afterAi = await kitText(second.view.id);
  const aiChanges = docs.views.get(second.view.id).nextSeq - seqBefore;
  check('an AI edit animates on the overlay and Kit receives only final values', applied.ok && captionSeen && afterAi.includes('AI 改写的标题') && afterAi.includes('AI 新增的说明') && !afterAi.includes('▍') && aiChanges <= 10 && docs.views.get(second.view.id).aiActive === 0,
    { ok: applied.ok, captionSeen, changes: aiChanges, editRevision: applied.result?.editRevision, error: applied.message });
  const undone = await within(page.evaluate(() => window.agent('undo', { editId: 'ai1' })), 30_000, 'undo');
  const afterUndo = await kitText(second.view.id);
  check('AI undo returns the deck to before the edit, in Kit too', undone.ok && !afterUndo.includes('AI 改写的标题') && !afterUndo.includes('AI 新增的说明') && afterUndo.includes('KIT-HAND'), undone.ok ? undone.result : undone.message);
  const long = Array.from({ length: 6 }, (_, index) => ({ op: 'setText', slideId, shapeId: title.id, text: `第 ${index} 次很长很长很长很长很长很长很长的改写` }));
  const cancelled = page.evaluate(ops => window.agent('apply', { ops, editId: 'ai2' }), long);
  await page.waitForTimeout(400);
  await page.evaluate(() => { const frame = document.getElementById('doc'); frame.contentWindow.postMessage({ type: 'officesdk:command', v: 1, id: frame.src.match(/view=([^&]+)/)[1], revision: 1, command: 'agent', requestId: 'cancel1', op: 'cancel', args: { target: 'r' + (window.events.filter(event => event.name === 'agent-result').length + 1) } }, location.origin); });
  const outcome = await within(cancelled, 60_000, 'cancel');
  await within(page.evaluate(() => window.agent('flush')), 20_000, 'flush');
  const afterCancel = await kitText(second.view.id);
  check('cancelling mid-animation leaves the deck as it was', outcome.ok === false && outcome.code === 'cancelled' && afterCancel.includes('KIT-HAND') && !afterCancel.includes('很长的改写'), { code: outcome.code });
  const again = await within(page.evaluate(ops => window.agent('apply', { ops, editId: 'ai3' }), [{ op: 'setText', slideId, shapeId: title.id, text: 'KIT-HAND AI' }]), 60_000, 'apply');
  await typeAt('AI', 'KIT-HAND');
  await within(page.evaluate(() => window.agent('flush')), 20_000, 'flush');
  const refusedUndo = await within(page.evaluate(() => window.agent('undo', { editId: 'ai3' })), 30_000, 'undo');
  check('after a hand edit the AI undo is refused and the hand edit stays', again.ok && refusedUndo.ok === false && refusedUndo.code === 'undo-unavailable', { code: refusedUndo.code });

  // Unsaved edits survive a Kit crash: the record is on disk.
  await typeAt('KIT', 'KIT-CRASH'); // the title now renders as separate words: KIT, -, HAND
  const crashFlush = await within(page.evaluate(() => window.agent('flush')), 20_000, 'flush').catch(error => ({ ok: false, message: error.message }));
  const generation = kit.generation;
  process.kill(kit.pid, 'SIGKILL');
  for (let i = 0; i < 200 && kit.generation === generation; i++) { await new Promise(resolve => setTimeout(resolve, 50)); await kit.ready().catch(() => {}); }
  await closeView(second.view, { autosave: false });
  const third = await open();
  const recovered = await docs.content(third.view.id, peer).catch(() => undefined);
  check('after a Kit crash the unsaved edit is still in the reopened document', crashFlush.ok && !third.failed && docs.views.get(third.view.id)?.unsaved === true && recovered?.content.includes('KIT-CRASH') && !savedText().includes('KIT-CRASH'), { generation: kit.generation, unsaved: docs.views.get(third.view.id)?.unsaved });
  if (values.shots) await page.screenshot({ path: path.join(values.shots, 'recovered.png') });

  // Another program rewrites the file while edits wait: save refuses until the user agrees to replace it.
  copyFileSync(path.resolve(values.deck), path.join(workspace, 'deck.pptx'));
  const changed = await docs.save(third.view.id, peer).then(() => 'saved', error => `${error.code}:${error.details?.overwrite}`);
  await docs.save(third.view.id, peer, { overwrite: true });
  check('a file changed on disk is only replaced with the user\'s overwrite', changed === 'source-changed:true' && savedText().includes('KIT-CRASH'), changed);

  // Unsaved edits, then the disk changes before the next open: keep editing, or use the disk version.
  await typeAt('KIT', 'KIT-DISK');
  await closeView(third.view, { autosave: false });
  copyFileSync(path.resolve(values.deck), path.join(workspace, 'deck.pptx'));
  const fourth = await open();
  const keep = await docs.resolveDiskChange(fourth.view.id, peer, { useDisk: false });
  const keptContent = await (await kit.ready()).request('content', { documentId: docs.views.get(fourth.view.id).documentId });
  liveView = undefined; // using the disk version retires this view; the tab reopens
  await docs.resolveDiskChange(fourth.view.id, peer, { useDisk: true });
  await closeView(fourth.view, { autosave: false });
  const fifth = await open();
  const diskContent = await (await kit.ready()).request('content', { documentId: docs.views.get(fifth.view.id).documentId });
  check('a disk change: keeping edits keeps them, using the disk replaces them', fourth.view.diskChanged === true && keep.kept && keptContent.content.includes('KIT-DISK') && !diskContent.content.includes('KIT-DISK'), { diskChanged: fourth.view.diskChanged, kept: keptContent.content.includes('KIT-DISK'), replaced: !diskContent.content.includes('KIT-DISK') });
  // Insert a picture from the toolbar: bytes go to Kit through the Host, the reference into a change.
  const mediaCount = () => execFileSync('sh', ['-c', `unzip -l "${path.join(workspace, 'deck.pptx')}" | grep -c 'ppt/media/' || true`], { encoding: 'utf8' }).trim();
  const before = Number(mediaCount());
  const picture = path.join(base, 'dot.png');
  writeFileSync(picture, Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAgAAAAICAIAAABLbSncAAAAEklEQVR4nGP4z8CAFWEXHbQSACj/P8Fu7N9hAAAAAElFTkSuQmCC', 'base64'));
  await frame.getByText('插入', { exact: true }).first().click();
  await page.waitForTimeout(300);
  const chooser = page.waitForEvent('filechooser', { timeout: 10_000 }).catch(() => null);
  await page.mouse.click(190, 75);
  const fileChooser = await chooser;
  if (fileChooser) await fileChooser.setFiles(picture);
  await page.waitForTimeout(1500);
  const inserted = await within(page.evaluate(() => window.agent('flush')), 20_000, 'flush').catch(error => ({ ok: false, message: error.message }));
  const savedPicture = await docs.save(fifth.view.id, peer).then(() => 'saved', error => error.code);
  const after = Number(mediaCount());
  check('a picture inserted from the toolbar is stored through the Host and saved into the PPTX', Boolean(fileChooser) && inserted.ok && savedPicture === 'saved' && after === before + 1, { chooser: Boolean(fileChooser), saved: savedPicture, media: [before, after] });
  if (values.shots) await page.screenshot({ path: path.join(values.shots, 'inserted.png') });

  await closeView(fifth.view, { autosave: false });
  const sixth = await open();
  const reopened = await (await kit.ready()).request('content', { documentId: docs.views.get(sixth.view.id).documentId });
  const references = text => new Set(text.match(/mop-asset:\/media\/[0-9a-f]{64}\.[a-z0-9]+/g) ?? []);
  check('the inserted picture is there after reopening', [...references(reopened.content)].some(reference => !references(diskContent.content).has(reference)), { pictures: references(reopened.content).size });
  await closeView(sixth.view, { autosave: false });

  if (values.large) {
    copyFileSync(path.resolve(values.large), path.join(workspace, 'large.pptx'));
    const opened = Date.now();
    const view = await docs.open(scope, 'large.pptx', peer);
    await page.evaluate(({ id, entry }) => window.openView(id, entry), { id: view.id, entry: `/shimo/kit/static/page/index.html?view=${view.id}` });
    await page.waitForFunction(() => window.ready || window.failed, null, { timeout: 180_000 });
    const mounted = Date.now() - opened;
    await (await kit.ready()).request('change', { documentId: view.documentId ?? docs.views.get(view.id).documentId, body: '{}' }).catch(() => {});
    const saving = Date.now();
    await docs.save(view.id, peer).catch(() => {});
    check('a large presentation opens and saves', !(await page.evaluate(() => window.failed)), { bytes: (await import('node:fs')).statSync(path.resolve(values.large)).size, openAndMountMs: mounted, saveMs: Date.now() - saving });
    await closeView(view, { autosave: false });
  }
  const fatal = errors.filter(text => !/font|favicon|ResizeObserver/i.test(text));
  check('no page errors', fatal.length === 0, fatal.slice(0, 5));
} finally {
  console.log('closing');
  await browser.close();
  console.log('browser closed');
  server.closeAllConnections?.();
  server.close();
  await docs.dispose();
  console.log('documents disposed');
  rmSync(base, { recursive: true, force: true, maxRetries: 5, retryDelay: 200 });
}
const failed = results.filter(result => !result.ok);
console.log(`${results.length - failed.length}/${results.length} passed`);
process.exitCode = failed.length ? 1 : 0;
