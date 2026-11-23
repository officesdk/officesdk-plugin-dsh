#!/usr/bin/env node
/**
 * DOCX and XLSX through office-kit's Kit in a real browser (spec §13.7 #3–4): the built package's Kit, editor
 * bundle, SDK resources, host fonts and page, served by kit-host's HTTP handler; the parent page plays the
 * DSH tab. Each suite starts from a fresh (cold) Kit work root.
 *
 *   node tooling/suite/e2e-kit-suites.mjs [--suite document|sheet|all] [--resources packages/dsh-adapter/resources]
 *        [--large-docx <file>] [--large-xlsx <file>] [--shots <dir>] [--headed]
 */
import { createServer } from 'node:http';
import { copyFileSync, mkdirSync, mkdtempSync, readFileSync, realpathSync, rmSync, statSync, writeFileSync } from 'node:fs';
import { execFileSync } from 'node:child_process';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { isDeepStrictEqual, parseArgs } from 'node:util';
import { crc32, deflateSync } from 'node:zlib';
import { chromium } from 'playwright';

const repo = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '../..');
const { values } = parseArgs({ options: { suite: { type: 'string', default: 'all' }, resources: { type: 'string', default: path.join(repo, 'packages/dsh-adapter/resources') },
  'large-docx': { type: 'string' }, 'large-xlsx': { type: 'string' }, shots: { type: 'string' }, headed: { type: 'boolean', default: false } } });
const { KitDocuments, KitProcess, createKitHttpHandler } = await import(path.join(repo, 'packages/kit-host/src/index.mjs'));
const resources = path.resolve(values.resources);
const fixtures = path.join(repo, 'test/fixtures');
/** A valid 32×32 RGB PNG (an orange square), so the suites' image decoders accept it. */
function picturePng(size = 32) {
  const chunk = (type, data) => { const length = Buffer.alloc(4); length.writeUInt32BE(data.length); const body = Buffer.concat([Buffer.from(type), data]); const crc = Buffer.alloc(4); crc.writeUInt32BE(crc32(body)); return Buffer.concat([length, body, crc]); };
  const header = Buffer.alloc(13); header.writeUInt32BE(size, 0); header.writeUInt32BE(size, 4); header[8] = 8; header[9] = 2;
  const rows = Buffer.concat(Array.from({ length: size }, () => Buffer.concat([Buffer.from([0]), Buffer.alloc(size * 3, Buffer.from([255, 140, 0]))])));
  return Buffer.concat([Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]), chunk('IHDR', header), chunk('IDAT', deflateSync(rows)), chunk('IEND', Buffer.alloc(0))]);
}
const PICTURE = picturePng();

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
  if (name === 'failed') window.failed = m;
  if (name === 'agent-result') { pending.get(m.requestId)?.(m); pending.delete(m.requestId); }
});
window.openView = (id, entry) => { view = id; window.ready = false; window.failed = undefined; window.events = []; frame.src = entry + '&revision=1&parentOrigin=' + encodeURIComponent(location.origin); };
window.agent = (op, args) => new Promise(resolve => { const requestId = 'r' + (++n); pending.set(requestId, resolve); frame.contentWindow.postMessage({ type: 'officesdk:command', v: 1, id: view, revision, command: 'agent', requestId, op, args: args ?? {} }, location.origin); });
</script>`;

const results = [];
const check = (suite, name, ok, detail) => { results.push({ suite, name, ok }); console.log(`${ok ? 'ok  ' : 'FAIL'} [${suite}] ${name}${detail !== undefined ? ` ${JSON.stringify(detail).slice(0, 300)}` : ''}`); };
const within = (promise, ms, what) => Promise.race([promise, new Promise((_, reject) => setTimeout(() => reject(new Error(`${what} timed out after ${ms} ms`)), ms))]);
const unzipText = (file, pattern) => execFileSync('sh', ['-c', `unzip -p "$1" "$2"`, 'sh', file, pattern], { encoding: 'utf8', maxBuffer: 256 * 1024 * 1024 });
const unzipOptional = (file, member) => execFileSync('sh', ['-c', 'unzip -p "$1" "$2" 2>/dev/null || true', 'sh', file, member], { encoding: 'utf8', maxBuffer: 256 * 1024 * 1024 });
const occurrences = (text, needle) => text.split(needle).length - 1;
/**
 * A MOW/MOP model for comparison, as office-kit's parity test reads it: `null` and absent members are the
 * same in BOS, and Writer stamps a per-load `meta.documentId` that never travels in a changeset.
 */
function model(json, { withoutTableNodeIds = false } = {}) {
  const parsed = JSON.parse(json, (key, value) => value === null || typeof value !== 'object' || Array.isArray(value) ? value
    : Object.fromEntries(Object.entries(value).filter(([name, member]) => member !== null
      // Writer numbers table nodes when it loads (`table-node:<n>`); proven page-only by the mount-time check below.
      && !(withoutTableNodeIds && name === 'nodeId' && typeof member === 'string' && /^table-node:\d+$/.test(member))
      // An empty attribute map is no attributes, as an absent one is.
      && !(name === 'attrs' && member && typeof member === 'object' && !Array.isArray(member) && Object.keys(member).length === 0)).sort(([a], [b]) => a.localeCompare(b))));
  if (typeof parsed.meta?.documentId === 'string' && parsed.meta.documentId.startsWith('writer-document:')) {
    delete parsed.meta.documentId;
    if (Object.keys(parsed.meta).length === 0) delete parsed.meta;
  }
  return parsed;
}
function differences(left, right, at = '$', out = []) {
  if (out.length >= 8 || Object.is(left, right)) return out;
  if (typeof left !== 'object' || typeof right !== 'object' || left === null || right === null || Array.isArray(left) !== Array.isArray(right)) {
    out.push(`${at}: ${JSON.stringify(left)?.slice(0, 80)} != ${JSON.stringify(right)?.slice(0, 80)}`);
    return out;
  }
  for (const key of new Set([...Object.keys(left), ...Object.keys(right)])) differences(left[key], right[key], Array.isArray(left) ? `${at}[${key}]` : `${at}.${key}`, out);
  return out;
}
const mediaCount = (file, dir) => Number(execFileSync('sh', ['-c', 'unzip -l "$1" | grep -c "$2" || true', 'sh', file, dir], { encoding: 'utf8' }).trim());
/**
 * Another program rewrites the file in place (same inode, as an editor saving over it): same content,
 * different bytes (a space in [Content_Types].xml).
 */
function alterOnDisk(dir, name) {
  const scratch = mkdtempSync(path.join(tmpdir(), 'kit-alter-'));
  execFileSync('sh', ['-c', 'cd "$1" && unzip -q "$2" -d z && cd z && printf " " >> "[Content_Types].xml" && zip -q -r -X ../altered.zip .', 'sh', scratch, path.join(dir, name)]);
  copyFileSync(path.join(scratch, 'altered.zip'), path.join(dir, name));
  rmSync(scratch, { recursive: true, force: true });
}

/** One suite's environment: a cold Kit work root, a Host over it, a browser page as the DSH tab. */
async function environment(browser, suite) {
  const base = realpathSync(mkdtempSync(path.join(tmpdir(), `kit-e2e-${suite}-`)));
  const workspace = path.join(base, 'ws');
  mkdirSync(workspace);
  const kit = new KitProcess({ binary: path.join(resources, 'kit', 'offline-kit-runtime'), work: path.join(base, 'kit') });
  kit.on('log', text => { if (/error|panic/i.test(text)) console.log('[kit]', text.trim().slice(0, 300)); });
  const docs = new KitDocuments({ kit, exportRoot: path.join(base, 'exports'),
    // Like WorkspaceFiles: a version that changes whenever the file is rewritten.
    authorize: async () => {}, stat: async (scope, file) => { const absolutePath = path.resolve(workspace, file); const info = statSync(absolutePath); return { absolutePath, version: `${info.mtimeMs}:${info.size}:${info.ino}` }; },
    writeAccess: async (scope, real) => real.startsWith(workspace + path.sep) ? 'allowed' : 'outside' });
  const scope = { sessionId: 's', workspaceRoot: workspace };
  const peer = { id: 'tab', principal: 'owner', signal: new AbortController().signal };
  let origin;
  const handler = createKitHttpHandler({ documents: docs, admit: () => ({ peer }), allowedOrigin: value => value === origin,
    staticRoots: { page: path.join(resources, 'kit-page'), editor: path.join(resources, 'editor'), resources: path.join(resources, 'office-kit'), fonts: path.join(resources, 'host-fonts') },
    meta: view => ({ v: 1, id: view.id, suite: view.suite, mode: view.writer ? 'edit' : 'view', name: path.basename(view.path), diskChanged: view.diskChanged, user: { uid: 'local:e2e', name: 'e2e' } }) });
  const server = createServer(async (req, res) => {
    if (req.url === '/') { res.writeHead(200, { 'content-type': 'text/html' }); res.end(PARENT); return; }
    if (!(await handler(req, res))) { res.statusCode = 404; res.end(); }
  });
  await new Promise(resolve => server.listen(0, '127.0.0.1', resolve));
  origin = `http://127.0.0.1:${server.address().port}`;
  const page = await browser.newPage({ viewport: { width: 1280, height: 800 } });
  const errors = [];
  let liveView;
  page.on('pageerror', error => errors.push(error.message.slice(0, 200)));
  page.on('console', message => { if (message.type() === 'error' && !/Failed to load resource/.test(message.text())) errors.push(message.text().slice(0, 200)); });
  page.on('response', response => { const where = new URL(response.url()).pathname; if (response.status() >= 400 && response.status() !== 416 && (!where.includes('/kit/v/') || where.includes(`/kit/v/${liveView}/`))) errors.push(`${response.status()} ${where}`); });
  await page.goto(`${origin}/`);
  const open = async (name, { as = peer, session = scope } = {}) => {
    const view = await docs.open(session, name, as);
    liveView = view.id;
    const started = Date.now();
    await page.evaluate(({ id, entry }) => window.openView(id, entry), { id: view.id, entry: `/shimo/kit/static/page/index.html?view=${view.id}` });
    await page.waitForFunction(() => window.ready || window.failed, null, { timeout: 120_000 });
    return { view, failed: await page.evaluate(() => window.failed), ms: Date.now() - started };
  };
  // Like the DSH tab: the frame goes away before its view is released.
  const closeView = async (view, options) => {
    // Requests already queued by the old iframe may finish after it is detached. They are stale once this view
    // closes, so the response probe must stop treating its expected 409s as page errors.
    if (liveView === view.id) liveView = undefined;
    await page.evaluate(() => { document.getElementById('doc').src = 'about:blank'; });
    await page.waitForTimeout(300);
    return docs.close(view.id, peer, options).catch(() => ({ closed: true }));
  };
  // Kit's record, read directly: `docs.content` after the page mounted would be a rebuild that restarts the change sequence.
  const kitContent = async view => { const read = await (await kit.ready()).request('content', { documentId: docs.views.get(view.id)?.documentId ?? view.documentId }); return { content: read.ok ? read.content : '' }; };
  const agent = (op, args) => within(page.evaluate(([o, a]) => window.agent(o, a), [op, args]), 30_000, op).catch(error => ({ ok: false, message: error.message }));
  const frame = () => page.frames().find(f => f.url().includes('/kit/static/page/'));
  const shot = async name => { if (values.shots) { mkdirSync(values.shots, { recursive: true }); await page.screenshot({ path: path.join(values.shots, `${suite}-${name}.png`) }); } };
  const dispose = async () => { await page.close(); server.close(); await docs.dispose(); rmSync(base, { recursive: true, force: true }); };
  return { base, workspace, kit, docs, scope, peer, page, errors, open, closeView, agent, frame, shot, dispose, kitContent };
}

/** Types until the editor reports a change (a click right after a mount or save can land early). */
async function typeUntilChanged(env, place, text, commit = async () => {}, type = value => env.page.keyboard.insertText(value)) {
  const before = (await env.agent('flush')).result?.editRevision ?? 0;
  for (let attempt = 0; attempt < 3; attempt++) {
    await place();
    await type(text);
    await commit();
    await env.page.waitForTimeout(300);
    const after = (await env.agent('flush')).result?.editRevision ?? 0;
    if (after > before) return after;
    await env.page.waitForTimeout(500);
  }
  throw new Error(`typing "${text}" never reached the editor`);
}

/** The shared matrix (spec §13.7): what Kit keeps, saves, recovers and refuses, for one suite. */
/** The model's way of editing: a Python library rewrites the file in place (spec §13.8). */
const PYTHON = path.join(process.env.HOME, '.dsh/dsh-runtimes/dsh-primary-runtime/dependencies/python/bin/python3');
const rewriteWithPython = (file, script) => execFileSync(PYTHON, ['-I', '-c', script, file]);

async function sharedChecks(env, suite, { name, place, savedText, typed, kitText, commit, type, rewrite }) {
  const { docs, kit, scope, peer } = env;
  // §13.8: a hand edit is written back by the autosave path (hand-over without settling, then save); then the
  // model rewrites the file; the disk check calls it someone else's write; the reopened editor has both edits.
  let auto = await env.open(name);
  await typeUntilChanged(env, place, typed('AUTO'), commit, type);
  const handOver = await env.agent('flush', { settle: false });
  const autosaved = await docs.save(auto.view.id, peer);
  const ownWrite = await docs.checkDisk(auto.view.id, peer).then(() => 'same', error => error.code);
  rewrite(path.join(env.workspace, name), typed('MODEL'));
  const external = await docs.checkDisk(auto.view.id, peer).then(() => 'same', error => error.code);
  await env.agent('flush');
  auto = await env.open(name);
  const both = kitText((await env.kitContent(auto.view)).content);
  check(suite, 'autosaved hand edits and the model\'s rewrite are both in the reopened editor', handOver.ok && autosaved.saved && ownWrite === 'same' && external === 'source-changed'
    && !auto.failed && auto.view.diskChanged !== true && both.includes(typed('AUTO')) && both.includes(typed('MODEL')) && savedText().includes(typed('AUTO')),
    { ownWrite, external, diskChanged: auto.view.diskChanged, auto: both.includes(typed('AUTO')), model: both.includes(typed('MODEL')) });
  await env.closeView(auto.view, { autosave: false });

  // Kit restart: the unsaved edit survives into the reopened document.
  let current = await env.open(name);
  await typeUntilChanged(env, place, typed('CRASH'), commit, type);
  const flushed = await env.agent('flush');
  (await kit.ready()).request('crash').catch(() => {});
  const generation = kit.generation;
  for (let i = 0; i < 100 && kit.generation === generation; i++) await new Promise(resolve => setTimeout(resolve, 50));
  await env.closeView(current.view, { autosave: false });
  current = await env.open(name);
  const recovered = await env.kitContent(current.view);
  check(suite, 'after a Kit crash the unsaved edit is still in the reopened document', flushed.ok && !current.failed && docs.views.get(current.view.id)?.unsaved === true && kitText(recovered?.content ?? '').includes(typed('CRASH')) && !savedText().includes(typed('CRASH')),
    { generation: kit.generation, unsaved: docs.views.get(current.view.id)?.unsaved });

  // A changed file: no silent overwrite; cancel keeps the edits, overwrite writes them.
  alterOnDisk(env.workspace, name);
  const refused = await docs.save(current.view.id, peer).then(() => 'saved', error => `${error.code}:${error.details?.overwrite}`);
  const keptAfterCancel = docs.views.get(current.view.id)?.unsaved === true && !savedText().includes(typed('CRASH'));
  await docs.save(current.view.id, peer, { overwrite: true });
  check(suite, 'a file changed on disk is only replaced with the user\'s overwrite; cancelling keeps the edits', refused === 'source-changed:true' && keptAfterCancel && savedText().includes(typed('CRASH')), { refused, keptAfterCancel });

  // A second conversation gets a read-only view.
  const other = { id: 'tab-2', principal: 'owner', signal: new AbortController().signal };
  const readOnly = await docs.open({ ...scope, sessionId: 's2' }, name, other);
  const denied = await docs.change(readOnly.id, other, { seq: 0, format: 'x', body: '{}' }).then(() => 'accepted', error => error.code);
  check(suite, 'a second conversation on the same file gets a read-only view', readOnly.mode === 'view' && denied === 'read-only', { mode: readOnly.mode, denied });
  await docs.close(readOnly.id, other);

  // A disk change while edits are kept: keep, then use the disk.
  await typeUntilChanged(env, place, typed('DISK'), commit, type);
  await env.agent('flush');
  await env.closeView(current.view, { autosave: false });
  alterOnDisk(env.workspace, name);
  const fourth = await env.open(name);
  const keep = await docs.resolveDiskChange(fourth.view.id, peer, { useDisk: false });
  const keptContent = await env.kitContent(fourth.view);
  await docs.resolveDiskChange(fourth.view.id, peer, { useDisk: true });
  await env.closeView(fourth.view, { autosave: false });
  const fifth = await env.open(name);
  const diskContent = await env.kitContent(fifth.view);
  check(suite, 'a disk change: keeping edits keeps them, using the disk replaces them', fourth.view.diskChanged === true && keep.kept && kitText(keptContent.content).includes(typed('DISK')) && !kitText(diskContent.content).includes(typed('DISK')),
    { diskChanged: fourth.view.diskChanged, kept: kitText(keptContent.content).includes(typed('DISK')), replaced: !kitText(diskContent.content).includes(typed('DISK')) });

  // Closing the editing tab saves it.
  await typeUntilChanged(env, place, typed('CLOSE'), commit, type);
  await env.agent('flush');
  const closed = await env.closeView(fifth.view);
  check(suite, 'closing the editing tab autosaves its edits', closed.saved === true && savedText().includes(typed('CLOSE')), closed);

  // Another conversation holds the writer: this tab's view is read only and its editor must still mount.
  const holder = { id: 'tab-hold', principal: 'owner', signal: new AbortController().signal };
  const held = await docs.open({ ...scope, sessionId: 's3' }, name, holder);
  const readOnlyTab = await env.open(name);
  check(suite, 'a read-only view mounts its editor', held.mode === 'edit' && readOnlyTab.view.mode === 'view' && !readOnlyTab.failed, { mode: readOnlyTab.view.mode, failed: readOnlyTab.failed?.message });
  await env.closeView(readOnlyTab.view, { autosave: false });
  await docs.close(held.id, holder, { autosave: false });
}

async function largeFile(env, suite, file) {
  if (!file) return;
  const name = `large${path.extname(file)}`;
  copyFileSync(path.resolve(file), path.join(env.workspace, name));
  const started = Date.now();
  const opened = await env.open(name);
  const mounted = Date.now() - started;
  const saving = Date.now();
  const saved = await env.docs.save(opened.view.id, env.peer).then(result => result.saved, error => error.code);
  check(suite, 'a large file opens and saves', !opened.failed && saved === true, { bytes: statSync(path.resolve(file)).size, openAndMountMs: mounted, saveMs: Date.now() - saving, saved, failed: opened.failed?.message });
  await env.closeView(opened.view, { autosave: false });
}

async function documentSuite(browser) {
  const suite = 'document';
  const env = await environment(browser, suite);
  const { docs, kit, peer, page } = env;
  try {
    copyFileSync(path.join(fixtures, 'kit-edit.docx'), path.join(env.workspace, 'doc.docx'));
    copyFileSync(path.join(fixtures, 'sample.docx'), path.join(env.workspace, 'tracked.docx'));
    const savedText = () => unzipText(path.join(env.workspace, 'doc.docx'), 'word/document.xml');
    const kitText = content => content;
    const typed = tag => `KIT-${tag}`;
    // Click into the body text of the first page, then go to the end of that paragraph.
    const place = async () => {
      const box = await env.frame().locator('[data-role="editor-mount"]').first().boundingBox().catch(() => null)
        ?? await page.locator('iframe').boundingBox();
      await page.mouse.click(box.x + box.width / 2, box.y + Math.min(box.height * 0.45, 360));
      await page.keyboard.press('End');
    };

    const first = await env.open('doc.docx');
    check(suite, 'a cold Kit work root opens the document within the 30 s font budget', !first.failed && first.ms < 30_000, { ms: first.ms, failed: first.failed?.message });
    if (first.failed) throw new Error('the document did not mount');
    await env.shot('mounted');
    // Before any edit the page holds what Kit imported, apart from the editor's own table node numbering.
    const imported = await (await kit.ready()).request('content', { documentId: docs.views.get(first.view.id).documentId });
    const mountedModel = await env.agent('snapshot');
    const atMount = imported.ok && mountedModel.ok ? differences(model(imported.content), model(mountedModel.result.content)) : ['unavailable'];
    const remaining = imported.ok && mountedModel.ok ? differences(model(imported.content, { withoutTableNodeIds: true }), model(mountedModel.result.content, { withoutTableNodeIds: true })) : ['unavailable'];
    check(suite, 'at mount the page model is Kit\'s import except Writer\'s load-time table node ids', remaining.length === 0 && atMount.every(line => line.includes('table-node:')), { remaining, atMount: atMount.length });
    // The zoom menu marks the active mode: "fit page width" must be the selected item.
    await page.waitForTimeout(1000);
    await env.frame().locator('.editor-statusbar__zoom-value').click();
    const fitItem = env.frame().locator('[data-menu-binding-token="root:fit-page-width"]');
    await fitItem.waitFor({ timeout: 5000 }).catch(() => {});
    const selected = await fitItem.getAttribute('data-selected').catch(() => null);
    const zoomText = (await env.frame().locator('.editor-statusbar__zoom-value').textContent().catch(() => '')).trim();
    await page.keyboard.press('Escape');
    check(suite, 'the document opens at page width, like the preview', selected === 'true', { selected, zoom: zoomText });

    await typeUntilChanged(env, place, typed('HAND'));
    const flushed = await env.agent('flush');
    check(suite, 'a hand edit reaches Kit as numbered changes and the page reports it', flushed.ok && flushed.result.editRevision > 0 && docs.views.get(first.view.id).unsaved, flushed.ok ? flushed.result : flushed);
    check(suite, 'the file is untouched until a save', !savedText().includes(typed('HAND')));
    const kitContent = await (await kit.ready()).request('content', { documentId: docs.views.get(first.view.id).documentId });
    const snapshot = await env.agent('snapshot');
    const strip = { withoutTableNodeIds: true };
    const diff = kitContent.ok && snapshot.ok ? differences(model(kitContent.content, strip), model(snapshot.result.content, strip)) : ['unavailable'];
    check(suite, 'Kit\'s composed content equals the page\'s own model', kitContent.ok && snapshot.ok && diff.length === 0, { kit: kitContent.ok, page: snapshot.ok ? 'ok' : snapshot.message, diff });

    // A composition still open when the save starts is committed by prepare, once.
    await place();
    const cdp = await page.context().newCDPSession(page);
    await cdp.send('Input.imeSetComposition', { text: 'KIT-IME', selectionStart: 7, selectionEnd: 7 });
    await page.waitForTimeout(300);
    const imeFlush = await env.agent('flush');
    await docs.save(first.view.id, peer);
    await cdp.detach().catch(() => {});
    const saved = savedText();
    check(suite, 'a composition in progress when saving is saved exactly once', imeFlush.ok && occurrences(saved, 'KIT-IME') === 1, { flush: imeFlush.ok, count: occurrences(saved, 'KIT-IME') });
    check(suite, 'save writes the edit into word/document.xml', saved.includes(typed('HAND')) && docs.views.get(first.view.id).unsaved === false);

    await env.closeView(first.view, { autosave: false });
    const second = await env.open('doc.docx');
    const reopened = await env.kitContent(second.view);
    check(suite, 'reopening shows the saved edit', !second.failed && reopened.content.includes(typed('HAND')));
    await typeUntilChanged(env, place, typed('AGAIN'));
    await env.agent('flush');
    await docs.save(second.view.id, peer);
    check(suite, 'a second edit after reopening saves too', savedText().includes(typed('AGAIN')) && savedText().includes(typed('HAND')));

    // A picture from the toolbar is stored through the Host and saved into the DOCX.
    const mediaBefore = mediaCount(path.join(env.workspace, 'doc.docx'), 'word/media/');
    await place();
    const inserted = await insertPicture(env, async () => {
      const f = env.frame();
      await f.locator('text=插入').first().click();
      await page.waitForTimeout(400);
      await f.locator('button:visible', { hasText: '图片' }).first().click();
      await f.locator('[data-menu-binding-token="root:insert-picture-local"]').click();
    });
    await env.agent('flush');
    await docs.save(second.view.id, peer).catch(error => console.log('save after insert:', error.code, error.message));
    const mediaAfter = mediaCount(path.join(env.workspace, 'doc.docx'), 'word/media/');
    check(suite, 'a picture inserted from the toolbar is stored through the Host and saved into the DOCX', inserted && mediaAfter === mediaBefore + 1, { inserted, media: [mediaBefore, mediaAfter] });
    await env.closeView(second.view, { autosave: false });
    const third = await env.open('doc.docx');
    const withPicture = await env.kitContent(third.view);
    check(suite, 'the inserted picture is there after reopening', /modoc-asset:\/image\//.test(withPicture.content) && mediaAfter > mediaBefore);
    await env.closeView(third.view, { autosave: false });

    const rewrite = (file, text) => rewriteWithPython(file, `import docx,sys\np=sys.argv[1]\nd=docx.Document(p)\nd.add_paragraph('${text}')\nd.save(p)`);
    await sharedChecks(env, suite, { name: 'doc.docx', place, savedText, typed, kitText, rewrite });

    // Tracked revisions: since office-kit 0.1.2 (Writer 0.1.1) the editor shows them without changing the document at
    // mount, so the view opens with nothing unsaved, the page holds Kit's import and the file stays as it was.
    const trackedFile = path.join(env.workspace, 'tracked.docx');
    const trackedBefore = readFileSync(trackedFile);
    const tracked = await env.open('tracked.docx');
    const trackedView = docs.views.get(tracked.view.id);
    const trackedUnsaved = trackedView?.unsaved;
    const trackedRevision = tracked.failed ? undefined : (await env.agent('flush')).result?.editRevision;
    const trackedImport = trackedView && await (await kit.ready()).request('content', { documentId: trackedView.documentId });
    const trackedModel = trackedView && await env.agent('snapshot');
    const trackedDiff = trackedImport?.ok && trackedModel?.ok ? differences(model(trackedImport.content, { withoutTableNodeIds: true }), model(trackedModel.result.content, { withoutTableNodeIds: true })) : ['unavailable'];
    if (trackedView) await env.closeView(tracked.view, { autosave: false });
    check(suite, 'a DOCX with tracked revisions opens unchanged and its file stays as it was', !tracked.failed && trackedUnsaved === false && trackedRevision === 0 && trackedDiff.length === 0
      && readFileSync(trackedFile).equals(trackedBefore) && unzipText(trackedFile, 'word/settings.xml').includes('<w:trackRevisions/>'),
      { failed: tracked.failed?.message, unsaved: trackedUnsaved, editRevision: trackedRevision, diff: trackedDiff.slice(0, 3) });

    await largeFile(env, suite, values['large-docx']);
    check(suite, 'no page errors', env.errors.length === 0, env.errors.slice(0, 5));
  } finally { await env.dispose(); }
}

/** Opens the suite's picture command, answers the file chooser with a PNG, and waits for the Host store. */
async function insertPicture(env, openCommand) {
  const picture = path.join(env.base, 'picture.png');
  writeFileSync(picture, PICTURE);
  const chooser = env.page.waitForEvent('filechooser', { timeout: 10_000 }).catch(() => null);
  const stored = env.page.waitForResponse(response => /\/kit\/v\/[^/]+\/media\?extension=/.test(response.url()) && response.request().method() === 'POST', { timeout: 20_000 }).catch(() => null);
  await openCommand().catch(error => console.log('insert command:', error.message));
  const fileChooser = await chooser;
  if (!fileChooser) { console.log('insert: no file chooser'); return false; }
  await fileChooser.setFiles(picture);
  const response = await stored;
  if (response?.status() !== 200) console.log('insert: store', response?.status() ?? 'never called');
  await env.page.waitForTimeout(1500);
  return response?.status() === 200;
}

async function sheetSuite(browser) {
  const suite = 'sheet';
  const env = await environment(browser, suite);
  const { docs, kit, peer, page } = env;
  try {
    copyFileSync(path.join(fixtures, 'kit-edit.xlsx'), path.join(env.workspace, 'book.xlsx'));
    // Strings may be shared or inline, depending on who wrote the file last.
    const savedText = () => unzipOptional(path.join(env.workspace, 'book.xlsx'), 'xl/sharedStrings.xml') + unzipText(path.join(env.workspace, 'book.xlsx'), 'xl/worksheets/sheet1.xml');
    const kitText = content => content;
    const typed = tag => `KIT-${tag}`;
    // Go to a cell through the name box (the input that shows the active cell's address), then type into it.
    const goTo = async cell => {
      const box = await env.frame().evaluate(() => {
        const input = [...document.querySelectorAll('input')].find(node => node.offsetParent && /^[A-Z]{1,3}[0-9]{1,7}$/.test(node.value));
        const rect = input?.getBoundingClientRect();
        return rect && { x: rect.x + rect.width / 2, y: rect.y + rect.height / 2 };
      });
      if (!box) throw new Error('the sheet has no name box');
      await page.mouse.click(box.x, box.y);
      await page.keyboard.press('Meta+A');
      await page.keyboard.type(cell);
      await page.keyboard.press('Enter');
      await page.waitForTimeout(200);
    };
    const write = async (cell, text) => { await goTo(cell); await page.keyboard.type(text); await page.keyboard.press('Enter'); };
    let row = 5;
    const place = async () => { await goTo(`A${row++}`); };

    const first = await env.open('book.xlsx');
    check(suite, 'a cold Kit work root opens the workbook', !first.failed, { ms: first.ms, failed: first.failed?.message });
    if (first.failed) throw new Error('the workbook did not mount');
    await env.shot('mounted');
    const iconFonts = await env.frame().evaluate(async () => { await document.fonts.ready; return ['LockIcons', 'DateMentionIcons'].map(family => [family, [...document.fonts].some(face => face.family.replaceAll('"', '') === family)]); });
    check(suite, 'the Host supplies the Sheet icon faces (LockIcons, DateMentionIcons)', iconFonts.every(([, found]) => found), iconFonts);

    for (let attempt = 0; attempt < 3 && !(await env.kitContent(first.view)).content.includes(typed('HAND')); attempt++) { await write('C1', typed('HAND')); await env.agent('flush'); }
    await env.shot('cell-edit');
    const flushed = await env.agent('flush');
    check(suite, 'a cell edit reaches Kit as numbered changes and the page reports it', flushed.ok && flushed.result.editRevision > 0 && docs.views.get(first.view.id).unsaved, flushed.ok ? flushed.result : flushed);
    check(suite, 'the file is untouched until a save', !savedText().includes(typed('HAND')));

    // A formula whose input changes: prepare stabilizes the formula cache, and the saved file carries it.
    // A formula input changes: the saved file must carry the recalculated cache (B3 = A3*3 = 15).
    await write('A3', '5');
    await page.waitForTimeout(1500);
    await env.agent('flush');
    await docs.save(first.view.id, peer);
    const sheetXml = unzipText(path.join(env.workspace, 'book.xlsx'), 'xl/worksheets/sheet1.xml');
    const b3 = /<c r="B3"[^>]*>(.*?)<\/c>/.exec(sheetXml)?.[1] ?? '';
    check(suite, 'save writes the cell edits into the XLSX', savedText().includes(typed('HAND')) && /<c r="A3"[^>]*><v>5<\/v>/.test(sheetXml) && /<f>A3\*3<\/f>/.test(b3),
      { row1: /<row r="1"[^>]*>.*?<\/row>/.exec(sheetXml)?.[0], shared: unzipOptional(path.join(env.workspace, 'book.xlsx'), 'xl/sharedStrings.xml').slice(-200) });
    // Known limitation (spec §13.6, Acceptance v8): the recalculated cache is not written; Excel recalculates on open.
    const calc = /<calcPr[^>]*>/.exec(unzipText(path.join(env.workspace, 'book.xlsx'), 'xl/workbook.xml'))?.[0] ?? '';
    check(suite, 'the formula is kept and the workbook recalculates on open (cache not written: known limitation)', /<f>A3\*3<\/f>/.test(b3) && /calcId="0"/.test(calc), { b3, calc, cacheWritten: /<v>15<\/v>/.test(b3) });

    await env.closeView(first.view, { autosave: false });
    const second = await env.open('book.xlsx');
    const reopened = await env.kitContent(second.view);
    check(suite, 'reopening shows the saved edit', !second.failed && reopened.content.includes(typed('HAND')));

    await goTo('D2');
    const inserted = await insertPicture(env, async () => {
      const f = env.frame();
      await f.locator('text=插入').first().click();
      await page.waitForTimeout(400);
      await env.shot('insert-menu');
      // The menu item is a canvas-era element Playwright's text locator does not resolve; click where it is drawn.
      const item = await f.evaluate(() => {
        const node = [...document.querySelectorAll('*')].find(element => element.childElementCount === 0 && element.textContent.trim() === '浮动图片' && element.getBoundingClientRect().width > 0);
        const rect = node?.getBoundingClientRect();
        return rect && { x: rect.x + rect.width / 2, y: rect.y + rect.height / 2 };
      });
      if (!item) throw new Error('no 浮动图片 menu item');
      await page.mouse.click(item.x, item.y);
      await page.waitForTimeout(400);
      const local = f.locator('text=/本地上传|从本地|本地图片/').first();
      if (await local.count()) await local.click();
    });
    await env.agent('flush');
    await docs.save(second.view.id, peer).catch(error => console.log('save after insert:', error.code, error.message));
    const media = mediaCount(path.join(env.workspace, 'book.xlsx'), 'xl/media/');
    check(suite, 'a picture inserted from the toolbar is stored through the Host and saved into the XLSX', inserted && media >= 1, { inserted, media });
    await env.closeView(second.view, { autosave: false });
    // The picture is drawn from this view's media/ route by its Kit reference. (With this Sheet SDK a PNG,
    // floating or in a cell, loads through the editor's resolved URL; the worker bridge is covered by unit tests.)
    const picture = page.waitForResponse(response => /\/kit\/v\/[^/]+\/media\/modoc-assets%3A%2Fmedia%2F[A-Za-z0-9._-]+/.test(response.url()) && response.status() === 200, { timeout: 30_000 }).then(response => response.headers()['content-type'], () => undefined);
    const third = await env.open('book.xlsx');
    const shown = await picture;
    check(suite, 'the picture is there after reopening, read through this view\'s media route', shown === 'image/png' && /modoc-assets:\/media\//.test((await env.kitContent(third.view)).content), { contentType: shown });
    await env.closeView(third.view, { autosave: false });

    const rewrite = (file, text) => rewriteWithPython(file, `import openpyxl,sys\np=sys.argv[1]\nw=openpyxl.load_workbook(p)\nw.active['F9']='${text}'\nw.save(p)`);
    await sharedChecks(env, suite, { name: 'book.xlsx', place, savedText, typed, kitText, commit: () => page.keyboard.press('Enter'), type: value => page.keyboard.type(value), rewrite });
    await largeFile(env, suite, values['large-xlsx']);
    check(suite, 'no page errors', env.errors.length === 0, env.errors.slice(0, 5));
  } finally { await env.dispose(); }
}

const browser = await chromium.launch({ headless: !values.headed });
try {
  for (const [name, run] of [['document', documentSuite], ['sheet', sheetSuite]]) {
    if (values.suite !== 'all' && values.suite !== name) continue;
    try { await run(browser); } catch (error) { check(name, `the ${name} run finished`, false, error.message); }
  }
} finally { await browser.close(); }
const failed = results.filter(result => !result.ok);
console.log(`${results.length - failed.length}/${results.length} passed`);
process.exit(failed.length ? 1 : 0);
