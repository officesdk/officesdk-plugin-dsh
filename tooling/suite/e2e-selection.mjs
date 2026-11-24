#!/usr/bin/env node
/**
 * Selection rewrite (spec 261109-selection-ai-edit) through office-kit's Kit in a real browser: the built package's Kit,
 * editor bundle, SDK resources and page, served by kit-host's HTTP handler; the parent page plays the DSH tab (it answers
 * the page's selection-request and delivers selection_rewrite as the Host would). Word: a selection inside a paragraph
 * (⌘J, Esc sends nothing, rewrite, undo), a selection over paragraphs rewritten into more paragraphs (and back); Sheet: a
 * block of cells rewritten and undone, values and formulas restored.
 *
 *   node tooling/suite/e2e-selection.mjs [--resources packages/dsh-adapter/resources] [--kit-page <dir>] [--kit <binary>] [--shots <dir>] [--headed]
 */
import { createServer } from 'node:http';
import { copyFileSync, mkdirSync, mkdtempSync, realpathSync, rmSync, statSync } from 'node:fs';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { parseArgs } from 'node:util';
import { chromium } from 'playwright';

const repo = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '../..');
const { values } = parseArgs({ options: { resources: { type: 'string', default: path.join(repo, 'packages/dsh-adapter/resources') }, 'kit-page': { type: 'string' }, kit: { type: 'string' }, shots: { type: 'string' }, headed: { type: 'boolean', default: false } } });
const { KitDocuments, KitProcess, createKitHttpHandler } = await import(path.join(repo, 'packages/kit-host/src/index.mjs'));
const resources = path.resolve(values.resources);
const fixtures = path.join(repo, 'test/fixtures');

const PARENT = `<!doctype html><meta charset="utf-8"><style>html,body{margin:0;height:100%}iframe{border:0;width:100%;height:100%}</style>
<iframe id="doc"></iframe><script type="module">
const frame = document.getElementById('doc');
const pending = new Map(); let view, revision = 1, n = 0;
window.requests = []; window.updates = [];
addEventListener('message', event => {
  const m = event.data;
  if (event.source !== frame.contentWindow || event.origin !== location.origin || !m || m.v !== 1 || m.id !== view || typeof m.type !== 'string') return;
  const name = m.type.replace('officesdk:', '');
  if (name === 'ready') window.ready = true;
  if (name === 'failed') window.failed = m;
  if (name === 'selection-request') window.requests.push(m);
  if (name === 'selection-update') window.updates.push(m);
  if (name === 'agent-result') { pending.get(m.requestId)?.(m); pending.delete(m.requestId); }
});
window.openView = (id, entry) => { view = id; window.ready = false; window.failed = undefined; frame.src = entry + '&revision=1&parentOrigin=' + encodeURIComponent(location.origin); };
window.agent = (op, args) => new Promise(resolve => { const requestId = 'r' + (++n); pending.set(requestId, resolve); frame.contentWindow.postMessage({ type: 'officesdk:command', v: 1, id: view, revision, command: 'agent', requestId, op, args: args ?? {} }, location.origin); });
</script>`;

const results = [];
const check = (suite, name, ok, detail) => { results.push({ suite, name, ok }); console.log(`${ok ? 'ok  ' : 'FAIL'} [${suite}] ${name}${detail !== undefined ? ` ${JSON.stringify(detail).slice(0, 400)}` : ''}`); };

async function environment(browser, suite) {
  const base = realpathSync(mkdtempSync(path.join(tmpdir(), `sel-e2e-${suite}-`)));
  const workspace = path.join(base, 'ws');
  mkdirSync(workspace);
  const kit = new KitProcess({ binary: values.kit ?? path.join(resources, 'kit', 'offline-kit-runtime'), work: path.join(base, 'kit') });
  const docs = new KitDocuments({ kit, exportRoot: path.join(base, 'exports'), authorize: async () => {},
    stat: async (scope, file) => { const absolutePath = path.resolve(workspace, file); const info = statSync(absolutePath); return { absolutePath, version: `${info.mtimeMs}:${info.size}:${info.ino}` }; },
    writeAccess: async (scope, real) => real.startsWith(workspace + path.sep) ? 'allowed' : 'outside' });
  const scope = { sessionId: 's', workspaceRoot: workspace };
  const peer = { id: 'tab', principal: 'owner', signal: new AbortController().signal };
  let origin;
  const handler = createKitHttpHandler({ documents: docs, admit: () => ({ peer }), allowedOrigin: value => value === origin,
    staticRoots: { page: values['kit-page'] ?? path.join(resources, 'kit-page'), editor: path.join(resources, 'editor'), resources: path.join(resources, 'office-kit'), fonts: path.join(resources, 'host-fonts') },
    meta: view => ({ v: 1, id: view.id, suite: view.suite, mode: view.writer ? 'edit' : 'view', name: path.basename(view.path), diskChanged: view.diskChanged, user: { uid: 'local:e2e', name: 'e2e' } }) });
  const server = createServer(async (req, res) => {
    if (req.url === '/') { res.writeHead(200, { 'content-type': 'text/html' }); res.end(PARENT); return; }
    if (!(await handler(req, res))) { res.statusCode = 404; res.end(); }
  });
  await new Promise(resolve => server.listen(0, '127.0.0.1', resolve));
  origin = `http://127.0.0.1:${server.address().port}`;
  const page = await browser.newPage({ viewport: { width: 1100, height: 760 } });
  const errors = [];
  page.on('pageerror', error => errors.push(error.message.slice(0, 200)));
  page.on('console', message => { if (/officesdk/.test(message.text())) console.log(`[page ${message.type()}] ${message.text().slice(0, 300)}`); });
  await page.goto(`${origin}/`);
  const open = async name => {
    copyFileSync(path.join(fixtures, name), path.join(workspace, name));
    const view = await docs.open(scope, name, peer, { mode: 'edit' });
    await page.evaluate(({ id, entry }) => window.openView(id, entry), { id: view.id, entry: `/shimo/kit/static/page/index.html?view=${view.id}` });
    await page.waitForFunction(() => window.ready || window.failed, null, { timeout: 120_000 });
    await page.waitForTimeout(2500);
    return { view, failed: await page.evaluate(() => window.failed) };
  };
  const agent = (op, args) => page.evaluate(([o, a]) => window.agent(o, a), [op, args]);
  const frame = () => page.frames().find(f => f.url().includes('/kit/static/page/'));
  const shown = mark => frame().evaluate(m => { const node = document.querySelector(`[data-shimo-selection="${m}"]`); return Boolean(node && node.style.display !== 'none'); }, mark);
  const waitShown = async (mark, on = true, ms = 4000) => { for (let waited = 0; waited < ms; waited += 100) { if (await shown(mark) === on) return true; await page.waitForTimeout(100); } return false; };
  const click = mark => frame().evaluate(m => document.querySelector(`[data-shimo-selection="${m}"]`)?.click(), mark);
  const requests = () => page.evaluate(() => window.requests.slice());
  /** Waits until the Keep / Undo bar says `text` (the undo finished, with its animation). */
  const resultSays = async (text, ms = 15_000) => { for (let waited = 0; waited < ms; waited += 200) { if (await frame().evaluate(t => document.querySelector('[data-shimo-selection="result"]')?.textContent?.includes(t), text)) return true; await page.waitForTimeout(200); } return false; };
  const shot = async name => { if (values.shots) { mkdirSync(values.shots, { recursive: true }); await page.screenshot({ path: path.join(values.shots, `${suite}-${name}.png`) }); } };
  const frameBox = async () => page.locator('iframe').boundingBox();
  const dispose = async () => { await page.close(); server.close(); await docs.dispose(); rmSync(base, { recursive: true, force: true }); };
  return { page, open, agent, frame, shown, waitShown, click, requests, resultSays, shot, frameBox, errors, dispose };
}

/** Sends the request typed into the box, answers it as the DSH tab does, and returns it. */
async function sendRequest(env, instruction, id) {
  const before = (await env.requests()).length;
  // Typed like a user: the editor must not take the keyboard away from the box.
  await env.frame().locator('[data-shimo-selection="input"]').click();
  await env.page.keyboard.type(instruction);
  const typed = await env.frame().evaluate(() => document.querySelector('[data-shimo-selection="input"]')?.value);
  if (typed !== instruction) console.log(`[debug] the box holds ${JSON.stringify(typed)} after typing`);
  await env.page.keyboard.press('Enter');
  await env.page.waitForFunction(count => window.requests.length > count, before, { timeout: 5000 }).catch(async error => {
    console.log('[debug]', await env.frame().evaluate(() => ({ active: document.activeElement?.outerHTML?.slice(0, 120), value: document.querySelector('[data-shimo-selection="input"]')?.value, hint: document.querySelector('[data-shimo-selection="hint"]')?.textContent })));
    await env.shot('debug-send');
    throw error;
  });
  const request = (await env.requests()).at(-1);
  await env.agent('selectionStatus', { requestId: request.requestId, ok: true, id });
  return request;
}

async function documentSuite(browser) {
  const suite = 'document';
  const env = await environment(browser, suite);
  try {
    const { failed } = await env.open('kit-edit.docx');
    check(suite, 'the document opens for editing', !failed, failed);
    if (failed) return;
    const read = async () => (await env.agent('inspect')).result.paragraphs.map(paragraph => paragraph.text);
    const original = await read();
    const lines = await env.frame().evaluate(() => Array.from(document.querySelectorAll('text[data-writer-document-content]')).slice(0, 4).map(element => { const box = element.getBoundingClientRect(); return { x: box.left, y: box.top, h: box.height }; }));
    const frame = await env.frameBox();
    // Select the first four characters of the first paragraph with the keyboard.
    await env.page.mouse.click(frame.x + lines[0].x + 2, frame.y + lines[0].y + lines[0].h / 2);
    await env.page.keyboard.press('Home');
    for (let i = 0; i < 4; i++) await env.page.keyboard.press('Shift+ArrowRight');
    check(suite, 'a selection shows the AI rewrite button', await env.waitShown('chip'));
    await env.shot('1-chip');
    // ⌘J opens the box; Esc closes it and nothing is sent.
    await env.page.keyboard.press('Meta+j');
    check(suite, '⌘J opens the request box', await env.waitShown('panel'));
    await env.frame().locator('[data-shimo-selection="input"]').fill('改得更正式');
    await env.shot('2-panel');
    await env.frame().locator('[data-shimo-selection="input"]').press('Escape');
    check(suite, 'Esc closes the box and sends nothing', await env.waitShown('panel', false) && (await env.requests()).length === 0);
    // Send for real.
    await env.waitShown('chip');
    await env.click('chip');
    await env.waitShown('panel');
    const request = await sendRequest(env, '改得更正式', 'sel-0000a001');
    const firstPart = Array.from(original[0]).slice(0, 4).join('');
    check(suite, 'the request carries the selection by paragraph and code point', request.kind === 'docx' && request.original === firstPart && request.anchor.paragraphs.join() === '0,0' && request.anchor.start === 0 && request.anchor.end === 4 && request.instruction === '改得更正式', request);
    check(suite, 'the selection is marked as waiting for the AI', await env.waitShown('pending'));
    await env.shot('3-pending');
    const text = '正式样稿😀，经审定';
    const rewrite = await env.agent('rewrite', { id: 'sel-0000a001', kind: 'docx', anchor: request.anchor, original: request.original, text });
    const after = await read();
    const expected = `${text}${Array.from(original[0]).slice(4).join('')}`;
    check(suite, 'selection_rewrite replaces only the selected text', rewrite.ok && after[0] === expected && after.slice(1).join('\n') === original.slice(1).join('\n'), { rewrite, first: after[0] });
    check(suite, 'Keep / Undo is offered after the rewrite and the waiting mark is gone', await env.waitShown('result') && !(await env.shown('pending')));
    await env.shot('4-rewritten');
    await env.click('revert');
    await env.resultSays('已撤回');
    const reverted = await read();
    check(suite, 'Undo puts the original text back', reverted[0] === original[0], reverted[0]);
    check(suite, 'the tab hears which selection was undone', (await env.page.evaluate(() => window.updates.map(update => `${update.selection}:${update.state}`))).includes('sel-0000a001:reverted'));
    await env.shot('5-reverted');
    // A second rewrite of the same text, stale: the text changed meanwhile.
    const stale = await env.agent('rewrite', { id: 'sel-0000a002', kind: 'docx', anchor: request.anchor, original: '不是这段', text: 'x' });
    check(suite, 'a selection whose text changed is refused as source-changed', !stale.ok && stale.code === 'source-changed', stale);
    // Over paragraphs: "123\n4" becomes three paragraphs, and back.
    const from = original.indexOf('123');
    const at = async index => (await env.frame().evaluate(i => { const box = document.querySelectorAll('text[data-writer-document-content]')[i].getBoundingClientRect(); return { x: box.left, y: box.top, h: box.height }; }, index));
    const [one, two] = [await at(1), await at(2)];
    await env.page.mouse.move(frame.x + one.x + 1, frame.y + one.y + one.h / 2);
    await env.page.mouse.down();
    await env.page.mouse.move(frame.x + two.x + 9, frame.y + two.y + two.h / 2, { steps: 8 });
    await env.page.mouse.up();
    await env.waitShown('chip');
    await env.click('chip');
    await env.waitShown('panel');
    const multi = await sendRequest(env, '拆成三段', 'sel-0000a003');
    check(suite, 'a selection over two paragraphs is anchored on both', multi.anchor.paragraphs.join() === `${from},${from + 1}` && multi.original === '123\n4', multi);
    const split = await env.agent('rewrite', { id: 'sel-0000a003', kind: 'docx', anchor: multi.anchor, original: multi.original, text: '一\n二\n三' });
    const parts = await read();
    check(suite, 'new text with line breaks becomes paragraphs; the rest stays', split.ok && parts.slice(from, from + 3).join('|') === '一|二|三56' && parts.length === original.length + 1 && parts.slice(from + 3).join('\n') === original.slice(from + 2).join('\n'), { split, parts: parts.slice(0, 6) });
    await env.shot('6-paragraphs');
    await env.click('revert');
    await env.resultSays('已撤回');
    const back = await read();
    check(suite, 'Undo puts the paragraphs back', back.join('\n') === original.join('\n'), back.slice(0, 6));
    check(suite, 'no page errors', env.errors.length === 0, env.errors);
  } finally { await env.dispose(); }
}

async function sheetSuite(browser) {
  const suite = 'sheet';
  const env = await environment(browser, suite);
  try {
    const { failed } = await env.open('kit-edit.xlsx');
    check(suite, 'the spreadsheet opens for editing', !failed, failed);
    if (failed) return;
    const frame = await env.frameBox();
    // A formula in the block's first cell, typed like a user, so the rewrite must replace it and the undo bring it back.
    await env.page.mouse.click(frame.x + 230, frame.y + 230);
    await env.page.keyboard.type('=2+3');
    await env.page.keyboard.press('Enter');
    await env.page.waitForTimeout(500);
    // Click that cell, then extend to 2 rows × 3 columns with the keyboard.
    await env.page.mouse.click(frame.x + 230, frame.y + 230);
    await env.page.keyboard.press('Shift+ArrowRight');
    await env.page.keyboard.press('Shift+ArrowRight');
    await env.page.keyboard.press('Shift+ArrowDown');
    check(suite, 'a block of cells shows the AI rewrite button', await env.waitShown('chip'));
    await env.shot('1-chip');
    await env.click('chip');
    check(suite, 'the button opens the request box', await env.waitShown('panel'));
    await env.shot('1b-panel');
    const request = await sendRequest(env, '翻成英文', 'sel-0000b001');
    check(suite, 'the request carries the 2×3 block with its cells and the formula', request.kind === 'xlsx' && request.anchor.rowCount === 2 && request.anchor.columnCount === 3 && request.cells.length === 2 && request.cells[0].length === 3 && /2\+3/.test(request.cells[0][0].formula ?? ''), { anchor: request.anchor, first: request.cells[0][0] });
    await env.shot('2-pending');
    const values = [['Alpha', 'Beta', 7], ['Gamma', '=1+2', 'Delta']];
    const wrong = await env.agent('rewrite', { id: 'sel-0000b001', kind: 'xlsx', anchor: request.anchor, cells: request.cells, values: [['x']] });
    check(suite, 'values of another size are refused', !wrong.ok && wrong.code === 'invalid-edit', wrong);
    const rewrite = await env.agent('rewrite', { id: 'sel-0000b001', kind: 'xlsx', anchor: request.anchor, cells: request.cells, values });
    const cells = async () => {
      const book = (await env.agent('inspect')).result;
      const sheet = book.sheets.find(item => item.name === request.anchor.sheet);
      const { row, column } = sheet.cells.range;
      return Array.from({ length: 2 }, (_, r) => Array.from({ length: 3 }, (_, c) => { const cell = sheet.cells.rows[request.anchor.row - row + r]?.[request.anchor.column - column + c]; return cell?.formula ? `=${String(cell.formula).replace(/^=/, '')}` : cell?.text ?? ''; }));
    };
    const written = await cells();
    check(suite, 'selection_rewrite writes exactly the block', rewrite.ok && JSON.stringify(written) === JSON.stringify([['Alpha', 'Beta', '7'], ['Gamma', '=1+2', 'Delta']]), { rewrite, written });
    check(suite, 'Keep / Undo is offered', await env.waitShown('result'));
    await env.shot('3-rewritten');
    await env.click('revert');
    await env.resultSays('已撤回');
    const back = await cells();
    const before = request.cells.map(row => row.map(cell => cell.formula ? `=${cell.formula}` : cell.text));
    check(suite, 'Undo puts the cells back', JSON.stringify(back) === JSON.stringify(before), { back, before });
    await env.shot('4-reverted');
    check(suite, 'no page errors', env.errors.length === 0, env.errors);
  } finally { await env.dispose(); }
}

const browser = await chromium.launch({ headless: !values.headed });
try { await documentSuite(browser); await sheetSuite(browser); }
finally { await browser.close(); }
const failedChecks = results.filter(result => !result.ok);
console.log(`${results.length - failedChecks.length}/${results.length} checks passed`);
process.exit(failedChecks.length ? 1 : 0);
