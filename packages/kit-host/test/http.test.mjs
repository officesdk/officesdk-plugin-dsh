import test from 'node:test';
import assert from 'node:assert/strict';
import { mkdirSync, symlinkSync, writeFileSync } from 'node:fs';
import { createServer } from 'node:http';
import path from 'node:path';
import { KIT_CSP, createKitHttpHandler, safeRelative, singleRange } from '../src/index.mjs';
import { owner, setup } from './helpers.mjs';

async function serve(t) {
  const env = setup(t);
  const statics = path.join(env.base, 'static');
  mkdirSync(path.join(statics, 'page'), { recursive: true });
  writeFileSync(path.join(statics, 'page', 'index.html'), '<!doctype html>');
  writeFileSync(path.join(env.base, 'secret.txt'), 'secret');
  symlinkSync(path.join(env.base, 'secret.txt'), path.join(statics, 'page', 'leak.txt'));
  const tabs = { 'tab-1': owner('tab-1'), 'tab-2': owner('tab-2') };
  const handler = createKitHttpHandler({ documents: env.docs, staticRoots: { page: path.join(statics, 'page') },
    admit: req => tabs[req.headers['x-tab']] ? { peer: tabs[req.headers['x-tab']] } : { rejection: 403 },
    allowedOrigin: origin => origin === 'http://app.local', meta: view => ({ mode: view.writer ? 'edit' : 'view' }) });
  const server = createServer(async (req, res) => { if (!(await handler(req, res))) { res.statusCode = 404; res.end(); } });
  await new Promise(resolve => server.listen(0, '127.0.0.1', resolve));
  t.after(() => new Promise(resolve => server.close(resolve)));
  const origin = `http://127.0.0.1:${server.address().port}`;
  const call = (route, { tab = 'tab-1', ...init } = {}) => fetch(`${origin}/shimo/kit/${route}`, { ...init, headers: { 'x-tab': tab, ...init.headers } });
  const view = await env.docs.open(env.scope, 'deck.pptx', tabs['tab-1']);
  return { ...env, call, view };
}

test('paths and ranges: traversal, encoded separators and hidden names are refused; one byte range at a time', () => {
  for (const bad of ['../x', 'a/../b', '%2e%2e/x', 'a%2fb', 'a\\b', '.env', 'a//b', '']) assert.throws(() => safeRelative(bad), /Invalid/, bad);
  assert.equal(safeRelative('ppt/ppt-sdk/code/sdk.js'), 'ppt/ppt-sdk/code/sdk.js');
  assert.equal(singleRange('bytes=0-99'), 'bytes=0-99');
  assert.equal(singleRange('bytes=-5'), 'bytes=-5');
  assert.equal(singleRange('bytes=0-1,5-9'), undefined, 'multi-part ranges are served whole');
});

test('content: first the opened content, then a rebuild drains the queue and restarts the change sequence', async t => {
  const { call, view } = await serve(t);
  const first = await call(`v/${view.id}/content`);
  assert.equal(await first.text(), '{"slides":[]}');
  for (const seq of [0, 1]) assert.equal((await call(`v/${view.id}/change`, { method: 'POST', body: JSON.stringify({ seq, format: 'mop-v1', body: `c${seq}` }) })).status, 200);
  const gap = await call(`v/${view.id}/change`, { method: 'POST', body: JSON.stringify({ seq: 7, format: 'mop-v1', body: 'late' }) });
  assert.deepEqual([gap.status, await gap.json()], [409, { error: 'edit-conflict', message: 'Change 7 arrived, 2 was expected.', rebuild: true }]);
  const rebuilt = await call(`v/${view.id}/content`);
  assert.deepEqual([await rebuilt.text(), rebuilt.headers.get('x-next-change-set-revision')], ['{"slides":[],"changes":2}', '2']);
  assert.equal((await call(`v/${view.id}/change`, { method: 'POST', body: JSON.stringify({ seq: 0, format: 'mop-v1', body: 'after rebuild' }) })).status, 200);
});

test('change: only a literal `inert: true` marks the editor\'s own change; anything else is an edit', async t => {
  const { call, view, docs } = await serve(t);
  assert.equal((await call(`v/${view.id}/change`, { method: 'POST', body: JSON.stringify({ seq: 0, format: 'mop-v1', body: 'layout', inert: true }) })).status, 200);
  assert.equal(docs.views.get(view.id).unsaved, false);
  assert.equal((await call(`v/${view.id}/change`, { method: 'POST', body: JSON.stringify({ seq: 1, format: 'mop-v1', body: 'typed', inert: 'yes' }) })).status, 200);
  assert.equal(docs.views.get(view.id).unsaved, true);
});

test('writes need the same origin (or the desktop shell\'s stripped forwarding); views belong to their tab', async t => {
  const { call, view } = await serve(t);
  const body = JSON.stringify({ seq: 0, format: 'mop-v1', body: 'x' });
  assert.equal((await call(`v/${view.id}/change`, { method: 'POST', body, headers: { origin: 'http://evil.example' } })).status, 403);
  assert.equal((await call(`v/${view.id}/change`, { method: 'POST', body, headers: { 'sec-fetch-site': 'cross-site' } })).status, 403, 'Sec-Fetch-Site without Origin is not the desktop shell');
  assert.equal((await call(`v/${view.id}/change`, { method: 'POST', body, headers: { origin: 'http://app.local' } })).status, 200);
  assert.equal((await call(`v/${view.id}/content`, { tab: 'tab-2' })).status, 404, 'another tab cannot read this view');
  assert.equal((await call(`v/${view.id}/content`, { tab: 'nobody' })).status, 403);
  const headers = (await call(`v/${view.id}/meta`)).headers;
  assert.match(headers.get('content-security-policy'), /frame-ancestors 'self'/);
  assert.equal(headers.get('x-content-type-options'), 'nosniff');
});

test('media: whole reads, a single range as 206, an impossible range as 416; static files stay inside their root', async t => {
  const { call, view } = await serve(t);
  const whole = await call(`v/${view.id}/media?reference=mop-asset:/media/x.png`);
  assert.deepEqual([whole.status, await whole.text()], [200, 'media mop-asset:/media/x.png']);
  const part = await call(`v/${view.id}/media?reference=v.mp4`, { headers: { range: 'bytes=2-4' } });
  assert.deepEqual([part.status, part.headers.get('content-range'), await part.text()], [206, 'bytes 2-4/10', 'dia']);
  assert.equal((await call(`v/${view.id}/media?reference=v.mp4`, { headers: { range: 'bytes=999-' } })).status, 416);
  const page = await call('static/page/index.html', { tab: 'nobody' });
  assert.deepEqual([page.status, page.headers.get('content-type')], [200, 'text/html; charset=utf-8']);
  assert.equal((await call('static/page/leak.txt')).status, 404, 'a link out of the root is not served');
  const escaped = await call('static/page/%2e%2e/secret.txt');
  assert.ok(escaped.status >= 400 && !(await escaped.text()).includes('secret'), 'a dotted path never reaches the file');
  assert.equal((await call('static/other/index.html')).status, 404);
});

test('downloads: one token, one tab, one read; the exported file is removed afterwards', async t => {
  const { call, view, docs } = await serve(t);
  const { token, name } = await docs.exportDownload(view.id, owner('tab-1'));
  const file = docs.downloads.get(token).file;
  assert.equal((await call(`v/${view.id}/download/${token}`, { tab: 'tab-2' })).status, 404, 'another tab cannot take it');
  assert.equal((await call(`v/${view.id}/download/${token}`, { method: 'HEAD' })).status, 404, 'a HEAD does not spend the token');
  assert.equal((await import('node:fs')).existsSync(file), true, 'a refused tab does not consume the token');
  const response = await call(`v/${view.id}/download/${token}`);
  assert.deepEqual([response.status, response.headers.get('content-disposition'), await response.text()], [200, `attachment; filename*=UTF-8''${name}`, 'pptx bytes']);
  const { existsSync } = await import('node:fs');
  for (let i = 0; i < 50 && existsSync(file); i++) await new Promise(resolve => setTimeout(resolve, 10));
  assert.equal(existsSync(file), false, 'the export is deleted once streamed');
  assert.equal((await call(`v/${view.id}/download/${token}`)).status, 404, 'a token works once');
});

test('the page CSP lets suite workers start from blob: and data: but nothing leaves the origin', async t => {
  const { call, view } = await serve(t);
  assert.equal((await call(`v/${view.id}/meta`)).headers.get('content-security-policy'), KIT_CSP);
  const directives = Object.fromEntries(KIT_CSP.split(';').map(part => part.trim().split(/\s+/)).map(([name, ...values]) => [name, values]));
  assert.deepEqual(directives['worker-src'], ["'self'", 'blob:', 'data:']);
  for (const values of Object.values(directives)) assert.ok(values.every(value => !/^https?:|^\*$/.test(value)), 'no remote sources');
  assert.deepEqual(directives['frame-ancestors'], ["'self'"]);
});

test('a sheet: media is posted without kind, and its download is an XLSX', async t => {
  const { call, docs, scope } = await serve(t);
  const book = await docs.open(scope, 'book.xlsx', owner('tab-1'));
  const posted = await call(`v/${book.id}/media?extension=png`, { method: 'POST', body: 'png', headers: { origin: 'http://app.local' } });
  assert.equal(posted.status, 200);
  const refused = await call(`v/${book.id}/media?extension=png&kind=image`, { method: 'POST', body: 'png', headers: { origin: 'http://app.local' } });
  assert.equal(refused.status, 400);
  docs.cancelMedia(book.id, owner('tab-1'), (await posted.json()).reference);
  const { token } = await docs.exportDownload(book.id, owner('tab-1'));
  const download = await call(`v/${book.id}/download/${token}`);
  assert.deepEqual([download.status, download.headers.get('content-type'), download.headers.get('content-disposition')],
    [200, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', "attachment; filename*=UTF-8''book.xlsx"]);
});

test('POST abandon closes the view for its own tab, same-origin only', async t => {
  const { call, view, docs } = await serve(t);
  assert.equal((await call(`v/${view.id}/abandon`, { method: 'POST', headers: { origin: 'http://evil.example' } })).status, 403);
  assert.equal((await call(`v/${view.id}/abandon`, { method: 'POST', tab: 'tab-2', headers: { origin: 'http://app.local' } })).status, 404);
  const abandoned = await call(`v/${view.id}/abandon`, { method: 'POST', headers: { origin: 'http://app.local' } });
  assert.deepEqual([abandoned.status, await abandoned.json()], [200, { closed: true, unsaved: false }]);
  assert.equal(docs.views.has(view.id), false);
});
