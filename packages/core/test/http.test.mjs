import test from 'node:test';
import assert from 'node:assert/strict';
import { createServer } from 'node:http';
import { mkdir, writeFile } from 'node:fs/promises';
import { join } from 'node:path';
import { SUITE_CSP, createHttpHandler } from '../src/index.mjs';
import { fixture, peerOf, scope, signal, zip } from './helpers.mjs';

async function serve(t, options) {
  const f = await fixture(t, options);
  const peer = peerOf();
  await mkdir(join(f.service.runtimeRoot, 'fontmetrics'));
  await writeFile(join(f.service.runtimeRoot, 'fontmetrics', 'a.json'), '{}');
  let admitted = true;
  const handler = createHttpHandler({ service: f.service, admit: () => admitted ? { peer } : { rejection: 403 }, allowedOrigin: origin => origin === 'http://app.local' });
  const server = createServer(handler);
  await new Promise(resolve => server.listen(0, '127.0.0.1', resolve));
  t.after(() => new Promise(resolve => server.close(resolve)));
  const base = `http://127.0.0.1:${server.address().port}`;
  const request = (path, init = {}) => fetch(base + path, init);
  return { ...f, peer, request, deny: () => { admitted = false; } };
}

const expectHeaders = response => {
  assert.equal(response.headers.get('cache-control'), 'no-store');
  assert.equal(response.headers.get('x-content-type-options'), 'nosniff');
  assert.equal(response.headers.get('referrer-policy'), 'no-referrer');
  assert.equal(response.headers.get('cross-origin-resource-policy'), 'same-origin');
  assert.equal(response.headers.get('content-security-policy'), SUITE_CSP);
};

test('read routes carry fixed security headers; errors expose only a code and fixed text', async t => {
  const s = await serve(t);
  const handle = await s.service.open(scope, 'a.docx', 1, signal(), s.peer);
  const ok = await s.request(`/shimo/doc/${handle.id}/content`);
  assert.equal(ok.status, 200);
  expectHeaders(ok);
  assert.equal(ok.headers.get('x-amz-meta-head'), '1');
  assert.equal(await ok.text(), '{"doc":1}');
  const runtime = await s.request('/shimo/runtime/font-metrics-mow/fontmetrics/a.json');
  assert.equal(runtime.status, 200);
  const missing = await s.request(`/shimo/doc/${handle.id}/resources/images/none.png`);
  assert.equal(missing.status, 404);
  expectHeaders(missing);
  const body = await missing.json();
  assert.deepEqual(body, { error: 'not-found', message: 'Not found.' });
  assert.doesNotMatch(JSON.stringify(body), /authorized|\/tmp|\/var/);
  s.deny();
  assert.equal((await s.request(`/shimo/doc/${handle.id}/content`)).status, 403);
});

test('rejects traversal, encoded separators and leftover escapes before routing', async t => {
  const s = await serve(t);
  const handle = await s.service.open(scope, 'a.docx', 1, signal(), s.peer);
  for (const path of [`/shimo/doc/${handle.id}/resources/%2e%2e/x`, `/shimo/doc/${handle.id}/resources%2fx`, `/shimo/doc/${handle.id}/resources/a%5cb`,
    `/shimo/doc/${handle.id}/resources/a%00`, `/shimo/doc/${handle.id}/resources/%2541`, '/shimo/runtime//x', '/shimo/doc/zz/content', '/elsewhere']) {
    const response = await s.request(path);
    assert.ok([400, 404].includes(response.status), `${path} -> ${response.status}`);
  }
});

test('writes require POST from an allowed origin on an edit session; reads stay read-only', async t => {
  const s = await serve(t);
  const view = await s.service.open(scope, 'a.xlsx', 1, signal(), s.peer);
  const edit = await s.service.open(scope, 'a.xlsx', 1, signal(), s.peer, { mode: 'edit' });
  const post = (path, init = {}) => s.request(path, { method: 'POST', ...init, headers: { 'content-type': 'text/plain', ...init.headers } });
  assert.equal((await post(`/shimo/doc/${edit.id}/stage/content?base=0`, { body: 'x', headers: { 'sec-fetch-site': 'same-origin' } })).status, 403, 'a browser request without Origin');
  assert.equal((await post(`/shimo/doc/${edit.id}/stage/content?base=0`, { body: 'x', headers: { 'content-type': 'text/html' } })).status, 400, 'the desktop forwarder strips Origin and Sec-Fetch-Site: past the origin fence, stopped by the content type');
  assert.equal((await post(`/shimo/doc/${edit.id}/stage/content?base=0`, { body: 'x', headers: { origin: 'http://evil' } })).status, 403);
  const origin = { origin: 'http://app.local' };
  assert.equal((await post(`/shimo/doc/${view.id}/stage/content?base=0`, { body: 'x', headers: origin })).status, 405);
  assert.equal((await post(`/shimo/doc/${edit.id}/stage/content?base=0`, { body: 'x', headers: { ...origin, 'content-type': 'text/html' } })).status, 400);
  const staged = await post(`/shimo/doc/${edit.id}/stage/content?base=0`, { body: 'modoc-9', headers: origin });
  assert.equal(staged.status, 200);
  assert.deepEqual(await staged.json(), { editRevision: 1 });
  assert.equal((await post(`/shimo/doc/${edit.id}/stage/content?base=0`, { body: 'y', headers: origin })).status, 409);
  const saved = await post(`/shimo/doc/${edit.id}/save`, { body: JSON.stringify({ editRevision: 1 }), headers: { ...origin, 'content-type': 'application/json' } });
  assert.equal(saved.status, 200);
  assert.equal((await saved.json()).version, 'v2');
  const exported = await post(`/shimo/doc/${edit.id}/export`, { body: JSON.stringify({ editRevision: 1, target: 'download' }), headers: { ...origin, 'content-type': 'application/json' } });
  const { url } = await exported.json();
  const file = await s.request(url);
  assert.equal(file.status, 200);
  assert.match(file.headers.get('content-disposition'), /attachment; filename\*=UTF-8''a\.xlsx/);
  for (const method of ['PUT', 'DELETE', 'PATCH']) assert.equal((await s.request(`/shimo/doc/${edit.id}/content`, { method })).status, 405);
  assert.equal((await s.request('/shimo/runtime/fontmetrics/a.json', { method: 'POST', headers: origin })).status, 405);
});

test('stage bodies over the limit are refused while streaming', async t => {
  const s = await serve(t, { limits: { maxStageBytes: 16 } });
  const edit = await s.service.open(scope, 'a.docx', 1, signal(), s.peer, { mode: 'edit' });
  const response = await s.request(`/shimo/doc/${edit.id}/stage/package?base=0`, { method: 'POST', body: zip({ 'content.json': '{"x":"' + 'a'.repeat(64) + '"}' }), headers: { origin: 'http://app.local', 'content-type': 'application/zip' } });
  assert.equal(response.status, 413);
});

test('spreadsheet SDK compatibility is an explicit allowlist bound to xlsx sessions', async t => {
  const s = await serve(t);
  const sheet = await s.service.open(scope, 'a.xlsx', 1, signal(), s.peer);
  const doc = await s.service.open(scope, 'a.docx', 1, signal(), s.peer);
  const api = `/shimo/doc/${sheet.id}/sdk/v2/api`;
  assert.equal(await (await s.request(`${api}/files/${sheet.id}/content`)).text(), 'modoc-1');
  assert.equal(await (await s.request(`${api}/files/${sheet.id}/snapshot`)).json(), 'modoc-1');
  assert.deepEqual(await (await s.request(`${api}/files/${sheet.id}/head`)).json(), { head: 1 });
  assert.equal(await (await s.request(`${api}/files/${sheet.id}/resources/media/a.png`)).text(), 'png');
  assert.deepEqual(await (await s.request(`${api}/globals/checkpoint`)).json(), { data: { features: ['support_mosheet'], quotas: {} } });
  assert.equal((await s.request(`${api}/files/${sheet.id}/anything`)).status, 404);
  assert.equal((await s.request(`${api}/upload/token`)).status, 404);
  assert.equal((await s.request(`/shimo/doc/${doc.id}/sdk/v2/api/globals/checkpoint`)).status, 404);
});
