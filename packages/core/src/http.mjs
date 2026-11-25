import { SuiteError, fail } from './errors.mjs';
import { requestPath, safeRelative } from './paths.mjs';
import { RuntimeReaders } from './runtime-readers.mjs';

export const SUITE_CSP = "default-src 'self'; script-src 'self' 'unsafe-inline' 'unsafe-eval' blob:; style-src 'self' 'unsafe-inline'; connect-src 'self' blob:; img-src 'self' data: blob:; font-src 'self' data:; media-src 'self' blob:; worker-src 'self' blob:; frame-src 'self' blob:; object-src 'none'; base-uri 'self'; form-action 'none'; frame-ancestors 'self'";
const json = value => ({ body: Buffer.from(JSON.stringify(value)), type: 'application/json; charset=utf-8' });
const STAGE_TYPES = ['application/zip', 'application/octet-stream', 'application/json', 'text/plain'];

/** Fixed messages per status keep paths and backend details out of every response. */
const PUBLIC_MESSAGE = { 400: 'Invalid request.', 403: 'Access denied.', 404: 'Not found.', 405: 'Method not allowed.', 409: 'The document changed.', 413: 'Too large.', 422: 'The document could not be processed.', 429: 'Busy, try again.', 499: 'Cancelled.', 503: 'Unavailable.', 504: 'Timed out.' };

async function readBody(req, cap, signal) {
  const declared = Number(req.headers['content-length']);
  if (Number.isFinite(declared) && declared > cap) fail('input-too-large', 'The request body is too large.', { limit: 'maxStageBytes' });
  const chunks = [];
  let total = 0;
  for await (const chunk of req) {
    signal.throwIfAborted();
    total += chunk.length;
    if (total > cap) fail('input-too-large', 'The request body is too large.', { limit: 'maxStageBytes' });
    chunks.push(chunk);
  }
  return Buffer.concat(chunks, total);
}

/** Explicit read compatibility surface for the spreadsheet SDK: no wildcard fallbacks. */
export async function sheetCompatibility(service, id, route, peer, signal) {
  const meta = JSON.parse((await service.resource(id, 'meta', peer, signal)).body);
  if (meta.format !== 'xlsx') fail('not-found', 'Route not found.');
  const parts = route.split('/');
  if (parts[0] === 'files' && parts[1] === id) {
    const action = parts.slice(2).join('/');
    if (action === 'content') return service.resource(id, 'content', peer, signal);
    if (action === 'snapshot') return json((await service.resource(id, 'content', peer, signal)).body.toString('utf8'));
    if (action === 'head') return json({ head: 1 });
    if (action === 'pull') return { body: Buffer.from('event: changes\ndata: []\n\n'), type: 'text/event-stream; charset=utf-8' };
    if (action === 'csHistories' || action === 'cellCsHistories' || action === 'comments') return json([]);
    if (action === 'collabInfo') return json({ editRecommended: false, fileSize: 0, bigFile: false, usersCount: 0 });
    if (action === 'demo/info') return json({ errorCode: 30705, message: 'No demo information.', errorDetail: {}, traceId: '' });
    if (action.startsWith('resources/')) return service.resource(id, action, peer, signal);
  }
  if (route === 'globals/checkpoint') return json({ data: { features: ['support_mosheet'], quotas: {} } });
  if (route === 'uploader/settings') return json({ fileTypeBlockList: { extensions: {} }, urlReplacerPatterns: [] });
  if (/^user\/users\/[^/]+\/info$/.test(route)) return json({ domain: 'user', requestId: '', data: { avatar: '', names: { name: 'Preview', alias: '' }, labels: [], infos: [{ type: 'builtin', key: 'id', value: parts[2] }], paths: [] } });
  if (/^user\/users\/[^/]+\/watermark$/.test(route)) return json({ data: { watermarks: [], downloadWatermarks: [] }, scope: 'user', requestId: '' });
  if (/^sheetcalc\/s\/[^/]+\/meta$/.test(route)) return json({ cells: 0, formulas: 0, refCells: 0, sheetList: {} });
  if (/^sheetcalc\/s\/[^/]+\/cell-cache$/.test(route)) return json({ code: 100, message: 'normal' });
  fail('not-found', 'SDK route not found.');
}

/**
 * Node HTTP handler for one route prefix. Hosts supply admission (who is asking) and the
 * set of origins allowed to POST; Core owns routing, limits, headers and error shape.
 */
export function createHttpHandler({ service, admit, allowedOrigin }) {
  const prefix = service.routePrefix;
  const runtimeReaders = new RuntimeReaders(service.limits.maxReaders);
  async function post(req, id, route, peer, signal) {
    const origin = req.headers.origin;
    // Browsers send Origin and Sec-Fetch-Site on every POST. The DSH desktop shell checks the
    // Origin is its own app and strips both before forwarding, so only that carrier sends neither.
    const forwarded = origin === undefined && req.headers['sec-fetch-site'] === undefined;
    if (!forwarded && (!origin || !allowedOrigin(origin, req))) fail('forbidden', 'Cross-origin writes are not allowed.');
    const session = await service.session(id, peer, signal);
    const url = new URL(req.url, 'http://local');
    if (route.startsWith('stage/')) {
      const type = String(req.headers['content-type'] ?? '').split(';')[0].trim().toLowerCase();
      if (!STAGE_TYPES.includes(type)) fail('invalid-body', 'Unsupported document part type.');
      const baseEditRevision = Number(url.searchParams.get('base'));
      const body = await readBody(req, service.limits.maxStageBytes, signal);
      return json(await service.stage(id, safeRelative(route.slice('stage/'.length)), body, { baseEditRevision }, peer, signal));
    }
    const input = JSON.parse((await readBody(req, 64 * 1024, signal)).toString('utf8') || '{}');
    if (route === 'save') return json(await service.save(session.scope, id, { editRevision: input.editRevision }, signal, peer));
    if (route === 'export') return json(await service.exportTo(session.scope, id, { editRevision: input.editRevision, target: input.target }, signal, peer));
    fail('not-found', 'Route not found.');
  }
  return async function handle(req, res) {
    const cancel = new AbortController();
    const abort = () => cancel.abort(new SuiteError('cancelled', 'The request disconnected.'));
    req.once('aborted', abort);
    res.once('close', () => { if (!res.writableFinished) abort(); });
    const signal = AbortSignal.any([cancel.signal, service.lifetime.signal]);
    try {
      const admission = await admit(req);
      if (!admission || 'rejection' in admission) fail('forbidden', 'Access denied.');
      const { peer } = admission;
      const path = requestPath(req.url, prefix);
      const method = req.method;
      const reading = method === 'GET' || method === 'HEAD';
      let result;
      if (path.startsWith('runtime/')) {
        if (!reading) { res.setHeader('Allow', 'GET, HEAD'); fail('read-only', 'Runtime resources are read only.'); }
        const relative = path.slice('runtime/'.length).replace(/^font-metrics-mow\/fontmetrics\//, 'fontmetrics/');
        result = await runtimeReaders.read(signal, () => service.runtimeResource(relative, signal));
      } else {
        const match = /^(doc|draft)\/([a-f0-9]{32,48})\/(.+)$/.exec(path);
        if (!match) fail('not-found', 'Route not found.');
        const [, kind, id, route] = match;
        if (kind === 'draft') {
          if (!reading) { res.setHeader('Allow', 'GET, HEAD'); fail('read-only', 'Draft downloads are read only.'); }
          result = await service.draftResource(id, route, peer, signal);
        } else if (method === 'POST') result = await post(req, id, route, peer, signal);
        else if (!reading) { res.setHeader('Allow', 'GET, HEAD, POST'); fail('read-only', 'Unsupported method.'); }
        else if (route.startsWith('sdk/v2/api/')) result = await sheetCompatibility(service, id, route.slice('sdk/v2/api/'.length), peer, signal);
        else result = await service.resource(id, route, peer, signal);
      }
      res.statusCode = 200;
      res.setHeader('Content-Type', result.type);
      res.setHeader('Content-Length', result.body.length);
      setSecurityHeaders(res);
      for (const [name, value] of Object.entries(result.headers ?? {})) res.setHeader(name, value);
      res.end(method === 'HEAD' ? undefined : result.body);
    } catch (error) {
      if (res.destroyed || res.headersSent) return;
      const known = error instanceof SuiteError;
      const status = known ? error.status : error instanceof SyntaxError ? 400 : 500;
      const code = known ? error.code : status === 400 ? 'invalid-body' : 'unavailable';
      res.statusCode = status;
      res.setHeader('Content-Type', 'application/json; charset=utf-8');
      setSecurityHeaders(res);
      const details = known && error.details?.draftId ? { draftId: error.details.draftId } : undefined;
      res.end(req.method === 'HEAD' ? undefined : JSON.stringify({ error: code, message: PUBLIC_MESSAGE[status] ?? 'Unavailable.', ...(details && { details }) }));
    } finally {
      req.removeListener('aborted', abort);
    }
  };
}

export function setSecurityHeaders(res) {
  // Runtime URLs survive package replacement and have no verified content digest.
  res.setHeader('Cache-Control', 'no-store');
  res.setHeader('X-Content-Type-Options', 'nosniff');
  res.setHeader('Referrer-Policy', 'no-referrer');
  res.setHeader('Content-Security-Policy', SUITE_CSP);
  res.setHeader('Cross-Origin-Resource-Policy', 'same-origin');
}
