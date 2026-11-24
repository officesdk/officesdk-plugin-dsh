import { createReadStream } from 'node:fs';
import { realpath, rm, stat } from 'node:fs/promises';
import * as nodePath from 'node:path';
import { pipeline } from 'node:stream/promises';
import { KitError } from './documents.mjs';

/**
 * Page routes for Kit documents (spec §12.3–12.4), under `<prefix>/kit/`:
 *   GET  v/<view>/meta | v/<view>/content          the page's mount data and Kit content (string, unchanged)
 *   POST v/<view>/change                           { seq, format, body, inert? }: one numbered suite change; `inert`: the
 *                                                  editor made it on its own before anyone acted (it edits nothing)
 *   GET  v/<view>/media/<reference>?purpose=       media bytes (single Range → 206); also ?reference=
 *   POST v/<view>/media?extension=&kind=           inserted bytes → { reference }; DELETE ?reference= cancels
 *   GET  v/<view>/download/<token>                 a one-time export download
 *   POST v/<view>/abandon                          the page could not mount an editor: close without autosave
 *   GET  fonts/catalog | fonts/file?resourceId=    system fonts through Kit
 *   GET  static/<root>/<file>                      editor bundle, SDK resources and the page itself
 * `admit(req)` returns `{ peer }` or `{ rejection }`; writes need a same-origin Origin (or the desktop shell's
 * forwarding, which strips both Origin and Sec-Fetch-Site).
 */
export const KIT_CSP = "default-src 'self'; script-src 'self' 'unsafe-inline' 'unsafe-eval' 'wasm-unsafe-eval' blob:; style-src 'self' 'unsafe-inline'; connect-src 'self' blob: data:; img-src 'self' data: blob:; font-src 'self' data: blob:; media-src 'self' blob:; worker-src 'self' blob: data:; frame-src 'self' blob:; object-src 'none'; base-uri 'self'; form-action 'none'; frame-ancestors 'self'";
const STATUS = { forbidden: 403, 'sandbox-denied': 403, 'invalid-path': 400, 'invalid-body': 400, 'unsupported-format': 400, 'password-protected': 422, 'invalid-format': 422,
  busy: 429, 'source-changed': 409, 'edit-conflict': 409, exists: 409, 'lease-unavailable': 404, 'not-found': 404, 'read-only': 405, unavailable: 503, 'input-too-large': 413 };
const TYPES = { '.html': 'text/html; charset=utf-8', '.js': 'text/javascript; charset=utf-8', '.mjs': 'text/javascript; charset=utf-8', '.css': 'text/css; charset=utf-8',
  '.json': 'application/json; charset=utf-8', '.wasm': 'application/wasm', '.png': 'image/png', '.jpg': 'image/jpeg', '.jpeg': 'image/jpeg', '.gif': 'image/gif', '.webp': 'image/webp',
  '.svg': 'image/svg+xml', '.woff': 'font/woff', '.woff2': 'font/woff2', '.ttf': 'font/ttf', '.otf': 'font/otf', '.map': 'application/json; charset=utf-8' };
const MAX_CHANGE = 64 * 1024 * 1024;
const MAX_MEDIA = 256 * 1024 * 1024;
const VIEW = /^[0-9a-f-]{36}$/;

export function setKitHeaders(res) {
  res.setHeader('Cache-Control', 'no-store');
  res.setHeader('X-Content-Type-Options', 'nosniff');
  res.setHeader('Referrer-Policy', 'no-referrer');
  res.setHeader('Content-Security-Policy', KIT_CSP);
  res.setHeader('Cross-Origin-Resource-Policy', 'same-origin');
}

/** A relative path inside one directory: no traversal, no encoded separators, no NUL, no hidden names. */
export function safeRelative(encoded) {
  if (/%2f|%5c|%00/i.test(encoded)) throw new KitError('invalid-path', 'Invalid resource path.');
  let decoded;
  try { decoded = decodeURIComponent(encoded); } catch { throw new KitError('invalid-path', 'Invalid resource path.'); }
  const parts = decoded.split('/');
  if (!decoded || decoded.includes('\\') || decoded.includes('\0') || parts.some(part => !part || part === '.' || part === '..' || part.startsWith('.'))) throw new KitError('invalid-path', 'Invalid resource path.');
  return parts.join('/');
}

/** One single-part byte range for Kit; anything else is served whole (kit-usage/03-calls.md). */
export function singleRange(header) {
  if (typeof header !== 'string') return undefined;
  const match = /^bytes=(\d*)-(\d*)$/.exec(header.trim());
  if (!match || (!match[1] && !match[2])) return undefined;
  return `bytes=${match[1]}-${match[2]}`;
}

async function readBody(req, cap) {
  const chunks = [];
  let size = 0;
  for await (const chunk of req) {
    size += chunk.length;
    if (size > cap) throw new KitError('input-too-large', 'The request body is too large.');
    chunks.push(chunk);
  }
  return Buffer.concat(chunks);
}

export function createKitHttpHandler({ documents, admit, allowedOrigin, staticRoots, meta = view => view, prefix = '/shimo', reviews }) {
  const base = `${prefix}/kit/`;
  const sameOrigin = req => {
    const origin = req.headers.origin;
    if (origin === undefined && req.headers['sec-fetch-site'] === undefined) return true;
    return typeof origin === 'string' && allowedOrigin(origin, req);
  };
  const json = (res, status, value) => { res.statusCode = status; res.setHeader('Content-Type', 'application/json; charset=utf-8'); res.end(JSON.stringify(value)); };

  async function serveStatic(res, root, relative, method) {
    const dir = staticRoots[root];
    if (!dir) throw new KitError('not-found', 'Not found.');
    const realRoot = await realpath(dir);
    const file = await realpath(nodePath.join(realRoot, safeRelative(relative))).catch(() => { throw new KitError('not-found', 'Not found.'); });
    if (!file.startsWith(realRoot + nodePath.sep)) throw new KitError('not-found', 'Not found.');
    const info = await stat(file);
    if (!info.isFile()) throw new KitError('not-found', 'Not found.');
    res.statusCode = 200;
    res.setHeader('Content-Type', TYPES[nodePath.extname(file).toLowerCase()] ?? 'application/octet-stream');
    res.setHeader('Content-Length', info.size);
    if (method === 'HEAD') return res.end();
    await pipeline(createReadStream(file), res);
  }

  async function streamBody(res, result, range) {
    const partial = result.totalLength !== undefined && range;
    res.statusCode = partial ? 206 : 200;
    res.setHeader('Content-Type', result.contentType ?? 'application/octet-stream');
    res.setHeader('Content-Length', result.length);
    res.setHeader('Accept-Ranges', 'bytes');
    if (partial) res.setHeader('Content-Range', `bytes ${result.offset}-${result.offset + result.length - 1}/${result.totalLength}`);
    await pipeline(result.body, res);
  }

  return async function handle(req, res) {
    const url = new URL(req.url, 'http://local');
    if (!url.pathname.startsWith(base)) return false;
    setKitHeaders(res);
    const parts = url.pathname.slice(base.length).split('/');
    const method = req.method ?? 'GET';
    try {
      if (parts[0] === 'static') {
        if (method !== 'GET' && method !== 'HEAD') throw new KitError('read-only', 'Static resources are read only.');
        await serveStatic(res, parts[1], parts.slice(2).join('/'), method);
        return true;
      }
      const admission = admit(req);
      if ('rejection' in admission) { res.statusCode = admission.rejection; res.end(); return true; }
      const { peer } = admission;
      const writing = method !== 'GET' && method !== 'HEAD';
      if (writing && !sameOrigin(req)) throw new KitError('forbidden', 'Cross-origin writes are not allowed.');
      if (parts[0] === 'fonts') {
        if (parts[1] === 'catalog' && !writing) return json(res, 200, await documents.fontCatalog()), true;
        if (parts[1] === 'file' && !writing) return await streamBody(res, await documents.readFont(url.searchParams.get('resourceId'))), true;
        throw new KitError('not-found', 'Not found.');
      }
      if (parts[0] !== 'v' || !VIEW.test(parts[1] ?? '')) throw new KitError('not-found', 'Not found.');
      const [, viewId, route] = parts;
      if (route === 'meta' && !writing) {
        const view=documents.view(viewId,peer), info=meta(view);
        return json(res,200,reviews && view.writer ? {...info,reviewHistory:true} : info),true;
      }
      if (route==='reviews' && reviews && (method==='GET' || method==='POST')) {
        const view=documents.view(viewId,peer);
        if (!view.writer) throw new KitError('read-only','This view has no editable review history.');
        const result=method==='GET' ? await reviews.read(view) :
          await reviews.commit(view,JSON.parse((await readBody(req,MAX_CHANGE)).toString('utf8')));
        return json(res,200,result),true;
      }
      if (route === 'content' && !writing) {
        const content = await documents.content(viewId, peer);
        res.setHeader('Content-Type', 'application/json; charset=utf-8');
        res.setHeader('X-Next-Change-Set-Revision', String(content.nextChangeSetRevision ?? ''));
        res.end(content.content);
        return true;
      }
      if (route === 'change' && method === 'POST') {
        const change = JSON.parse((await readBody(req, MAX_CHANGE)).toString('utf8'));
        return json(res, 200, await documents.change(viewId, peer, { seq: change?.seq, format: change?.format, body: change?.body, inert: change?.inert === true })), true;
      }
      if (route === 'abandon' && method === 'POST') return json(res, 200, documents.abandon(viewId, peer)), true;
      // GET only: a HEAD must not spend the one-time token.
      if (route === 'download' && method === 'GET') {
        const download = documents.takeDownload(viewId, peer, parts[3]);
        try {
          const info = await stat(download.file);
          res.statusCode = 200;
          res.setHeader('Content-Type', download.contentType);
          res.setHeader('Content-Length', info.size);
          res.setHeader('Content-Disposition', `attachment; filename*=UTF-8''${encodeURIComponent(download.name)}`);
          await pipeline(createReadStream(download.file), res);
        } finally { await rm(download.file, { force: true }); }
        return true;
      }
      if (route === 'media') {
        if (method === 'GET' || method === 'HEAD') {
          const range = singleRange(req.headers.range);
          // `media/<encoded reference>` (what the page's asset proxy requests) or `media?reference=`.
          const reference = parts[3] ? decodeURIComponent(parts[3]) : url.searchParams.get('reference');
          return await streamBody(res, await documents.readMedia(viewId, peer, { reference, purpose: url.searchParams.get('purpose') ?? 'display', range }), range), true;
        }
        if (method === 'POST') return json(res, 200, await documents.writeMedia(viewId, peer, { bytes: await readBody(req, MAX_MEDIA), extension: url.searchParams.get('extension'), kind: url.searchParams.get('kind') ?? undefined })), true;
        if (method === 'DELETE') return json(res, 200, documents.cancelMedia(viewId, peer, url.searchParams.get('reference'))), true;
      }
      throw new KitError('not-found', 'Not found.');
    } catch (error) {
      if (res.headersSent) { res.destroy(error); return true; }
      if (error instanceof SyntaxError) error = new KitError('invalid-body', 'The request body is not JSON.');
      const code = error instanceof KitError ? error.code : 'unavailable';
      if (code === 'range-unsatisfiable') { res.statusCode = 416; res.end(); return true; }
      const { rebuild, overwrite, outside } = error.details ?? {};
      json(res, STATUS[code] ?? 500, { error: code, message: error instanceof KitError ? error.message : 'The document service is unavailable.', ...(rebuild && { rebuild }), ...(overwrite && { overwrite }), ...(outside && { outside }) });
    }
    return true;
  };
}
