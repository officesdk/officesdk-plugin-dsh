/**
 * Where a download comes from (spec 261110-universal-package §4.4). A source is never trusted: whatever it returns is
 * checked against the manifest's size and hashes by the downloader. Each opener resolves to
 * `{ body, total?, etag?, resumed, resumable, format }`: `body` an async iterable of Buffers, `format` `tgz` (an archive
 * whose sha512 the manifest pins) or `tar` (a plain tar stream, checked file by file).
 *
 *   url             `base + file` over HTTP(S): office-kit's release directory (its GitHub release) or a mirror with the same layout
 *   npm             `{ registry, package, version }`: the packument's dist.tarball (reserved for office-kit's npm)
 * DSH installs its proxy policy as undici's global dispatcher, so the global fetch here goes through the system proxy.
 */
export class SourceError extends Error {
  constructor(message, { code = 'source-failed', status, retry = true } = {}) { super(message); this.name = 'SourceError'; this.code = code; this.status = status; this.retry = retry; }
}

const CONNECT_TIMEOUT_MS = 30_000;
const withTimeout = (signal, ms) => (signal ? AbortSignal.any([signal, AbortSignal.timeout(ms)]) : AbortSignal.timeout(ms));
// Every `//user:password@` in a URL or in an error text that quotes one (fetch names the URL it refused).
const redactUrl = url => String(url).replace(/(\/\/)[^/@\s]+@/g, '$1');

/** Body of a fetch Response as an async iterable of Buffers. */
async function* bodyOf(response) {
  for await (const chunk of response.body) yield Buffer.from(chunk);
}

async function openUrl(url, { offset = 0, etag, signal, headers = {}, fetchImpl = fetch, format = 'tgz', connectTimeoutMs = CONNECT_TIMEOUT_MS }) {
  const request = { ...headers };
  if (offset > 0 && etag) Object.assign(request, { Range: `bytes=${offset}-`, 'If-Range': etag });
  let response;
  // The connect timeout covers the headers only; the downloader times out a body that stops sending.
  const connect = new AbortController();
  const timer = setTimeout(() => connect.abort(new SourceError(`no answer from ${redactUrl(url)} in ${connectTimeoutMs / 1000} s`, { code: 'timeout' })), connectTimeoutMs);
  try { response = await fetchImpl(url, { headers: request, signal: signal ? AbortSignal.any([signal, connect.signal]) : connect.signal, redirect: 'follow' }); }
  catch (error) {
    if (signal?.aborted) throw signal.reason;
    throw connect.signal.aborted ? connect.signal.reason : new SourceError(redactUrl(`${url}: ${error?.cause?.code ?? error?.message ?? error}`), { code: 'network' });
  } finally { clearTimeout(timer); }
  if (response.status === 401 || response.status === 403 || response.status === 404) {
    throw new SourceError(`${redactUrl(url)}: HTTP ${response.status}`, { status: response.status, code: response.status === 404 ? 'not-found' : 'denied', retry: false });
  }
  if (!response.ok) throw new SourceError(`${redactUrl(url)}: HTTP ${response.status}`, { status: response.status });
  const resumed = response.status === 206;
  const length = Number(response.headers.get('content-length'));
  const total = Number.isFinite(length) && length > 0 ? (resumed ? offset + length : length) : undefined;
  return { body: bodyOf(response), total, etag: response.headers.get('etag') ?? undefined, resumed, resumable: response.headers.get('accept-ranges') === 'bytes' || resumed, format };
}

/**
 * Opens one item from one source. `context`: `{ offset, etag, signal, fetchImpl, spawn }`. A source may set its own
 * `connectTimeoutMs` (answer and headers; default 30 s): the npm copy of the editor uses a short one so a registry that
 * cannot be reached gives way to the next source quickly.
 */
export async function openSource(source, item, context = {}) {
  item = { ...item, file: item.file ?? item.archive?.file, integrity: item.integrity ?? item.archive?.integrity };
  if (source.connectTimeoutMs) context = { ...context, connectTimeoutMs: source.connectTimeoutMs };
  switch (source.kind) {
    case 'url': {
      if (!item.file) throw new SourceError('a url source needs a file name', { retry: false });
      return openUrl(`${source.base.replace(/\/?$/, '/')}${encodeURI(item.file)}`, context);
    }
    case 'npm': {
      const fetchImpl = context.fetchImpl ?? fetch;
      const registry = source.registry.replace(/\/?$/, '/');
      const headers = source.token ? { Authorization: `Bearer ${source.token}` } : {};
      const limit = context.connectTimeoutMs ?? CONNECT_TIMEOUT_MS;
      let response, dist;
      // The packument is small: its whole body counts against the same limit as its headers.
      const packument = withTimeout(context.signal, limit);
      try {
        response = await fetchImpl(`${registry}${source.package.replace('/', '%2F')}`, { headers: { ...headers, Accept: 'application/json' }, signal: packument });
        if (response.ok) dist = (await response.json())?.versions?.[source.version]?.dist;
      } catch (error) {
        if (context.signal?.aborted) throw context.signal.reason;
        const timedOut = packument.aborted;
        throw new SourceError(timedOut ? `no answer from ${redactUrl(registry)} in ${limit / 1000} s` : redactUrl(`${registry}: ${error?.cause?.code ?? error?.message ?? error}`), { code: timedOut ? 'timeout' : 'network' });
      }
      if (!response.ok) throw new SourceError(`${redactUrl(registry)}${source.package}: HTTP ${response.status}`, { status: response.status, retry: response.status >= 500 });
      if (!dist?.tarball) throw new SourceError(`${source.package}@${source.version} is not in ${redactUrl(registry)}`, { code: 'not-found', retry: false });
      // A registry that states an integrity must state the manifest's (GitLab's states only sha1, which is not compared).
      if (dist.integrity?.startsWith('sha512-') && item.integrity && dist.integrity !== item.integrity) throw new SourceError(`${source.package}@${source.version}: registry integrity differs from the manifest`, { code: 'integrity', retry: false });
      return openUrl(dist.tarball, { ...context, headers });
    }
    default:
      throw new SourceError(`unknown source kind ${source.kind}`, { retry: false });
  }
}
