/**
 * Sheet media for a Kit page (spec §13.3). The Sheet SDK and its image worker request the raw
 * `modoc-assets:` reference of a picture instead of the URL the editor resolves, and the worker starts from a
 * Blob or a `data:` URL (an opaque origin that cannot call the Host). The page therefore answers those
 * requests itself: `fetch` on the page and a message bridge in each image worker read the bytes through this
 * view's `media/` route. Nothing here talks to Kit. Modelled on office-kit's reference Host (demo-shared/sheet-assets.ts).
 */
const SCHEME = 'modoc-assets:';
const SHIM_MARK = 'officesdk-sheet-media-shim';

/** `modoc-assets:/media/<name>` (or `modoc-assets://media/<name>`) → Kit's reference form, else undefined. */
export function sheetMediaReference(url) {
  if (typeof url !== 'string' || !url.startsWith(SCHEME)) return undefined;
  const rest = url.slice(SCHEME.length).replace(/^\/+/, '').split(/[?#]/, 1)[0];
  if (!rest.startsWith('media/')) return undefined;
  let name;
  try { name = decodeURIComponent(rest.slice('media/'.length)); } catch { return undefined; }
  return /^[A-Za-z0-9._-]+$/.test(name) && !name.startsWith('.') ? `${SCHEME}/media/${name}` : undefined;
}


const WORKER_SHIM = `/* ${SHIM_MARK} */
(() => {
  const pending = new Map();
  let next = 1;
  const nativeAdd = self.addEventListener.bind(self);
  nativeAdd('message', event => {
    const data = event.data;
    if (!data || data.kind !== 'kit-result') return;
    const settle = pending.get(data.id);
    if (settle) { pending.delete(data.id); settle(data); }
  });
  try {
    self.addEventListener = function (type, listener, options) {
      if (type !== 'message' || typeof listener !== 'function') return nativeAdd(type, listener, options);
      return nativeAdd(type, function (event) {
        const data = event && event.data;
        if (data && (data.kind === 'kit-result' || data.kind === 'kit-fetch')) return;
        return listener.call(this, event);
      }, options);
    };
  } catch (error) { void error; }
  const nativeFetch = fetch.bind(self);
  self.fetch = function (input, init) {
    const url = typeof input === 'string' ? input : input instanceof URL ? input.toString() : (input && input.url) || '';
    if (!url.startsWith('${SCHEME}')) return nativeFetch(input, init);
    const id = String(next++);
    return new Promise((resolve, reject) => {
      const timer = setTimeout(() => { if (pending.delete(id)) reject(new Error('media fetch timed out')); }, 30000);
      pending.set(id, data => {
        clearTimeout(timer);
        resolve(new Response(data.ok ? new Uint8Array(data.bytes) : null, { status: data.ok ? 200 : (data.status || 404), headers: data.contentType ? { 'content-type': data.contentType } : undefined }));
      });
      self.postMessage({ kind: 'kit-fetch', id, url });
    });
  };
})();
`;

/** Only the Sheet image worker gets the shim: a script that fetches by request key and listens for messages. */
export function injectWorkerSource(source, mime = '') {
  if (source.includes(SHIM_MARK)) return source;
  if (mime && !mime.includes('javascript') && !mime.includes('ecmascript')) return source;
  if (!source.includes('requestKey') || !source.includes('fetch(')) return source;
  if (!source.includes('addEventListener("message"') && !source.includes("addEventListener('message'")) return source;
  return `${WORKER_SHIM}\n${source}`;
}

/**
 * Installs the bridge on `window`. `read(reference, purpose)` returns `{ ok, status, contentType, bytes }`
 * for a Kit reference. The SDK only asks this way for pictures it draws, so the purpose is always `display`
 * (as in the reference Host); downloads go through the editor's own resolved URLs. Returns an uninstall
 * function (tests; the page keeps it for its lifetime).
 */
export function installSheetMediaBridge(window, read) {
  const { Blob: NativeBlob, Worker: NativeWorker } = window;
  const nativeFetch = window.fetch.bind(window);
  const answer = (url, purpose) => {
    const reference = sheetMediaReference(url);
    return reference ? read(reference, purpose).catch(() => ({ ok: false, status: 404 })) : Promise.resolve({ ok: false, status: 404 });
  };
  window.fetch = (input, init) => {
    const url = typeof input === 'string' ? input : input instanceof window.URL ? input.href : input?.url ?? '';
    if (!url.startsWith(SCHEME)) return nativeFetch(input, init);
    return answer(url, 'display').then(asset => new window.Response(asset.ok ? asset.bytes : null, { status: asset.ok ? 200 : asset.status || 404, headers: asset.contentType ? { 'content-type': asset.contentType } : undefined }));
  };
  function KitBlob(parts, options) {
    const text = Array.isArray(parts) && parts.every(part => typeof part === 'string') ? parts.join('') : null;
    if (text !== null) {
      const injected = injectWorkerSource(text, options?.type ?? '');
      if (injected !== text) return new NativeBlob([injected], { type: options?.type || 'application/javascript' });
    }
    return new NativeBlob(parts ?? [], options);
  }
  KitBlob.prototype = NativeBlob.prototype;
  function KitWorker(url, options) {
    let script = url;
    const prefix = 'data:application/javascript,';
    if (typeof script === 'string' && script.startsWith(prefix)) {
      const source = decodeURIComponent(script.slice(prefix.length));
      const injected = injectWorkerSource(source, 'application/javascript');
      if (injected !== source) script = `${prefix}${encodeURIComponent(injected)}`;
    }
    const worker = new NativeWorker(script, options);
    worker.addEventListener('message', event => {
      const data = event.data;
      if (!data || data.kind !== 'kit-fetch' || !data.id || typeof data.url !== 'string') return;
      void answer(data.url, 'display').then(asset => {
        const bytes = asset.ok && asset.bytes ? new Uint8Array(asset.bytes) : undefined;
        worker.postMessage({ kind: 'kit-result', id: data.id, ok: Boolean(bytes), status: asset.status, contentType: asset.contentType, bytes }, bytes ? [bytes.buffer] : []);
      });
    });
    return worker;
  }
  KitWorker.prototype = NativeWorker.prototype;
  window.Blob = KitBlob;
  window.Worker = KitWorker;
  return () => { window.fetch = nativeFetch; window.Blob = NativeBlob; window.Worker = NativeWorker; };
}
