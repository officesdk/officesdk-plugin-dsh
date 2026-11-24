/** Only Session addresses carry an authorizing Session; absolute addresses never borrow one. */
export function parseSource(address) {
  const prefix = 'dsh-resource://file/session/';
  if (typeof address !== 'string' || !address.startsWith(prefix)) throw new Error('Documents open only from a Session file address');
  let parts;
  try { parts = address.slice(prefix.length).split(/[?#]/, 1)[0].split('/').map(decodeURIComponent); }
  catch { throw new Error('Invalid encoded document address'); }
  const [sessionId, ...rest] = parts;
  const path = rest.join('/');
  if (!sessionId || !path || /[\u0000-\u001f\u007f]/.test(sessionId + path)) throw new Error('Documents require a Session and a file path');
  return { sessionId, path };
}

export function unwrapRemote(result) {
  if (result?.status === 'success' || result?.type === 'success') return result.value;
  if (result?.ok === true) return result.value;
  const error = result?.error ?? result;
  if (result?.status === 'failure' || result?.type === 'failure' || result?.type === 'error' || result?.ok === false) {
    const failure = new Error(error?.message || 'Document request failed');
    failure.code = typeof error?.code === 'string' ? error.code : undefined;
    failure.details = error?.details;
    throw failure;
  }
  throw new Error('Invalid document RPC result');
}

/** Keep leading Host routes under the application mount, including dsh-app://app/. */
export function hostUrl(path, base) {
  if (typeof path !== 'string' || !path.startsWith('/shimo/')) throw new Error('Invalid local document route');
  return new URL(path.slice(1), base).href;
}

/** Remote codes arrive as officesdk/<code>; the UI groups them by what the user can do. */
export function failureKind(error) {
  const code = typeof error?.code === 'string' && error.code.startsWith('officesdk/') ? error.code.slice('officesdk/'.length) : undefined;
  if (code === 'busy' || code === 'timeout' || code === 'unavailable' || error?.code === 'gateway/cancelled') return { code, kind: 'retry' };
  if (code === 'source-changed' || code === 'edit-conflict') return { code, kind: 'conflict', draftId: error.details?.draftId };
  return { code, kind: 'fixed' };
}

/**
 * The Session file address for a path: workspace-relative, or absolute (its leading slash kept, as DSH spells it).
 * Every tab and AI command builds it here, so one file always has one address. `platform` is the Host's (a lease's
 * `platform`): on Windows '\' separates names too (a Windows name cannot contain it), so `sub\a.docx` and
 * `sub/a.docx` are one address. How DSH itself spells Windows Session addresses (drive letter, separators) is still
 * to be confirmed on a real machine.
 */
export function fileAddress(sessionId, path, platform) {
  return `dsh-resource://file/session/${encodeURIComponent(sessionId)}/${String(path).split(platform === 'win32' ? /[\\/]/ : '/').map(encodeURIComponent).join('/')}`;
}
