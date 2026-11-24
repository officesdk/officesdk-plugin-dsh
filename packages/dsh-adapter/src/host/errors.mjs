import { ERROR_STATUS, SuiteError, isSuiteError } from '../../../core/src/index.mjs';

/** Harness WorkspaceFiles failures become suite codes; callers never branch on message text. */
const WORKSPACE_FILE = {
  'workspace-file/not-found': 'not-found',
  'workspace-file/too-large': 'input-too-large',
  'workspace-file/not-regular-file': 'invalid-path',
  'workspace-file/not-directory': 'invalid-path',
  'workspace-file/unknown-workspace': 'forbidden',
  'workspace-file/unsupported-address': 'forbidden',
  'workspace-file/outside-workspace': 'forbidden',
};

export function fromWorkspaceFile(error) {
  if (isSuiteError(error) || error?.name === 'AbortError') return error;
  const code = typeof error?.code === 'string' ? error.code : '';
  if (code.startsWith('workspace-file/')) return new SuiteError(WORKSPACE_FILE[code] ?? 'unavailable', 'The workspace file is unavailable.', { source: code });
  if (code.startsWith('gateway/lookup')) return new SuiteError('forbidden', 'The Session is unavailable.');
  return error;
}

/** Kit-host errors share Core's codes; one Core has no name for becomes `unavailable`. */
export function fromKit(error) {
  if (error?.name !== 'KitError') return error;
  return new SuiteError(error.code in ERROR_STATUS ? error.code : 'unavailable', error.message, error.details);
}

/** Suite codes ride under one namespace with their HTTP status; anything unknown stays opaque. */
export function toRemote(error, signal, RemoteError) {
  error = fromKit(error);
  if (signal?.aborted) return new RemoteError('gateway/cancelled', 'The document operation was cancelled.', {});
  if (isSuiteError(error)) {
    const { draftId, limit, resource, status: componentStatus, outside, rebuild, overwrite, engine } = error.details ?? {};
    return new RemoteError(`officesdk/${error.code}`, error.message, { status: error.status,
      ...(draftId && { draftId }), ...(limit && { limit }), ...(resource && { resource }), ...(componentStatus !== undefined && { componentStatus }),
      ...(outside === true && { outside }), ...(rebuild === true && { rebuild }), ...(overwrite === true && { overwrite }), ...(typeof engine === 'string' && { engine }) });
  }
  if (error?.isDSHRemoteError) return error;
  return new RemoteError('gateway/internal', 'The document service is unavailable.', {});
}

/** Stable identity for drafts: owners share one key, other principals key by every field. */
export function principalKey(principal) {
  if (!principal || principal.kind === 'owner') return 'owner';
  return JSON.stringify(Object.keys(principal).sort().map(key => [key, principal[key]]));
}
