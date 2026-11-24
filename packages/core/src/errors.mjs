/** One table owns every code and its HTTP status; callers branch on code, never on message. */
export const ERROR_STATUS = Object.freeze({
  'forbidden': 403, 'sandbox-denied': 403,
  'invalid-path': 400, 'invalid-revision': 400, 'invalid-body': 400, 'invalid-edit': 400,
  'unsupported-format': 400, 'unsupported-operation': 400,
  'password-protected': 422, 'invalid-format': 422,
  'input-too-large': 413, 'output-too-large': 413,
  'busy': 429,
  'source-changed': 409, 'edit-conflict': 409, 'exists': 409, 'same-file': 409, 'undo-unavailable': 409,
  // Save As over a file another program holds open (Windows refuses the replace): close it there and save again.
  'target-busy': 409,
  // A plugin update while a document still has unsaved changes (in any window): save first.
  'unsaved-changes': 409,
  // Selection rewrite (selection_rewrite): an id this Session does not have, a selection the user changed or took back,
  // and one already rewritten once.
  'selection-unknown': 404, 'selection-stale': 409, 'selection-used': 409,
  'lease-unavailable': 404, 'not-found': 404,
  'read-only': 405,
  'timeout': 504, 'cancelled': 499,
  'converter-failed': 422, 'invalid-output': 422, 'editor-failed': 422,
  'unavailable': 503, 'no-editor': 503,
});

export class SuiteError extends Error {
  constructor(code, message, details) {
    if (!(code in ERROR_STATUS)) throw new TypeError(`Unknown suite error code ${code}.`);
    super(message);
    this.name = 'SuiteError';
    this.code = code;
    this.status = ERROR_STATUS[code];
    if (details !== undefined) this.details = details;
  }
}
export const fail = (code, message, details) => { throw new SuiteError(code, message, details); };
/**
 * The release bundles Core once per entry (lib/host, lib/tools), so an error thrown by the Host's copy is not an
 * instance of the tools' class: a SuiteError is also anything named SuiteError whose code is one of ours.
 * Own keys only: `in` would also accept inherited names such as `toString`.
 */
export const isSuiteError = error => error instanceof SuiteError
  || (error?.name === 'SuiteError' && typeof error.code === 'string' && Object.hasOwn(ERROR_STATUS, error.code));

/** Abort reasons from timers and services become their own codes; anything else is a cancellation. */
export function fromAbort(signal) {
  const reason = signal?.reason;
  return isSuiteError(reason) ? reason : new SuiteError('cancelled', 'The document operation was cancelled.');
}
