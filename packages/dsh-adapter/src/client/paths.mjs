/**
 * Paths as the Host's platform spells them. The Host says which (`platform` on a document lease: Node's
 * `process.platform`); the Client never guesses from the browser. Without it, or for any platform but 'win32',
 * POSIX rules apply, so macOS behaves exactly as before.
 */
const windows = platform => platform === 'win32';

/**
 * The root a path starts from, separator included ('' for a relative path): '/'; on Windows also 'C:\' or 'C:/',
 * '\\server\share\' and a rooted '\…' or '/…' (the Host resolves those on the workspace's drive).
 */
export function rootOf(path, platform) {
  if (typeof path !== 'string') return '';
  if (!windows(platform)) return path.startsWith('/') ? '/' : '';
  return /^(?:[A-Za-z]:[\\/]|[\\/]{2}[^\\/]+[\\/][^\\/]+[\\/]?|[\\/])/.exec(path)?.[0] ?? '';
}

/** Absolute: `/…`; on Windows also `X:\…`, `X:/…` and `\\server\share\…`. */
export const isAbsolute = (path, platform) => rootOf(path, platform) !== '';

/** The path's names in order, separators and empty names dropped (a root's drive, 'C:', is the first name). */
export const segments = (path, platform) => String(path ?? '').split(windows(platform) ? /[\\/]+/ : /\/+/).filter(Boolean);

function lastSeparator(path, platform) {
  if (!windows(platform)) return path.lastIndexOf('/');
  return Math.max(path.lastIndexOf('/'), path.lastIndexOf('\\'));
}

/** The last name: 'a.docx' for 'sub/a.docx' (and for 'C:\sub\a.docx' on Windows). */
export function baseName(path, platform) {
  path = String(path ?? '');
  return path.slice(lastSeparator(path, platform) + 1);
}

/** The folder holding `path`: '' for a bare name, the root itself for a file at the root ('/', 'C:\'). */
export function parentOf(path, platform) {
  path = String(path ?? '');
  const at = lastSeparator(path, platform);
  if (at < 0) return '';
  const root = rootOf(path, platform);
  return at < root.length ? root : path.slice(0, at);
}

/** `name` inside `folder` ('' is the workspace), in the folder's own spelling: a Windows folder written with '\' keeps it. */
export function joinPath(folder, name, platform) {
  folder = String(folder ?? '');
  if (!folder) return name;
  const glue = windows(platform) && folder.includes('\\') ? '\\' : '/';
  return `${folder.replace(windows(platform) ? /[\\/]$/ : /\/$/, '')}${glue}${name}`;
}
