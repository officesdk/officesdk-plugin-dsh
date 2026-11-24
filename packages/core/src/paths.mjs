import { constants } from 'node:fs';
import { lstat, open, readdir, realpath } from 'node:fs/promises';
import { extname, join, sep } from 'node:path';
import { fail } from './errors.mjs';

/** Names have already crossed the single HTTP decode boundary; never decode them again. */
export function safeRelative(path) {
  if (typeof path !== 'string' || !path || path.length > 4096 || /%[0-9a-f]{2}|[\\\0?#]/i.test(path)
      || path.startsWith('/') || path.split('/').some(p => !p || p === '.' || p === '..')) {
    fail('invalid-path', 'Invalid resource path.');
  }
  return path;
}

/** Raw request paths are checked before URL normalization can remove traversal segments. */
export function requestPath(url, prefix) {
  const raw = (url ?? '').split('?')[0];
  if (/%(?:2f|5c|00)/i.test(raw)) fail('invalid-path', 'Invalid resource path.');
  let path;
  try { path = decodeURIComponent(raw); } catch { fail('invalid-path', 'Invalid resource path.'); }
  if (!path.startsWith(prefix + '/')) fail('not-found', 'Route not found.');
  return safeRelative(path.slice(prefix.length + 1));
}

export async function confinedFile(root, relative) {
  safeRelative(relative);
  const rootReal = await realpath(root);
  let candidate = rootReal;
  const parts = relative.split('/');
  for (let i = 0; i < parts.length; i++) {
    candidate = join(candidate, parts[i]);
    const info = await lstat(candidate);
    if (info.isSymbolicLink()) fail('invalid-output', 'Document output contains a symlink.');
    if (i < parts.length - 1 && !info.isDirectory()) fail('invalid-path', 'Invalid resource path.');
    if (i === parts.length - 1 && !info.isFile()) fail('invalid-path', 'Resource path is not a file.');
  }
  const actual = await realpath(candidate);
  if (!actual.startsWith(rootReal + sep)) fail('invalid-path', 'Invalid resource path.');
  return actual;
}

export async function readConfined(root, path, cap, signal) {
  signal.throwIfAborted();
  const target = await confinedFile(root, path);
  const handle = await open(target, constants.O_RDONLY | constants.O_NOFOLLOW);
  try {
    const info = await handle.stat();
    if (!info.isFile() || info.size > cap) fail('output-too-large', 'Document output exceeds the byte limit.');
    const body = await handle.readFile({ signal });
    if (body.byteLength > cap) fail('output-too-large', 'Document output exceeds the byte limit.');
    return body;
  } finally { await handle.close(); }
}

/**
 * Strict trees (published output) reject symlinks and special files and count entries.
 * Lenient trees (converter scratch) only count bytes: components may leave sockets or links there.
 */
export async function measure(root, { maxBytes, maxEntries, strict = true }) {
  let bytes = 0, count = 0;
  async function walk(dir, depth) {
    if (depth > 32) fail('invalid-output', 'Document output path is too deep.');
    let names;
    try { names = await readdir(dir); } catch (error) { if (error.code === 'ENOENT' && depth === 0) return; throw error; }
    for (const name of names) {
      const path = join(dir, name);
      let info;
      try { info = await lstat(path); } catch (error) { if (error.code === 'ENOENT') continue; throw error; }
      if (strict && ++count > maxEntries) fail('output-too-large', 'Document output exceeds the entry limit.');
      if (info.isSymbolicLink()) { if (strict) fail('invalid-output', 'Document output contains a symlink.'); }
      else if (info.isDirectory()) await walk(path, depth + 1);
      else if (info.isFile()) {
        bytes += info.size;
        if (bytes > maxBytes) fail('output-too-large', 'Document output exceeds the byte limit.');
      } else if (strict) fail('invalid-output', 'Document output is not a regular file.');
    }
  }
  await walk(root, 0);
  return bytes;
}

const TYPES = { '.html': 'text/html; charset=utf-8', '.js': 'text/javascript; charset=utf-8', '.mjs': 'text/javascript; charset=utf-8', '.css': 'text/css; charset=utf-8', '.json': 'application/json; charset=utf-8', '.wasm': 'application/wasm', '.png': 'image/png', '.jpg': 'image/jpeg', '.jpeg': 'image/jpeg', '.gif': 'image/gif', '.webp': 'image/webp', '.bmp': 'image/bmp', '.svg': 'image/svg+xml', '.woff': 'font/woff', '.woff2': 'font/woff2', '.ttf': 'font/ttf', '.otf': 'font/otf', '.mp4': 'video/mp4', '.mp3': 'audio/mpeg', '.docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', '.pptx': 'application/vnd.openxmlformats-officedocument.presentationml.presentation' };
export const mime = path => TYPES[extname(path).toLowerCase()] ?? 'application/octet-stream';
