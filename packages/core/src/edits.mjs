import { createHash } from 'node:crypto';
import { cp, mkdir, readFile, stat, writeFile } from 'node:fs/promises';
import { dirname, join } from 'node:path';
import { fail } from './errors.mjs';
import { safeRelative } from './paths.mjs';
import { listZip, readZipEntry } from './zip.mjs';

const MiB = 1024 * 1024;
const MAX_ASSET_BYTES = 100 * MiB;

/** Same whitelist as Writer's normalizeMowArchivePath and the converter's, so round-trips cannot drop files. */
export function mowArchivePath(name) {
  if (!name || name.startsWith('/') || name.includes('\\') || name.split('/').some(s => !s || s === '.' || s === '..')) fail('invalid-body', 'Invalid MOW archive path.');
  const [root] = name.split('/');
  if (name !== 'content.json' && root !== 'image' && root !== 'embedding') fail('invalid-body', 'Unsupported MOW archive entry.');
  return name;
}

/** Writer leaves runtime nodeId keys behind; the pinned word2mow rejects them inside tables. */
export function dropWriterRuntimeNodeIds(value) {
  let removed = false;
  const visit = node => {
    if (Array.isArray(node)) { for (const child of node) visit(child); return; }
    if (node === null || typeof node !== 'object') return;
    if (Object.hasOwn(node, 'nodeId')) { delete node.nodeId; removed = true; }
    for (const child of Object.values(node)) visit(child);
  };
  visit(value);
  return removed;
}

const INLINE_EXTENSIONS = { 'image/jpeg': '.jpg', 'image/jpg': '.jpg', 'image/png': '.png', 'image/webp': '.webp', 'image/gif': '.gif', 'image/svg+xml': '.svg', 'image/bmp': '.bmp', 'video/mp4': '.mp4', 'audio/mpeg': '.mp3' };

/** Editor fills arrive as data: URIs; staging them as package media keeps the snapshot exportable. */
export async function stageInlineResources(text, directory) {
  if (!text.includes('data:')) return text;
  let document;
  try { document = JSON.parse(text); } catch { fail('invalid-body', 'The presentation snapshot is not valid JSON.'); }
  const staged = new Map();
  const pending = [];
  const visit = node => {
    if (Array.isArray(node)) { for (const item of node) visit(item); return; }
    if (node === null || typeof node !== 'object') return;
    if (typeof node.resourceUri === 'string' && node.resourceUri.startsWith('data:')) pending.push(node);
    for (const item of Object.values(node)) visit(item);
  };
  visit(document);
  for (const node of pending) {
    const uri = node.resourceUri;
    let resource = staged.get(uri);
    if (!resource) {
      const comma = uri.indexOf(',');
      const header = uri.slice(5, comma);
      if (comma < 0 || !header.includes('base64')) continue;
      const contentType = header.split(';')[0].trim().toLowerCase();
      const data = Buffer.from(uri.slice(comma + 1), 'base64');
      if (!data.length) continue;
      if (data.length > MAX_ASSET_BYTES) fail('input-too-large', 'An inline resource exceeds the byte limit.');
      const digest = createHash('sha256').update(data).digest('hex');
      const extension = INLINE_EXTENSIONS[contentType] ?? '.bin';
      const name = digest + extension;
      const target = join(directory, 'media', name);
      await mkdir(dirname(target), { recursive: true, mode: 0o700 });
      const existing = await stat(target).catch(() => undefined);
      if (existing?.size !== data.length) await writeFile(target, data, { mode: 0o600 });
      resource = { resourceUri: 'mop-asset:/media/' + name, contentType, extension: extension.slice(1), digest: 'sha256:' + digest, resourceSize: data.length };
      staged.set(uri, resource);
    }
    Object.assign(node, resource);
  }
  return staged.size ? JSON.stringify(document) : text;
}

/** Each format's editor writes back through a fixed set of parts; anything else is rejected. */
export const EDITS = Object.freeze({
  docx: {
    /** A complete MOW zip replaces the snapshot; the caller swaps directories atomically. */
    async package(snapshot, part, body, limits) {
      if (part !== 'package') fail('not-found', 'Unknown document part.');
      const archive = listZip(body, { maxEntries: limits.maxOutputFiles });
      let total = 0, sawContent = false;
      const files = [];
      for (const entry of archive.entries) {
        if (entry.name.endsWith('/')) continue;
        const name = mowArchivePath(entry.name);
        const data = readZipEntry(archive, entry, limits.maxStageBytes - total);
        total += data.length;
        if (name === 'content.json') sawContent = data.length > 0;
        files.push([name, data]);
      }
      if (!sawContent) fail('invalid-body', 'The MOW package has no content.json at its root.');
      return { replace: true, files, content: 'content.json', assets: ['image', 'embedding'].filter(root => files.some(([name]) => name.startsWith(root + '/'))) };
    },
    /** Export works on a scratch copy so the live snapshot keeps what the editor sent. */
    async prepareExport(snapshot, scratch) {
      const copy = join(scratch, 'export-source');
      await cp(snapshot.directory, copy, { recursive: true, verbatimSymlinks: true });
      const contentPath = join(copy, snapshot.content);
      const document = JSON.parse(await readFile(contentPath, 'utf8'));
      if (dropWriterRuntimeNodeIds(document)) await writeFile(contentPath, JSON.stringify(document));
      return copy;
    },
  },
  xlsx: {
    async package(snapshot, part, body) {
      if (part !== 'content') fail('not-found', 'Unknown spreadsheet part.');
      if (!body.length) fail('invalid-body', 'The spreadsheet content is empty.');
      return { files: [[snapshot.content, body]] };
    },
    async prepareExport(snapshot) { return snapshot.directory; },
  },
  pptx: {
    async package(snapshot, part, body) {
      if (part === 'snapshot') {
        if (!body.length) fail('invalid-body', 'The presentation snapshot is empty.');
        // Binary MOP snapshots are written as sent; only JSON snapshots can carry inline data: fills.
        if (!snapshot.content.endsWith('.json')) return { files: [[snapshot.content, body]] };
        const text = await stageInlineResources(body.toString('utf8'), snapshot.directory);
        return { files: [[snapshot.content, Buffer.from(text)]] };
      }
      if (part.startsWith('asset/')) {
        const path = safeRelative(part.slice('asset/'.length));
        const [root] = path.split('/');
        if (root !== 'media' && root !== 'embeddings') fail('invalid-path', 'Invalid presentation asset path.');
        if (!body.length || body.length > MAX_ASSET_BYTES) fail('invalid-body', 'Invalid presentation asset size.');
        const digest = createHash('sha256').update(body).digest('hex');
        if (!path.split('/').at(-1).toLowerCase().startsWith(digest + '.')) fail('invalid-body', 'The asset name does not match its content digest.');
        return { files: [[path, body]] };
      }
      fail('not-found', 'Unknown presentation part.');
    },
    async prepareExport(snapshot) { return snapshot.directory; },
  },
});
