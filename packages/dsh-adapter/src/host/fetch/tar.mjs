import { createHash } from 'node:crypto';
import { once } from 'node:events';
import { createWriteStream } from 'node:fs';
import { chmod, mkdir } from 'node:fs/promises';
import { dirname, join, resolve, sep } from 'node:path';
import { createGunzip } from 'node:zlib';

/**
 * A small streaming tar reader for the archives the plugin downloads (spec 261110-universal-package §4.2 step 5):
 * office-kit's Kit archives and editor package, and `git archive` output for skills. Node has no tar module and
 * Windows may have no `tar`, so this reads ustar with pax (`x`/`g`) and GNU long-name (`L`/`K`) extensions, including
 * UTF-8 names. Only regular files and directories are accepted: links, devices and anything else are rejected, as are
 * absolute names and names with `..`, so an archive can never write outside the target directory.
 */

const BLOCK = 512;
export class TarError extends Error {
  constructor(message) { super(message); this.name = 'TarError'; this.code = 'bad-archive'; }
}

const field = (block, offset, length) => {
  const raw = block.subarray(offset, offset + length);
  const end = raw.indexOf(0);
  return (end < 0 ? raw : raw.subarray(0, end)).toString('utf8');
};
function octal(block, offset, length) {
  const raw = block.subarray(offset, offset + length);
  // GNU base-256 for sizes over 8 GiB: not expected here, refused rather than misread.
  if (raw[0] & 0x80) throw new TarError('base-256 numbers are not supported');
  const text = field(block, offset, length).trim();
  if (!text) return 0;
  if (!/^[0-7]+$/.test(text)) throw new TarError(`bad number ${JSON.stringify(text)}`);
  return parseInt(text, 8);
}
function checksumOk(block) {
  const stored = octal(block, 148, 8);
  let sum = 0;
  for (let i = 0; i < BLOCK; i++) sum += i >= 148 && i < 156 ? 32 : block[i];
  return sum === stored;
}
/** pax records: `<len> <key>=<value>\n`. */
function parsePax(buffer) {
  const out = {};
  let at = 0;
  while (at < buffer.length) {
    const space = buffer.indexOf(0x20, at);
    if (space < 0) break;
    const length = Number(buffer.subarray(at, space).toString('ascii'));
    if (!Number.isInteger(length) || length <= 0) throw new TarError('bad pax record');
    const record = buffer.subarray(space + 1, at + length - 1).toString('utf8');
    const eq = record.indexOf('=');
    if (eq > 0) out[record.slice(0, eq)] = record.slice(eq + 1);
    at += length;
  }
  return out;
}

/**
 * The relative path an entry may be written to, or null to skip it. `strip` is a prefix every entry must have
 * (`packages/dsh-adapter/skills/shimo-pptx/` for a git archive, `package/` for an npm tarball); entries outside it
 * are rejected unless `outside: 'skip'`.
 */
export function safeName(name, { strip = '', outside = 'reject' } = {}) {
  let rel = name.replace(/\\/g, '/').replace(/^(\.\/)+/, '');
  if (!rel || rel === '.') return null;
  if (rel.startsWith('/') || /^[A-Za-z]:/.test(rel)) throw new TarError(`absolute path ${JSON.stringify(name)}`);
  const parts = rel.split('/').filter(part => part !== '' && part !== '.');
  if (parts.some(part => part === '..')) throw new TarError(`path with .. ${JSON.stringify(name)}`);
  rel = parts.join('/');
  if (strip) {
    const prefix = strip.replace(/\/+$/, '');
    if (rel === prefix) return null;
    if (!rel.startsWith(`${prefix}/`)) {
      if (outside === 'skip') return null;
      throw new TarError(`${JSON.stringify(name)} is outside ${prefix}/`);
    }
    rel = rel.slice(prefix.length + 1);
  }
  return rel || null;
}

/**
 * Reads a tar (gzip when `gzip`) from an async iterable of Buffers. For each regular file calls
 * `onFile({ name, size, mode, executable }, chunks)` where `chunks` is an async iterable of its bytes that must be
 * consumed before the next entry. Returns the number of files.
 */
export async function readTar(source, onFile, { gzip = false, strip = '', outside = 'reject', onDirectory = () => {} } = {}) {
  const input = gzip ? source.pipe ? source.pipe(createGunzip()) : gunzipIterable(source) : source;
  const reader = new ChunkReader(input[Symbol.asyncIterator]());
  let pax = {};
  let globalPax = {};
  let longName;
  let count = 0;
  for (;;) {
    const block = await reader.read(BLOCK);
    if (!block) throw new TarError('archive ended without its end blocks');
    if (block.every(byte => byte === 0)) break;
    if (!checksumOk(block)) throw new TarError('header checksum mismatch (not a tar, or corrupted)');
    const type = String.fromCharCode(block[156] || 0x30);
    let size = octal(block, 124, 12);
    const headerPadded = Math.ceil(size / BLOCK) * BLOCK;
    if (type === 'x' || type === 'g' || type === 'L' || type === 'K') {
      const data = await reader.read(headerPadded);
      if (!data) throw new TarError('archive ended inside an extended header');
      const body = data.subarray(0, size);
      if (type === 'x') pax = parsePax(body);
      else if (type === 'g') globalPax = { ...globalPax, ...parsePax(body) };
      else if (type === 'L') longName = field(body, 0, body.length);
      // 'K' (long link name) only matters for links, which are rejected below.
      continue;
    }
    const prefix = field(block, 345, 155);
    let name = longName ?? pax.path ?? globalPax.path ?? (prefix ? `${prefix}/${field(block, 0, 100)}` : field(block, 0, 100));
    if (pax.size !== undefined) size = Number(pax.size);
    const mode = octal(block, 100, 8);
    const entry = { pax, type };
    pax = {};
    longName = undefined;
    if (!Number.isSafeInteger(size) || size < 0) throw new TarError(`bad size for ${JSON.stringify(name)}`);
    const dataLength = Math.ceil(size / BLOCK) * BLOCK;
    if (type === '5') {
      // The prefix's own parent directories (git archive lists them) are not content: skipped, never rejected.
      const rel = safeName(name, { strip, outside: 'skip' });
      if (rel) onDirectory(rel);
      await reader.skip(dataLength);
      continue;
    }
    if (type !== '0' && type !== '\0' && type !== '7') throw new TarError(`${JSON.stringify(name)} is not a regular file (type ${JSON.stringify(type)}); links are not accepted`);
    const rel = safeName(name, { strip, outside });
    // macOS AppleDouble companions (`._name`, written by a Mac's tar for extended attributes) are not content; office-kit's
    // Mac archive carries them, and npm drops them from packages, so they are skipped everywhere.
    if (!rel || rel.split('/').at(-1).startsWith('._')) { await reader.skip(dataLength); continue; }
    let left = size;
    const chunks = (async function* () {
      while (left > 0) {
        const chunk = await reader.readSome(left);
        if (!chunk) throw new TarError(`archive ended inside ${rel}`);
        left -= chunk.length;
        yield chunk;
      }
    })();
    await onFile({ name: rel, size, mode, executable: Boolean(mode & 0o111) }, chunks);
    // A caller that stopped reading early: drain the rest of the file.
    for await (const _ of chunks) { /* drain */ }
    await reader.skip(dataLength - size);
    count++;
  }
  return count;
}

async function* gunzipIterable(source) {
  const gunzip = createGunzip();
  const feed = (async () => { try { for await (const chunk of source) if (!gunzip.write(chunk)) await new Promise(r => gunzip.once('drain', r)); gunzip.end(); } catch (error) { gunzip.destroy(error); } })();
  yield* gunzip;
  await feed;
}

class ChunkReader {
  constructor(iterator) { this.iterator = iterator; this.buffer = Buffer.alloc(0); this.done = false; }
  async fill(n) {
    while (this.buffer.length < n && !this.done) {
      const { value, done } = await this.iterator.next();
      if (done) { this.done = true; break; }
      this.buffer = this.buffer.length ? Buffer.concat([this.buffer, value]) : Buffer.from(value);
    }
  }
  async read(n) {
    if (n === 0) return Buffer.alloc(0);
    await this.fill(n);
    if (this.buffer.length < n) return null;
    const out = this.buffer.subarray(0, n);
    this.buffer = this.buffer.subarray(n);
    return out;
  }
  async readSome(max) {
    if (!this.buffer.length) await this.fill(1);
    if (!this.buffer.length) return null;
    const out = this.buffer.subarray(0, Math.min(max, this.buffer.length));
    this.buffer = this.buffer.subarray(out.length);
    return out;
  }
  async skip(n) {
    while (n > 0) {
      const chunk = await this.readSome(n);
      if (!chunk) throw new TarError('archive ended early');
      n -= chunk.length;
    }
  }
}

/**
 * Extracts into `dest` (which must exist or is created), hashing each file as it is written. Returns
 * `{ files: { [rel]: { size, sha256, sha512, executable } } }`; `sha256`/`sha512` are SRI strings (`sha256-<base64>`).
 * `keep(rel)` may skip files (return false).
 */
export async function extractTar(source, dest, { gzip = false, strip = '', outside = 'reject', keep = () => true } = {}) {
  const root = resolve(dest);
  await mkdir(root, { recursive: true });
  const files = {};
  await readTar(source, async (entry, chunks) => {
    if (!keep(entry.name)) return;
    if (Object.hasOwn(files, entry.name)) throw new TarError(`${entry.name} appears twice`);
    const target = join(root, ...entry.name.split('/'));
    if (target !== root && !target.startsWith(root + sep)) throw new TarError(`${entry.name} escapes the target`);
    await mkdir(dirname(target), { recursive: true });
    const h256 = createHash('sha256');
    const h512 = createHash('sha512');
    const out = createWriteStream(target, { flags: 'wx', mode: entry.executable ? 0o755 : 0o644 });
    const closed = new Promise((ok, fail) => { out.once('close', ok); out.once('error', fail); });
    try {
      for await (const chunk of chunks) {
        h256.update(chunk); h512.update(chunk);
        // events.once rejects on 'error' and removes both listeners afterwards.
        if (!out.write(chunk)) await once(out, 'drain');
      }
    } finally { out.end(); }
    await closed;
    if (entry.executable) await chmod(target, 0o755);
    files[entry.name] = { size: entry.size, sha256: `sha256-${h256.digest('base64')}`, sha512: `sha512-${h512.digest('base64')}`, executable: entry.executable };
  }, { gzip, strip, outside });
  return { files };
}

/** Lists and hashes every file without writing anything (tooling: manifests from an archive in memory). */
export async function hashTar(source, { gzip = false, strip = '', outside = 'reject' } = {}) {
  const files = {};
  await readTar(source, async (entry, chunks) => {
    const h256 = createHash('sha256');
    const h512 = createHash('sha512');
    for await (const chunk of chunks) { h256.update(chunk); h512.update(chunk); }
    if (Object.hasOwn(files, entry.name)) throw new TarError(`${entry.name} appears twice`);
    files[entry.name] = { size: entry.size, sha256: `sha256-${h256.digest('base64')}`, sha512: `sha512-${h512.digest('base64')}`, executable: entry.executable };
  }, { gzip, strip, outside });
  return { files };
}
