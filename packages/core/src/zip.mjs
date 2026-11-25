import { inflateRawSync } from 'node:zlib';
import { fail } from './errors.mjs';

/** A bounded reader for the stored and deflated entries Office files and MOW packages use. */
export function listZip(data, { maxEntries = 10_000 } = {}) {
  const buffer = Buffer.from(data.buffer, data.byteOffset, data.byteLength);
  const floor = Math.max(0, buffer.length - 0xffff - 22);
  let end = -1;
  for (let i = buffer.length - 22; i >= floor; i--) if (buffer.readUInt32LE(i) === 0x06054b50) { end = i; break; }
  if (end < 0) fail('invalid-format', 'The file is not a ZIP package.');
  const count = buffer.readUInt16LE(end + 10), size = buffer.readUInt32LE(end + 12);
  let offset = buffer.readUInt32LE(end + 16);
  if (count > maxEntries) fail('input-too-large', 'The ZIP package has too many entries.');
  if (offset + size > end) fail('invalid-format', 'The ZIP directory is out of range.');
  const entries = [];
  for (let i = 0; i < count; i++) {
    if (buffer.readUInt32LE(offset) !== 0x02014b50) fail('invalid-format', 'The ZIP directory is corrupt.');
    const flags = buffer.readUInt16LE(offset + 8), method = buffer.readUInt16LE(offset + 10);
    const compressedSize = buffer.readUInt32LE(offset + 20), uncompressedSize = buffer.readUInt32LE(offset + 24);
    const nameLength = buffer.readUInt16LE(offset + 28), extraLength = buffer.readUInt16LE(offset + 30), commentLength = buffer.readUInt16LE(offset + 32);
    const local = buffer.readUInt32LE(offset + 42);
    const name = buffer.toString(flags & 0x800 ? 'utf8' : 'latin1', offset + 46, offset + 46 + nameLength);
    if (flags & 0x1) fail('password-protected', 'Encrypted ZIP entries are not supported.');
    entries.push({ name, method, compressedSize, uncompressedSize, local });
    offset += 46 + nameLength + extraLength + commentLength;
  }
  return { buffer, entries };
}

export function readZipEntry({ buffer }, entry, maxBytes) {
  if (entry.uncompressedSize > maxBytes) fail('input-too-large', 'A ZIP entry exceeds the byte limit.');
  if (buffer.readUInt32LE(entry.local) !== 0x04034b50) fail('invalid-format', 'The ZIP entry header is corrupt.');
  const start = entry.local + 30 + buffer.readUInt16LE(entry.local + 26) + buffer.readUInt16LE(entry.local + 28);
  const raw = buffer.subarray(start, start + entry.compressedSize);
  if (raw.length !== entry.compressedSize) fail('invalid-format', 'The ZIP entry is truncated.');
  if (entry.method === 0) return raw;
  if (entry.method !== 8) fail('invalid-format', 'Unsupported ZIP compression method.');
  try { return inflateRawSync(raw, { maxOutputLength: Math.max(1, Math.min(maxBytes, entry.uncompressedSize || maxBytes)) }); }
  catch { fail('invalid-format', 'A ZIP entry could not be decompressed.'); }
}

/** Saved Office files must at least carry their content types and the format's main part. */
export function checkOfficePackage(data, mainPart) {
  const { entries } = listZip(data);
  const names = new Set(entries.map(entry => entry.name));
  if (!names.has('[Content_Types].xml') || !names.has(mainPart)) fail('invalid-output', 'The exported file is not a complete Office package.');
}
