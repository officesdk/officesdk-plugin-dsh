import { gzipSync } from 'node:zlib';

/**
 * Builds tar archives for tests. Entries: `{ name, data?, mode?, type? ('0' file, '5' dir, '2' symlink, '1' hard link),
 * linkname?, pax? ({ key: value }), longName? (GNU 'L' record) }`. A name over 100 bytes without `pax`/`longName` uses
 * the ustar prefix field.
 */
const BLOCK = 512;
function header({ name, size = 0, mode = 0o644, type = '0', linkname = '' }) {
  const block = Buffer.alloc(BLOCK);
  let base = name;
  let prefix = '';
  if (Buffer.byteLength(name) > 100) {
    const cut = name.lastIndexOf('/', 155);
    prefix = name.slice(0, cut); base = name.slice(cut + 1);
  }
  block.write(base, 0, 100, 'utf8');
  block.write(mode.toString(8).padStart(7, '0') + '\0', 100, 8, 'ascii');
  block.write('0000000\0', 108, 8, 'ascii');
  block.write('0000000\0', 116, 8, 'ascii');
  block.write(size.toString(8).padStart(11, '0') + '\0', 124, 12, 'ascii');
  block.write('00000000000\0', 136, 12, 'ascii');
  block.write('        ', 148, 8, 'ascii');
  block.write(type, 156, 1, 'ascii');
  block.write(linkname, 157, 100, 'utf8');
  block.write('ustar\0', 257, 6, 'ascii');
  block.write('00', 263, 2, 'ascii');
  block.write(prefix, 345, 155, 'utf8');
  let sum = 0;
  for (const byte of block) sum += byte;
  block.write(sum.toString(8).padStart(6, '0') + '\0 ', 148, 8, 'ascii');
  return block;
}
const pad = data => Buffer.concat([data, Buffer.alloc((BLOCK - (data.length % BLOCK)) % BLOCK)]);
const paxBody = fields => Buffer.from(Object.entries(fields).map(([key, value]) => {
  const tail = ` ${key}=${value}\n`;
  let length = Buffer.byteLength(tail) + 1;
  while (String(length).length + Buffer.byteLength(tail) !== length) length = String(length).length + Buffer.byteLength(tail);
  return `${length}${tail}`;
}).join(''), 'utf8');

export function makeTar(entries, { gzip = false, end = true } = {}) {
  const parts = [];
  for (const entry of entries) {
    const data = entry.data === undefined ? Buffer.alloc(0) : Buffer.from(entry.data);
    if (entry.pax) {
      const body = paxBody(entry.pax);
      parts.push(header({ name: 'PaxHeader', size: body.length, type: 'x' }), pad(body));
    }
    if (entry.longName) {
      const body = Buffer.from(`${entry.longName}\0`, 'utf8');
      parts.push(header({ name: '././@LongLink', size: body.length, type: 'L' }), pad(body));
    }
    const type = entry.type ?? '0';
    parts.push(header({ name: entry.longName ? entry.longName.slice(0, 99) : entry.name, size: type === '0' ? data.length : 0, mode: entry.mode, type, linkname: entry.linkname }));
    if (type === '0' && data.length) parts.push(pad(data));
  }
  if (end) parts.push(Buffer.alloc(BLOCK * 2));
  const tar = Buffer.concat(parts);
  return gzip ? gzipSync(tar) : tar;
}
