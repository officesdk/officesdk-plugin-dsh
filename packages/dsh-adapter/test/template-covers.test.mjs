import test from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync, readdirSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import path from 'node:path';
import { BUILTIN } from '../src/client/inspiration.mjs';
import { TEMPLATE_COVERS } from '../src/client/template-covers.mjs';

const covers = fileURLToPath(new URL('../src/client/covers/', import.meta.url));
const pngs = readdirSync(covers).filter(name => name.endsWith('.png')).sort();
/** Width and height from a PNG's IHDR. */
const size = buffer => {
  assert.equal(buffer.subarray(1, 4).toString('latin1'), 'PNG');
  return [buffer.readUInt32BE(16), buffer.readUInt32BE(20)];
};

test('covers: template-covers.mjs carries exactly the PNGs in src/client/covers, byte for byte', () => {
  const fromFiles = { word: {}, excel: {} };
  for (const name of pngs) {
    const [, kind, id] = name.match(/^(word|excel)-([a-z0-9-]+)\.png$/) ?? [];
    assert.ok(kind, `${name}: named <word|excel>-<id>.png`);
    fromFiles[kind][id] = `data:image/png;base64,${readFileSync(path.join(covers, name)).toString('base64')}`;
  }
  assert.deepEqual(JSON.parse(JSON.stringify(TEMPLATE_COVERS)), fromFiles, 'rerun node tooling/covers/covers-module.mjs after changing a PNG');
});

test('covers: every built-in Word and Excel template has one, and nothing else does', () => {
  for (const kind of ['word', 'excel']) assert.deepEqual(Object.keys(TEMPLATE_COVERS[kind]).sort(), BUILTIN[kind].map(item => item.id).sort());
});

test('covers: Word portrait 595:842, Excel 4:3, at least 2× the card width; all together within 100 KB (they ship in client.js)', () => {
  let total = 0;
  for (const name of pngs) {
    const buffer = readFileSync(path.join(covers, name));
    total += buffer.length;
    const [width, height] = size(buffer);
    const [ratio, minWidth] = name.startsWith('word-') ? [595 / 842, 360] : [4 / 3, 400];
    assert.ok(Math.abs(width / ratio - height) <= 1, `${name}: ${width}×${height} keeps ${ratio.toFixed(4)} within 1px`);
    assert.ok(width >= minWidth, `${name}: ${width}px wide`);
  }
  assert.ok(total <= 100 * 1024, `covers total ${total} bytes`);
});
