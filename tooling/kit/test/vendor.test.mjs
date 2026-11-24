import test from 'node:test';
import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import { mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { gzipSync } from 'node:zlib';
import { fetchVendor, pinProblems, tarballUrl, vendorProblems } from '../vendor.mjs';

const sha256 = bytes => createHash('sha256').update(bytes).digest('hex');
const sri512 = bytes => `sha512-${createHash('sha512').update(bytes).digest('base64')}`;
/** A gzipped ustar archive of `files` ({ name: text }) under package/. */
function tarball(files) {
  const blocks = [];
  for (const [name, text] of Object.entries(files)) {
    const body = Buffer.from(text);
    const header = Buffer.alloc(512);
    header.write(`package/${name}`, 0, 'utf8');
    header.write('0000644\0', 100); header.write('0000000\0', 108); header.write('0000000\0', 116);
    header.write(body.length.toString(8).padStart(11, '0') + '\0', 124);
    header.write('00000000000\0', 136); header.write('        ', 148); header.write('0', 156); header.write('ustar\0', 257); header.write('00', 263);
    let sum = 0; for (const byte of header) sum += byte;
    header.write(sum.toString(8).padStart(6, '0') + '\0 ', 148);
    blocks.push(header, body, Buffer.alloc((512 - (body.length % 512)) % 512));
  }
  blocks.push(Buffer.alloc(1024));
  return gzipSync(Buffer.concat(blocks));
}
const MANIFEST = JSON.stringify({ officeKit: { version: '1.0.0' }, kit: { 'darwin-arm64': { binary: 'k' }, 'openharmony-arm64': { module: 'k.so' } }, editor: { bundle: { sha256: `sha256-${createHash('sha256').update('editor').digest('base64')}` } } });
const RELEASE = 'https://releases.example/core/v1.0.0/';
const TO = { manifest: 'components/kit-manifest.json', editor: 'components/vendor/editor/bundle/offline-editor.js', fonts: 'components/vendor/host-fonts/sheet-host-fonts.css' };
const EDITOR_TGZ = { 'dist/bundle/offline-editor.js': 'editor' };
const FONTS_TGZ = { 'resources/host-fonts/sheet-host-fonts.css': 'fonts' };

/** A repository whose pin takes the manifest from a release (with SHA256SUMS) and the other two files from two npm tarballs. */
function repo(t, { tamper, manifest = MANIFEST, sums } = {}) {
  const root = mkdtempSync(path.join(tmpdir(), 'vendor-'));
  t.after(() => rmSync(root, { recursive: true, force: true }));
  const editor = tarball(EDITOR_TGZ), fonts = tarball(FONTS_TGZ);
  const pin = { v: 2, officeKit: { version: '1.0.0', platforms: ['darwin-arm64'] },
    sources: {
      release: { url: RELEASE, sums: 'SHA256SUMS' },
      editor: { registry: 'https://registry.example/', package: '@scope/editor', version: '1.0.0', integrity: sri512(editor) },
      fonts: { registry: 'https://registry.example/', package: '@scope/plugin', version: '0.0.1', integrity: sri512(fonts) },
    },
    files: {
      'kit-manifest.json': { from: 'release', path: 'kit-manifest.json', to: TO.manifest, sha256: sha256(Buffer.from(manifest)) },
      'offline-editor.js': { from: 'editor', path: 'dist/bundle/offline-editor.js', to: TO.editor, sha256: sha256(Buffer.from('editor')) },
      'sheet-host-fonts.css': { from: 'fonts', path: 'resources/host-fonts/sheet-host-fonts.css', to: TO.fonts, sha256: sha256(Buffer.from('fonts')) },
    } };
  tamper?.(pin);
  mkdirSync(path.join(root, 'components'), { recursive: true });
  writeFileSync(path.join(root, 'components/office-kit.json'), JSON.stringify(pin));
  const served = {
    [`${RELEASE}kit-manifest.json`]: Buffer.from(manifest),
    [`${RELEASE}SHA256SUMS`]: Buffer.from(sums ?? `${sha256(Buffer.from(manifest))}  kit-manifest.json\n${'1'.repeat(64)}  other.tgz\n`),
    'https://registry.example/@scope/editor/-/editor-1.0.0.tgz': editor,
    'https://registry.example/@scope/plugin/-/plugin-0.0.1.tgz': fonts,
  };
  const asked = [];
  const fetchImpl = async url => { asked.push(url); return served[url] ? new Response(served[url]) : new Response('', { status: 404 }); };
  return { root, fetchImpl, asked, served };
}

test('the registry address of a pinned npm source', () => {
  assert.equal(tarballUrl({ registry: 'https://registry.npmjs.org', package: '@officesdk/officesdk-core-editor', version: '0.3.0-rc.1' }),
    'https://registry.npmjs.org/@officesdk/officesdk-core-editor/-/officesdk-core-editor-0.3.0-rc.1.tgz');
});

test('fetch takes each file from its own source and check accepts them; an extra or edited file is reported', async t => {
  const { root, fetchImpl, asked } = repo(t);
  const written = await fetchVendor({ repo: root, fetchImpl, log: () => {} });
  assert.deepEqual(written.sort(), Object.values(TO).sort());
  assert.deepEqual(vendorProblems(root), []);
  assert.ok(asked.includes(`${RELEASE}SHA256SUMS`), 'the release sums were read');
  // A second fetch reads the cache alone (offline): SHA256SUMS is cached too, since it lists the pinned file as pinned.
  asked.length = 0;
  await fetchVendor({ repo: root, fetchImpl: async url => { asked.push(url); throw new Error('offline'); }, log: () => {} });
  assert.deepEqual(asked, []);
  writeFileSync(path.join(root, TO.editor), 'edited by hand');
  writeFileSync(path.join(root, 'components/vendor/extra.js'), 'x');
  const problems = vendorProblems(root).sort();
  assert.equal(problems.length, 2, problems.join('\n'));
  assert.match(problems[0], /^components\/vendor\/editor\/bundle\/offline-editor\.js differs from @scope\/editor@1\.0\.0 dist\/bundle\/offline-editor\.js/);
  assert.equal(problems[1], 'components/vendor/extra.js is not in components/office-kit.json');
});

test('a release file is checked against its pinned sha256 and against the release SHA256SUMS', async t => {
  const wrongSums = repo(t, { sums: `${'0'.repeat(64)}  kit-manifest.json\n` });
  await assert.rejects(fetchVendor({ repo: wrongSums.root, fetchImpl: wrongSums.fetchImpl, log: () => {} }), /SHA256SUMS does not list kit-manifest\.json with the pinned sha256/);
  const unlisted = repo(t, { sums: `${'0'.repeat(64)}  other.json\n` });
  await assert.rejects(fetchVendor({ repo: unlisted.root, fetchImpl: unlisted.fetchImpl, log: () => {} }), /does not list kit-manifest\.json/);
  const other = repo(t);
  other.served[`${RELEASE}kit-manifest.json`] = Buffer.from('{"changed":true}');
  await assert.rejects(fetchVendor({ repo: other.root, fetchImpl: other.fetchImpl, log: () => {} }), /kit-manifest\.json is not the pinned sha256/);
  assert.equal(vendorProblems(other.root).filter(problem => problem.includes('missing')).length, 3, 'nothing was written');
});

test('an npm tarball that is not the pinned one, or a file in it that is not, changes nothing on disk', async t => {
  const { root, fetchImpl, served } = repo(t);
  await fetchVendor({ repo: root, fetchImpl, log: () => {} });
  const before = readFileSync(path.join(root, TO.manifest), 'utf8');
  rmSync(path.join(root, '.vendor-cache'), { recursive: true, force: true });
  served['https://registry.example/@scope/editor/-/editor-1.0.0.tgz'] = Buffer.from('other');
  await assert.rejects(fetchVendor({ repo: root, fetchImpl, log: () => {} }), /sha512 is not the pinned/);
  assert.equal(readFileSync(path.join(root, TO.manifest), 'utf8'), before);
  const bad = repo(t, { tamper: pin => { pin.files['sheet-host-fonts.css'].sha256 = '0'.repeat(64); } });
  await assert.rejects(fetchVendor({ repo: bad.root, fetchImpl: bad.fetchImpl, log: () => {} }), /not the pinned sha256/);
  assert.equal(vendorProblems(bad.root).filter(problem => problem.includes('missing')).length, 3, 'nothing was written for the bad pin');
});

test('a tampered cache is not used: the source is downloaded again and the cache rewritten', async t => {
  const { root, fetchImpl, asked } = repo(t);
  await fetchVendor({ repo: root, fetchImpl, log: () => {} });
  const cache = path.join(root, '.vendor-cache');
  writeFileSync(path.join(cache, 'editor-1.0.0.tgz'), 'tampered');
  writeFileSync(path.join(cache, 'release-kit-manifest.json'), 'tampered');
  writeFileSync(path.join(cache, 'release-SHA256SUMS'), `${'0'.repeat(64)}  kit-manifest.json\n`);
  asked.length = 0;
  await fetchVendor({ repo: root, fetchImpl, log: () => {} });
  assert.deepEqual(asked.sort(), [`${RELEASE}SHA256SUMS`, `${RELEASE}kit-manifest.json`, 'https://registry.example/@scope/editor/-/editor-1.0.0.tgz'].sort());
  assert.deepEqual(vendorProblems(root), []);
  assert.notEqual(readFileSync(path.join(cache, 'editor-1.0.0.tgz'), 'utf8'), 'tampered', 'the cache holds the pinned tarball again');
});

test('OFFICESDK_VENDOR_DIR: local copies named as in the cache, still checked', async t => {
  const first = repo(t);
  await fetchVendor({ repo: first.root, fetchImpl: first.fetchImpl, log: () => {} });
  const dir = path.join(first.root, '.vendor-cache');
  writeFileSync(path.join(dir, 'release-SHA256SUMS'), first.served[`${RELEASE}SHA256SUMS`]);
  const offline = repo(t);
  const written = await fetchVendor({ repo: offline.root, dir, fetchImpl: async () => { throw new Error('no network'); } });
  assert.equal(written.length, 3);
  assert.deepEqual(vendorProblems(offline.root), []);
});

test('check: runnable platforms only, and the editor bundle the manifest names', async t => {
  const { root, fetchImpl } = repo(t);
  await fetchVendor({ repo: root, fetchImpl, log: () => {} });
  const pin = JSON.parse(readFileSync(path.join(root, 'components/office-kit.json'), 'utf8'));
  pin.officeKit.platforms = ['darwin-arm64', 'openharmony-arm64'];
  writeFileSync(path.join(root, 'components/office-kit.json'), JSON.stringify(pin));
  assert.deepEqual(vendorProblems(root), ['kit-manifest.json has runnable Kits for darwin-arm64, the pin lists darwin-arm64, openharmony-arm64']);
  const other = JSON.stringify({ ...JSON.parse(MANIFEST), editor: { bundle: { sha256: `sha256-${createHash('sha256').update('another').digest('base64')}` } } });
  const mismatch = repo(t, { manifest: other });
  await fetchVendor({ repo: mismatch.root, fetchImpl: mismatch.fetchImpl, log: () => {} });
  assert.deepEqual(vendorProblems(mismatch.root), ['components/vendor/editor/bundle/offline-editor.js is not the editor bundle kit-manifest.json names (editor.bundle.sha256)']);
});

test('the pin itself: v2, https sources, sha512 integrity, known sources, safe paths, targets under components/', () => {
  assert.deepEqual(pinProblems({ v: 1 }), ['components/office-kit.json: v must be 2']);
  // Paths stay inside their root on every platform; source ids are plain.
  const escape = { v: 2, officeKit: { version: '1', platforms: [] }, sources: { 'Bad/Id': { url: 'https://x/y/' } },
    files: { a: { from: 'Bad/Id', path: '..\\a', to: 'components/x', sha256: '0'.repeat(64) }, b: { from: 'Bad/Id', path: 'C:/b', to: 'components/..\\b', sha256: '0'.repeat(64) }, c: { from: 'Bad/Id', path: 'a//b', to: 'components/c', sha256: '0'.repeat(64) } } };
  assert.deepEqual(pinProblems(escape), ['components/office-kit.json: sources.Bad/Id: an id is lower-case letters, digits and -',
    'components/office-kit.json: files.a.path must be a relative path in its source', 'components/office-kit.json: files.b.path must be a relative path in its source',
    'components/office-kit.json: files.b.to must stay under components/', 'components/office-kit.json: files.c.path must be a relative path in its source']);
  const pin = { v: 2, officeKit: { version: '1', commit: 3, platforms: [] },
    sources: { a: { registry: 'http://x/', package: 'p', version: '1', integrity: 'sha1-x' }, b: { url: 'http://x/y/', sums: '../S' } },
    files: { f: { from: 'c', path: '../f', to: '../a', sha256: 'zz' }, g: { from: 'a', path: 'g', to: 'components/x', sha256: '0'.repeat(64) }, h: { from: 'a', path: 'h', to: 'components/x', sha256: '0'.repeat(64) } } };
  assert.deepEqual(pinProblems(pin), [
    'components/office-kit.json: sources.a.registry must be https', 'components/office-kit.json: sources.a.integrity must be sha512',
    'components/office-kit.json: sources.b.url must be an https directory (ending in /)', 'components/office-kit.json: sources.b.sums must be a file name in that directory',
    'components/office-kit.json: files.f.from must name one of sources', 'components/office-kit.json: files.f.path must be a relative path in its source',
    'components/office-kit.json: files.f.sha256 must be 64 hex digits', 'components/office-kit.json: files.f.to must stay under components/',
    'components/office-kit.json: files.h.to is written twice', 'components/office-kit.json: officeKit.commit must be a string when present']);
});

test('the committed pin is valid', () => {
  assert.deepEqual(pinProblems(JSON.parse(readFileSync(new URL('../../../components/office-kit.json', import.meta.url), 'utf8'))), []);
});
