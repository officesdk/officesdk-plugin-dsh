import test from 'node:test';
import assert from 'node:assert/strict';
import { cpSync, mkdirSync, mkdtempSync, readFileSync, realpathSync, rmSync, symlinkSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { withFileLock, writeFileAtomic } from '@deepseek-ai/dsh-atomic-write';
import ShimoPreview, { styleNames } from '../src/host/index.mjs';
import { SuiteError } from '../../core/src/index.mjs';
import { StyleMemory, THUMB_MAX_BYTES, localized, readStyleDefaults, styleCard, styleCards } from '../src/host/style-catalog.mjs';

const STYLES = fileURLToPath(new URL('../skills/shimo-pptx/styles/', import.meta.url));
const PNG = Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 1, 2, 3]);
const atomic = { withFileLock, writeFileAtomic };
const signal = () => new AbortController().signal;

/**
 * The real packs plus a made-up second one (`second`, dark, with meta and a thumbnail); test-only, never shipped. Its
 * meta follows the style distill toolkit's settled fields (its spec §1.1).
 */
export function withSecondPack(t, { meta, thumb = PNG } = {}) {
  const root = realpathSync(mkdtempSync(join(tmpdir(), 'officesdk-picker-')));
  t.after(() => rmSync(root, { recursive: true, force: true }));
  const styles = join(root, 'styles');
  cpSync(join(STYLES, 'default'), join(styles, 'default'), { recursive: true });   // only default: the tests count the packs
  cpSync(join(styles, 'default'), join(styles, 'second'), { recursive: true });
  rmSync(join(styles, 'second', 'thumb.png'), { force: true });
  const style = JSON.parse(readFileSync(join(styles, 'second', 'style.json'), 'utf8'));
  style.name = 'second';
  style.title = '测试黑底';
  style.default_palette = 'color-2';
  style.meta = meta ?? {
    id: 'second', name: { zh: '黑底画报', en: 'Dark editorial' }, aliases: ['黑底', ' photo-noir ', '黑底', 7],
    summary: { zh: '黑底大图、白字标题的摄影画报', en: 'Black pages, big photos, white titles' },
    scenarios: { zh: ['城市', '旅行', '品牌故事'], en: [] },
    requires: { photos: { slots: 19, required: 12, per_slide_max: 3 }, note: { zh: '每份约 19 张照片', en: 'About 19 photos per deck' } },
    thumb: 'thumb.png', fonts: [], design_source: { from: 'test', note: '', cleared: true },
  };
  writeFileSync(join(styles, 'second', 'style.json'), JSON.stringify(style));
  if (thumb) writeFileSync(join(styles, 'second', 'thumb.png'), thumb);
  return { root, styles, style };
}

test('localized text falls back language to language, then to the fallback', () => {
  assert.deepEqual(localized('松烟'), { zh: '松烟', en: '松烟' });
  assert.deepEqual(localized({ en: 'Ink' }), { zh: 'Ink', en: 'Ink' });
  assert.deepEqual(localized({ zh: ' ', en: '' }, 'x'), { zh: 'x', en: 'x' });
  assert.deepEqual(localized(undefined), { zh: '', en: '' });
});

test('a pack card: meta fields, aliases cleaned, primary colour, thumbnail as a data URI', t => {
  const { styles, style } = withSecondPack(t);
  const card = styleCard(styles, 'second');
  assert.deepEqual(card.name, { zh: '黑底画报', en: 'Dark editorial' });
  assert.deepEqual(card.aliases, ['黑底', 'photo-noir']);
  assert.deepEqual(card.summary, { zh: '黑底大图、白字标题的摄影画报', en: 'Black pages, big photos, white titles' });
  assert.deepEqual(card.scenarios, { zh: ['城市', '旅行', '品牌故事'], en: ['城市', '旅行', '品牌故事'] }, 'a missing language borrows the other');
  assert.deepEqual(card.requires, { note: { zh: '每份约 19 张照片', en: 'About 19 photos per deck' }, photos: { slots: 19, required: 12, perSlideMax: 3 } });
  assert.equal(card.primary, style.palettes['color-2'].primary);
  assert.equal(card.thumb, `data:image/png;base64,${PNG.toString('base64')}`);
});

test('a pack without meta falls back to its style.json title, then its id; a bad or oversized thumbnail is left out', t => {
  const { styles } = withSecondPack(t, { meta: 'nonsense', thumb: Buffer.from('GIF89a') });
  const card = styleCard(styles, 'second');
  assert.deepEqual(card.name, { zh: '测试黑底', en: '测试黑底' });
  assert.deepEqual([card.summary, card.scenarios, card.requires], [{ zh: '', en: '' }, { zh: [], en: [] }, { note: { zh: '', en: '' }, photos: null }]);
  assert.deepEqual(card.aliases, []);
  assert.equal(card.thumb, null);
  writeFileSync(join(styles, 'second', 'thumb.png'), Buffer.concat([PNG, Buffer.alloc(THUMB_MAX_BYTES)]));
  assert.equal(styleCard(styles, 'second').thumb, null);
  const style = JSON.parse(readFileSync(join(styles, 'second', 'style.json'), 'utf8'));
  delete style.title; delete style.meta;
  writeFileSync(join(styles, 'second', 'style.json'), JSON.stringify(style));
  assert.deepEqual(styleCard(styles, 'second').name, { zh: 'second', en: 'second' });
});

test('meta.thumb names a PNG inside the pack; a path that leaves the pack is not followed', t => {
  const { root, styles } = withSecondPack(t);
  const style = JSON.parse(readFileSync(join(styles, 'second', 'style.json'), 'utf8'));
  writeFileSync(join(styles, 'second', 'cover.png'), Buffer.concat([PNG, Buffer.from([9])]));
  writeFileSync(join(root, 'outside.png'), Buffer.concat([PNG, Buffer.from([8])]));
  const withThumb = thumb => { writeFileSync(join(styles, 'second', 'style.json'), JSON.stringify({ ...style, meta: { ...style.meta, thumb } })); return styleCard(styles, 'second').thumb; };
  assert.equal(withThumb('cover.png'), `data:image/png;base64,${Buffer.concat([PNG, Buffer.from([9])]).toString('base64')}`);
  for (const bad of ['../outside.png', '/etc/x.png', 'sub/cover.png', '.hidden.png', 'cover.jpg']) assert.equal(withThumb(bad), `data:image/png;base64,${PNG.toString('base64')}`, `${bad}: falls back to thumb.png`);
});

test('a pack that cannot be read is skipped; the others are still listed', t => {
  const { styles } = withSecondPack(t);
  writeFileSync(join(styles, 'second', 'style.json'), '{ broken');
  const warnings = [];
  assert.deepEqual(styleCards(styles, styleNames(styles), text => warnings.push(text)).map(card => card.id), ['default']);
  assert.match(warnings[0], /style second skipped/);
});

test('style memory: per workspace, the longest remembered root covers folders inside it, forgetting removes it, the file survives a restart', async t => {
  const root = realpathSync(mkdtempSync(join(tmpdir(), 'officesdk-memory-')));
  t.after(() => rmSync(root, { recursive: true, force: true }));
  const ws = join(root, 'ws'), inner = join(ws, 'inner'), other = join(root, 'ws-other');
  for (const dir of [join(inner, 'deep'), other]) mkdirSync(dir, { recursive: true });
  symlinkSync(ws, join(root, 'link'));
  const file = join(root, 'data', 'style-defaults.json');
  const memory = new StyleMemory({ file, atomic, now: () => 't0' });
  assert.equal(memory.rememberedFor(ws), undefined);
  await memory.set(ws, 'second');
  await memory.set(inner, 'default');
  assert.equal(memory.rememberedFor(ws), 'second');
  assert.equal(memory.rememberedFor(join(ws, 'sub')), 'second', 'a folder inside the workspace');
  assert.equal(memory.rememberedFor(join(inner, 'deep')), 'default', 'the nearer root wins');
  assert.equal(memory.rememberedFor(join(root, 'link')), 'second', 'a symlinked path is the same workspace');
  assert.equal(memory.rememberedFor(other), undefined, 'a sibling whose name only starts the same');
  assert.equal(memory.rememberedFor(''), undefined);
  assert.deepEqual(await readStyleDefaults(file), { v: 1, workspaces: { [ws]: { style: 'second', updated: 't0' }, [inner]: { style: 'default', updated: 't0' } } });
  await memory.set(inner, null);
  const again = new StyleMemory({ file, atomic });
  assert.equal(again.rememberedFor(join(inner, 'deep')), 'second');
  // A second writer (another Host process) is not overwritten: each write re-reads the file under the lock.
  await new StyleMemory({ file, atomic }).set(other, 'default');
  await again.set(ws, null);
  assert.deepEqual(Object.keys((await readStyleDefaults(file)).workspaces), [other]);
});

test('style memory: an unreadable file is an empty memory, and the next write replaces it', async t => {
  const root = realpathSync(mkdtempSync(join(tmpdir(), 'officesdk-memory-')));
  t.after(() => rmSync(root, { recursive: true, force: true }));
  const file = join(root, 'style-defaults.json');
  writeFileSync(file, '{ nope');
  const warnings = [];
  const memory = new StyleMemory({ file, atomic, warn: text => warnings.push(text) });
  assert.equal(memory.rememberedFor(root), undefined);
  assert.match(warnings[0], /unreadable/);
  await memory.set(root, 'default');
  assert.equal((await readStyleDefaults(file)).workspaces[root].style, 'default');
});

/** ShimoPreview's style methods over a plain object: the Remote peer and authorization are stand-ins. */
function picker(t, { principal = 'owner', authorize } = {}) {
  const { root, styles } = withSecondPack(t);
  const ws = join(root, 'ws');
  mkdirSync(join(ws, 'docs'), { recursive: true });
  const authorized = [];
  const self = Object.create(ShimoPreview.prototype);
  Object.assign(self, {
    ctx: {}, styleRoot: styles, styleCache: undefined,
    styleMemory: new StyleMemory({ file: join(root, 'data', 'style-defaults.json'), atomic }),
    remotePeer: () => ({ principal }),
    suite: { authorize: async (_peer, scope, _signal, action) => { authorized.push([scope.sessionId, action]); await authorize?.(); } },
  });
  return { self, root, styles, ws, scope: { sessionId: 's1', workspaceRoot: ws }, authorized };
}

test('listStyles: every pack\'s card, the workspace default and whether the caller may remember one', async t => {
  const { self, scope, authorized } = picker(t);
  const listed = await self.listStyles(scope, signal());
  assert.deepEqual(listed.styles.map(card => card.id), ['default', 'second']);
  assert.equal(listed.styles[1].name.zh, '黑底画报');
  assert.match(listed.styles[1].thumb, /^data:image\/png;base64,/);
  assert.deepEqual([listed.defaultStyle, listed.remembered, listed.canRemember], ['default', null, true]);
  assert.deepEqual(authorized, [['s1', 'open']]);
  const guest = picker(t, { principal: 'share:g' });
  assert.equal((await guest.self.listStyles(guest.scope, signal())).canRemember, false);
  const denied = picker(t, { authorize: async () => { throw new SuiteError('forbidden', 'Access to the document was denied.'); } });
  await assert.rejects(denied.self.listStyles(denied.scope, signal()), /denied/);
});

test('rememberStyle: owner only, known packs only; null forgets; a deleted pack is no default any more', async t => {
  const { self, scope, styles, ws } = picker(t);
  assert.deepEqual(await self.rememberStyle(scope, 'second', signal()), { defaultStyle: 'second', remembered: 'second', rememberedFrom: null });
  assert.deepEqual(self.defaultStyleFor(join(ws, 'docs')), { style: 'second', source: 'workspace' });
  await assert.rejects(self.rememberStyle(scope, 'nope', signal()), error => error.code === 'officesdk/invalid-edit' || /Unknown style "nope"/.test(error.message));
  await assert.rejects(self.rememberStyle(scope, '../default', signal()), /Unknown style/);
  const guest = picker(t, { principal: 'share:g' });
  await assert.rejects(guest.self.rememberStyle(guest.scope, 'second', signal()), /Only the owner/);
  rmSync(join(styles, 'second'), { recursive: true });
  self.styleCache = undefined;
  assert.deepEqual(self.defaultStyleFor(ws), { style: 'default', source: 'default' });
  const stale = await self.listStyles(scope, signal());
  assert.deepEqual([stale.defaultStyle, stale.remembered], ['default', 'second'], 'the stale choice is still reported');
  assert.deepEqual(await self.rememberStyle(scope, null, signal()), { defaultStyle: 'default', remembered: null, rememberedFrom: null });
});

test('rememberStyle: a write that fails is reported as unavailable, the memory unchanged', async t => {
  const { self, scope } = picker(t);
  self.styleMemory.atomic = { withFileLock: async () => { throw Object.assign(new Error('disk full'), { code: 'ENOSPC' }); } };
  await assert.rejects(self.rememberStyle(scope, 'second', signal()), /could not be saved \(ENOSPC\)/);
  assert.equal(self.styleMemory.rememberedFor(scope.workspaceRoot), undefined);
});

test('a workspace inside a remembered folder inherits its choice and says so; its own choice wins and is its own to clear', async t => {
  const { self, ws } = picker(t);
  const inner = join(ws, 'docs');
  const scope = { sessionId: 's1', workspaceRoot: inner };
  await self.styleMemory.set(ws, 'second');
  assert.deepEqual((await self.listStyles(scope, signal())).rememberedFrom, ws);
  assert.deepEqual(await self.rememberStyle(scope, 'default', signal()), { defaultStyle: 'default', remembered: 'default', rememberedFrom: null });
  assert.deepEqual(await self.rememberStyle(scope, null, signal()), { defaultStyle: 'second', remembered: 'second', rememberedFrom: ws }, 'clearing its own choice falls back to the enclosing one');
});

test('runtime context reads only the name, never the thumbnail', async t => {
  const { self, scope, ws, styles } = picker(t);
  await self.rememberStyle(scope, 'second', signal());
  writeFileSync(join(styles, 'second', 'thumb.png'), Buffer.concat([PNG, Buffer.alloc(1024)]));
  const card = styleCard(styles, 'second', { thumb: false });
  assert.equal(card.thumb, null);
  assert.match(self.styleContext(ws), /黑底画报/);
});

test('runtime context: names the remembered style for any folder of that workspace, nothing for default or none', async t => {
  const { self, scope, ws, styles } = picker(t);
  assert.equal(self.styleContext(ws), '');
  await self.rememberStyle(scope, 'default', signal());
  assert.equal(self.styleContext(ws), '', 'default needs no reminder');
  await self.rememberStyle(scope, 'second', signal());
  const text = self.styleContext(join(ws, 'docs'));
  assert.match(text, /^Workspace default PPT style \(shimo-pptx\): second \(黑底画报 \/ Dark editorial\)\./);
  assert.match(text, /leave style out of presentation_create/);
  assert.match(text, /--style second/);
  assert.equal(self.styleContext(undefined), '');
  assert.equal(self.styleContext(tmpdir()), '');
  rmSync(join(styles, 'second'), { recursive: true });
  self.styleCache = undefined;
  assert.equal(self.styleContext(ws), '', 'a deleted pack is not offered to the model');
});

test('the shipped packs\' cards speak to the user: the photo note names no script or file', () => {
  for (const card of styleCards(STYLES, styleNames(STYLES))) {
    for (const lang of ['zh', 'en']) {
      const note = card.requires.note[lang];
      assert.ok(note, `${card.id}: a ${lang} note`);
      assert.doesNotMatch(note, /\.py\b|\.json\b|scripts\//, `${card.id} ${lang}: ${note}`);
    }
  }
  assert.equal(styleCards(STYLES, ['default'])[0].requires.note.zh, '照片可选，只有图像陈述页需要一张；缺图时自动生成配图');
});
