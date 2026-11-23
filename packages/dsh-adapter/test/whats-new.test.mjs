// 「本次更新」: CHANGELOG highlights → whats-new.json → Host notes → the card. No DSH packages: CI runs this file.
import test from 'node:test';
import assert from 'node:assert/strict';
import { mkdir, mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { dirname, join } from 'node:path';
import { tmpdir } from 'node:os';
import { fileURLToPath } from 'node:url';
import { parseWhatsNew, renderWhatsNew, missingEnglish, MAX_HIGHLIGHTS } from '../../../scripts/whats-new.mjs';
import { WhatsNew, compareVersions, notesBetween, readNotes, releaseOf } from '../src/host/whats-new.mjs';
import { WHATS_NEW_SEEN_MS, WHATS_NEW_VISIBLE, createWhatsNewWatch, fill, notesLanguage, whatsNewView } from '../src/client/whats-new.mjs';

const adapter = fileURLToPath(new URL('..', import.meta.url));
const tick = () => new Promise(resolve => setImmediate(resolve));

const CHANGELOG = `# 更新日志

说明。

## 未发布

### 更新要点

- 新功能 \`x\` 和 **重点**。

### 详细

- 内部细节，不进提示。

## 0.0.3 — 2026-11-03

只有详细条目的版本。

- 没有要点。

## 0.0.2 — 2026-11-02

导语。

### 更新要点

- 第一条
  接着写的第二行。

- 第二条（隔一个空行仍是同一个列表）

一段正文结束要点。

- 详细条目，不算要点。
`;

test('the committed whats-new.json is what CHANGELOG.md makes (run node scripts/whats-new.mjs after editing highlights)', async () => {
  const changelog = await readFile(join(adapter, 'CHANGELOG.md'), 'utf8');
  const committed = await readFile(join(adapter, 'whats-new.json'), 'utf8');
  assert.equal(committed, renderWhatsNew(parseWhatsNew(changelog)), 'whats-new.json is out of date: run node scripts/whats-new.mjs');
  const notes = JSON.parse(committed);
  for (const entry of notes.entries) assert.ok(entry.highlights.length >= 1 && entry.highlights.length <= MAX_HIGHLIGHTS, `${entry.version}: 1–${MAX_HIGHLIGHTS} points`);
  const manifest = JSON.parse(await readFile(join(adapter, 'package.json'), 'utf8'));
  assert.ok(manifest.files.includes('whats-new.json'), 'the notes ship in every package');
});

test('the open-source CHANGELOG has one initial section, 0.0.1-rc.1; updates from 0.0.0-alpha.3 and new installs show it', async () => {
  const notes = JSON.parse(await readFile(join(adapter, 'whats-new.json'), 'utf8'));
  assert.deepEqual(notes.entries.map(entry => entry.version), ['0.0.1-rc.1']);
  assert.equal(notes.latestReleased, '0.0.1-rc.1', 'the dated section is the latest release');
  const { highlights: points, highlightsEn: pointsEn } = notes.entries[0];
  assert.equal(points.length, 2);
  assert.equal(pointsEn.length, 2);
  assert.ok(points[0].startsWith('初始开源版本'), 'the entry says it is the initial open-source version');
  const shipped = readNotes(join(adapter, 'whats-new.json'));
  const section = [{ version: '0.0.1-rc.1', highlights: points, highlightsEn: pointsEn }];
  assert.deepEqual(notesBetween(shipped, undefined, '0.0.1-rc.1'), section, 'a new install shows the points');
  assert.deepEqual(notesBetween(shipped, '0.0.0-alpha.3', '0.0.1-rc.1'), section, 'an update from 0.0.0-alpha.3 shows them once');
  assert.deepEqual(notesBetween(shipped, undefined, '0.0.0-alpha.4'), [], 'an older build has no section of its own');
  assert.deepEqual(notesBetween({ latestReleased: '0.0.0-alpha.3', entries: [{ version: 'unreleased', highlights: ['n'] }, { version: '0.0.0-alpha.3', highlights: ['a3'] }] }, '0.0.0-alpha.3', '0.0.0-alpha.4'), [{ version: '0.0.0-alpha.4', highlights: ['n'] }], 'a later alpha takes the 未发布 points');
  assert.deepEqual(notesBetween({ latestReleased: '0.0.0-alpha.3', entries: [{ version: '0.0.0-beta.1', highlights: ['b'] }, { version: '0.0.0-alpha.4', highlights: ['a4'] }] }, '0.0.0-alpha.3', '0.0.0-beta.1').map(entry => entry.version), ['0.0.0-beta.1', '0.0.0-alpha.4'], 'previews are sections of their own, in stage order');
  const changelog = await readFile(join(adapter, 'CHANGELOG.md'), 'utf8');
  assert.equal(changelog.match(/^## /gm).length, 1, 'history is not kept');
});

test('highlights come from ### 更新要点 only, as plain text; versions without them are left out', () => {
  assert.deepEqual(parseWhatsNew(CHANGELOG), {
    latestReleased: '0.0.3',
    entries: [
      { version: 'unreleased', highlights: ['新功能 x 和 重点。'] },
      { version: '0.0.2', highlights: ['第一条 接着写的第二行。', '第二条（隔一个空行仍是同一个列表）'] },
    ],
  });
});

test('a highlights list must have one to five points', () => {
  const many = Array.from({ length: MAX_HIGHLIGHTS + 1 }, (_, index) => `- 第 ${index} 条`).join('\n');
  assert.throws(() => parseWhatsNew(`## 0.0.1 — d\n\n### 更新要点\n\n${many}\n`), /at most 5/);
  assert.throws(() => parseWhatsNew('## 0.0.1 — d\n\n### 更新要点\n\n### 详细\n'), /no points/);
});

test('versions compare by X.Y.Z, whatever the test or Windows suffix', () => {
  assert.equal(releaseOf('0.0.47-test.12'), '0.0.47');
  assert.equal(releaseOf('0.0.47-win32'), '0.0.47');
  assert.equal(releaseOf('0.0.47-win32-test.3'), '0.0.47');
  assert.equal(releaseOf('unreleased'), undefined);
  assert.equal(compareVersions('0.0.47-win32', '0.0.47'), 0);
  assert.ok(compareVersions('0.0.10', '0.0.9') > 0);
  assert.ok(Number.isNaN(compareVersions('x', '0.0.1')));
});

const NOTES = { latestReleased: '0.0.46', entries: [
  { version: 'unreleased', highlights: ['u1'] },
  { version: '0.0.46', highlights: ['a', 'b'] },
  { version: '0.0.44', highlights: ['c'] },
  { version: '0.0.40', highlights: ['old'] },
] };

test('an update shows every version above the last one shown, newest first; 未发布 counts as the newer package', () => {
  assert.deepEqual(notesBetween(NOTES, '0.0.43', '0.0.47-test.5'), [
    { version: '0.0.47', highlights: ['u1'] }, { version: '0.0.46', highlights: ['a', 'b'] }, { version: '0.0.44', highlights: ['c'] }]);
  assert.deepEqual(notesBetween(NOTES, '0.0.44', '0.0.46'), [{ version: '0.0.46', highlights: ['a', 'b'] }], '未发布 is not the 0.0.46 package');
  assert.deepEqual(notesBetween(NOTES, undefined, '0.0.47-win32'), [{ version: '0.0.47', highlights: ['u1'] }], 'no record: only this version');
  assert.deepEqual(notesBetween(NOTES, undefined, '0.0.45'), [], 'no record and no points of its own');
  // 0.0.48 released without points while 未发布 already holds 0.0.49's: a 0.0.48 package must not show them.
  assert.deepEqual(notesBetween({ ...NOTES, latestReleased: '0.0.48' }, '0.0.47', '0.0.48'), []);
});

test('notes that are missing or malformed are no notes', async () => {
  const dir = await mkdtemp(join(tmpdir(), 'whats-new-'));
  try {
    assert.deepEqual(readNotes(join(dir, 'none.json')), { entries: [] });
    await writeFile(join(dir, 'bad.json'), '{"v":2}');
    assert.deepEqual(readNotes(join(dir, 'bad.json')), { entries: [] });
    await writeFile(join(dir, 'mixed.json'), JSON.stringify({ v: 1, latestReleased: '0.0.2', entries: [{ version: '0.0.2', highlights: ['ok'] }, { version: '0.0.1', highlights: [] }, { version: 'x', highlights: ['?'] }] }));
    assert.deepEqual(readNotes(join(dir, 'mixed.json')), { latestReleased: '0.0.2', entries: [{ version: '0.0.2', highlights: ['ok'] }] });
  } finally { await rm(dir, { recursive: true, force: true }); }
});

/** A plugin directory at `version` with NOTES, and the Host's record under its own data directory. */
async function plugin(version, notes = NOTES) {
  const dir = await mkdtemp(join(tmpdir(), 'whats-new-host-'));
  await writeFile(join(dir, 'package.json'), JSON.stringify({ version }));
  if (notes) await writeFile(join(dir, 'whats-new.json'), JSON.stringify({ v: 1, ...notes }));
  const file = join(dir, 'data', 'plugins', 'officesdk', 'whats-new-seen.json');
  const locks = [];
  const atomic = { withFileLock: async (lock, fn) => { locks.push(lock); return fn(); }, writeFileAtomic: (target, text) => writeFile(target, text) };
  const record = async () => JSON.parse(await readFile(file, 'utf8'));
  const seen = async version => { await mkdir(dirname(file), { recursive: true }); await writeFile(file, JSON.stringify({ v: 1, version })); };
  return { dir, file, locks, record, seen, whatsNew: new WhatsNew({ packageRoot: dir, file, atomic, now: () => 'T' }), done: () => rm(dir, { recursive: true, force: true }) };
}

test('after an update the Host offers the notes until the card is shown, then never again for that version', async () => {
  const p = await plugin('0.0.47-test.3');
  try {
    await p.seen('0.0.45');
    const pending = await p.whatsNew.pending();
    assert.deepEqual(pending, { current: '0.0.47', previous: '0.0.45', entries: [{ version: '0.0.47', highlights: ['u1'] }, { version: '0.0.46', highlights: ['a', 'b'] }] });
    assert.deepEqual(await p.whatsNew.pending(), pending, 'asking does not record it: a card that never appeared asks again');
    await p.whatsNew.markSeen('0.0.47');
    assert.deepEqual(await p.record(), { v: 1, version: '0.0.47-test.3', at: 'T' });
    assert.equal(p.locks.length, 1);
    assert.equal(await p.whatsNew.pending(), null);
    await assert.rejects(p.whatsNew.markSeen('0.0.46'), /not this plugin's version/);
  } finally { await p.done(); }
});

test('no record (a new install or one from before this feature): only this version, under a neutral title', async () => {
  const p = await plugin('0.0.47');
  try {
    const pending = await p.whatsNew.pending();
    assert.deepEqual(pending, { current: '0.0.47', entries: [{ version: '0.0.47', highlights: ['u1'] }] });
    assert.equal(whatsNewView(pending).title, 'whatsNewFresh');
  } finally { await p.done(); }
});

test('nothing to show (same version, a downgrade, no points, no notes file) records this version quietly', async () => {
  for (const [version, seen, notes] of [['0.0.47-win32', '0.0.47', NOTES], ['0.0.46', '0.0.47', NOTES], ['0.0.45', '0.0.44', NOTES], ['0.0.47', '0.0.40', null]]) {
    const p = await plugin(version, notes);
    try {
      await p.seen(seen);
      assert.equal(await p.whatsNew.pending(), null, `${seen} → ${version}`);
      assert.equal((await p.record()).version, seen === '0.0.47' && version === '0.0.47-win32' ? '0.0.47' : version, 'the same version keeps its record');
    } finally { await p.done(); }
  }
});

test('a broken record counts as none', async () => {
  const p = await plugin('0.0.47');
  try {
    await p.seen('0.0.46');
    await writeFile(p.file, '{not json');
    assert.deepEqual((await p.whatsNew.pending()).entries.map(entry => entry.version), ['0.0.47']);
  } finally { await p.done(); }
});

test('the card asks once for every tab, says nothing when the Host cannot answer, and marks the version once', async () => {
  const asked = [];
  const marked = [];
  let answer;
  const timers = [];
  const scheduler = { setTimeout: (fn, ms) => { timers.push({ fn, ms }); return timers.length; }, clearTimeout: id => { if (timers[id - 1]) timers[id - 1].cleared = true; } };
  const watch = createWhatsNewWatch({ scheduler, fetch: () => { asked.push(1); return answer; }, mark: version => { marked.push(version); } });
  let renders = 0;
  watch.subscribe(() => renders++);
  await watch.check();
  assert.equal(watch.snapshot(), undefined, 'no tab could ask yet');
  answer = Promise.resolve({ current: '0.0.47', entries: [{ version: '0.0.47', highlights: ['u1'] }] });
  await Promise.all([watch.check(), watch.check()]);
  assert.equal(asked.length, 2, 'a second tab shares the answer');
  assert.equal(watch.snapshot().current, '0.0.47');
  await watch.check();
  assert.equal(asked.length, 2, 'asked once per load');
  const goneA = watch.shown();
  goneA();
  await tick();
  assert.deepEqual(marked, [], 'a card that only flashed past is not seen');
  assert.equal(timers[0].cleared, true, 'its clock stops with it');
  watch.shown();
  const goneB = watch.shown();
  goneB();
  assert.deepEqual(timers.filter(timer => !timer.cleared).map(timer => timer.ms), [WHATS_NEW_SEEN_MS], 'one clock while any tab shows it');
  timers.at(-1).fn();
  await tick();
  assert.deepEqual(marked, ['0.0.47'], 'on screen long enough: seen');
  watch.dismiss();
  await tick();
  assert.deepEqual(marked, ['0.0.47'], 'marked once');
  assert.equal(watch.snapshot(), undefined, 'closed in every tab');

  const closing = createWhatsNewWatch({ scheduler, fetch: () => ({ current: '0.0.48', entries: [{ version: '0.0.48', highlights: ['x'] }] }), mark: version => { marked.push(version); } });
  await closing.check();
  closing.shown();
  closing.dismiss();
  await tick();
  assert.deepEqual(marked, ['0.0.47', '0.0.48'], 'closing the card marks it at once');
  assert.equal(timers.at(-1).cleared, true, 'and stops the clock');
  assert.ok(renders >= 2);

  let fails = 0;
  const failing = createWhatsNewWatch({ fetch: () => { fails++; return Promise.reject(new Error('unknown method whatsNew')); }, mark: () => {} });
  for (let i = 0; i < 5; i++) await failing.check();
  assert.equal(failing.snapshot(), undefined, 'an older Host shows nothing');
  assert.equal(fails, 3, 'a failed ask is tried again by later tabs, three times in all');
  let flaky = 0;
  const recovering = createWhatsNewWatch({ fetch: () => (++flaky === 1 ? Promise.reject(new Error('session starting')) : { current: '0.0.47', entries: [{ version: '0.0.47', highlights: ['u1'] }] }), mark: () => {} });
  await recovering.check();
  await recovering.check();
  assert.equal(recovering.snapshot().current, '0.0.47', 'a passing failure does not lose the card');
  const empty = createWhatsNewWatch({ fetch: () => null, mark: () => {} });
  await empty.check();
  assert.equal(empty.snapshot(), undefined);
});

test('the card groups several versions and shows four points until expanded', () => {
  const notes = { current: '0.0.47', previous: '0.0.44', entries: [
    { version: '0.0.47', highlights: ['1', '2', '3'] }, { version: '0.0.46', highlights: ['4', '5'] }, { version: '0.0.45', highlights: ['6'] }] };
  const short = whatsNewView(notes);
  assert.equal(WHATS_NEW_VISIBLE, 4);
  assert.deepEqual(short.groups, [{ version: '0.0.47', highlights: ['1', '2', '3'] }, { version: '0.0.46', highlights: ['4'] }]);
  assert.deepEqual({ title: short.title, grouped: short.grouped, total: short.total, hidden: short.hidden }, { title: 'whatsNewRange', grouped: true, total: 6, hidden: 2 });
  const all = whatsNewView(notes, true);
  assert.equal(all.hidden, 0);
  assert.equal(all.groups.length, 3);
  const one = whatsNewView({ current: '0.0.47', previous: '0.0.44', entries: [{ version: '0.0.47', highlights: ['1'] }] });
  assert.deepEqual([one.title, one.grouped, one.hidden], ['whatsNewRange', false, 0], 'the starting version shows even when only the newest has points');
  assert.equal(fill('已从 v{previous} 更新到 v{current}', short.values), '已从 v0.0.44 更新到 v0.0.47');
  assert.equal(fill('展开全部（共 {n} 条）', { n: 6 }), '展开全部（共 6 条）');
});

test('0.0.0-alpha.1 installed over a git release (0.0.50) shows its own points; a real downgrade within a line still shows nothing (review R1-3)', async () => {
  const notes = { latestReleased: '0.0.0-alpha.1', entries: [{ version: '0.0.0-alpha.1', highlights: ['new'] }, { version: '0.0.50', highlights: ['old'] }] };
  const p = await plugin('0.0.0-alpha.1', notes);
  try {
    await p.seen('0.0.50');
    assert.deepEqual(await p.whatsNew.pending(), { current: '0.0.0-alpha.1', entries: [{ version: '0.0.0-alpha.1', highlights: ['new'] }] });
  } finally { await p.done(); }
  const back = await plugin('0.0.0-alpha.1', notes);
  try {
    await back.seen('0.0.0-alpha.2');
    assert.equal(await back.whatsNew.pending(), null, 'alpha.2 → alpha.1 is a downgrade');
  } finally { await back.done(); }
});

// STRATEGY-7308: the card follows DSH's interface language. English points come from each section's ### Highlights.
test('### Highlights gives a section its English points, one per Chinese point', () => {
  const both = '## 0.0.2 — d\n\n### 更新要点\n\n- 第一条\n- 第二条\n\n### Highlights\n\n- First `x`\n  continued.\n- **Second**\n\n### 详细\n\n- not a point\n';
  assert.deepEqual(parseWhatsNew(both).entries, [{ version: '0.0.2', highlights: ['第一条', '第二条'], highlightsEn: ['First x continued.', 'Second'] }]);
  assert.throws(() => parseWhatsNew('## 0.0.2 — d\n\n### 更新要点\n\n- 一\n- 二\n\n### Highlights\n\n- One\n'), /1 points.*2.*one English point for each/);
  assert.throws(() => parseWhatsNew('## 0.0.2 — d\n\n### Highlights\n\n- One\n'), /without ### 更新要点/);
  assert.throws(() => parseWhatsNew('## 0.0.2 — d\n\n### 更新要点\n\n- 一\n\n### Highlights\n\n- \n'), /### Highlights has a blank point/);
  assert.deepEqual(missingEnglish(parseWhatsNew(CHANGELOG)), ['unreleased', '0.0.2'], 'sections without ### Highlights are named');
});

test('every committed section with highlights has its English points', async () => {
  const notes = parseWhatsNew(await readFile(join(adapter, 'CHANGELOG.md'), 'utf8'));
  assert.deepEqual(missingEnglish(notes), []);
  for (const entry of notes.entries) {
    assert.equal(entry.highlightsEn.length, entry.highlights.length, entry.version);
    for (const text of entry.highlightsEn) assert.doesNotMatch(text, /[一-鿿]/, `${entry.version}: English points carry no Chinese`);
  }
});

test('the Host keeps English points per version, drops malformed ones, and merges them only when every part has them', async () => {
  const dir = await mkdtemp(join(tmpdir(), 'whats-new-en-'));
  try {
    const file = join(dir, 'whats-new.json');
    await writeFile(file, JSON.stringify({ v: 1, latestReleased: '0.0.46', entries: [
      { version: 'unreleased', highlights: ['u1'], highlightsEn: ['U1'] },
      { version: '0.0.46', highlights: ['a', 'b'], highlightsEn: ['A'] },
      { version: '0.0.45', highlights: ['c'], highlightsEn: [''] },
      { version: '0.0.44', highlights: ['d'], highlightsEn: ['D'] },
    ] }));
    const notes = readNotes(file);
    assert.deepEqual(notes.entries.map(entry => entry.highlightsEn), [['U1'], undefined, undefined, ['D']], 'a list of another length or with blanks is dropped');
    assert.deepEqual(notesBetween(notes, '0.0.43', '0.0.47'), [
      { version: '0.0.47', highlights: ['u1'], highlightsEn: ['U1'] }, { version: '0.0.46', highlights: ['a', 'b'] },
      { version: '0.0.45', highlights: ['c'] }, { version: '0.0.44', highlights: ['d'], highlightsEn: ['D'] }]);
    const merged = { latestReleased: '0.0.46', entries: [{ version: '0.0.47', highlights: ['x'], highlightsEn: ['X'] }, { version: '0.0.47', highlights: ['y'] }] };
    assert.deepEqual(notesBetween(merged, undefined, '0.0.47'), [{ version: '0.0.47', highlights: ['x', 'y'] }], 'never half English');
  } finally { await rm(dir, { recursive: true, force: true }); }
});

test('the card shows English points under an English interface and Chinese otherwise, falling back per version', () => {
  const notes = { current: '0.0.47', previous: '0.0.45', entries: [
    { version: '0.0.47', highlights: ['一', '二'], highlightsEn: ['One', 'Two'] }, { version: '0.0.46', highlights: ['三'] }] };
  assert.deepEqual(whatsNewView(notes, true, undefined, 'en').groups, [{ version: '0.0.47', highlights: ['One', 'Two'] }, { version: '0.0.46', highlights: ['三'] }]);
  assert.deepEqual(whatsNewView(notes, true, undefined, 'en-US').groups[0].highlights, ['One', 'Two']);
  for (const lang of ['zh', 'zh-CN', undefined, '', 'ja']) assert.deepEqual(whatsNewView(notes, true, undefined, lang).groups[0].highlights, ['一', '二'], String(lang));
  assert.deepEqual(whatsNewView(notes).groups, whatsNewView(notes, false, WHATS_NEW_VISIBLE, 'zh').groups, 'no language: Chinese, as before');
  assert.deepEqual([notesLanguage('EN'), notesLanguage('zh'), notesLanguage(null)], ['en', 'zh', 'zh']);
  const short = whatsNewView(notes, false, 1, 'en');
  assert.deepEqual([short.groups, short.hidden], [[{ version: '0.0.47', highlights: ['One'] }], 2]);
});
