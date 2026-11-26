import test from 'node:test';
import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import { readFileSync } from 'node:fs';
import { BUILTIN, BULB_PATHS, CARD_WIDTH, FRAME_RATIO, HOVER_DELAY_MS, HOVER_EVERY_MS, HOVER_PAGES, PROMPTS, SCOPE_FEATURED, SCOPE_MINE, UPLOAD_EXTENSIONS, categoriesOf, checkUpload, columnCount, composeDraft, featurePage, libraryTemplates, masonry, modeKey, pageKey, promptOf, scopeTemplates, stripPrompt, summaryLines, templatesOf } from '../src/client/inspiration.mjs';
import { FILE_ICONS, FILE_ICON_SHA256, OUTLINE_ICONS, OUTLINE_ICON_SHA256 } from '../src/client/inspiration-icons.mjs';
import { BUILTIN_TEMPLATES } from '../src/host/template-pick.mjs';

const ICONS = new URL('../src/client/icons/', import.meta.url);
// deliveredSha256 of icons/来源.json in the inspiration library design handoff v1.0 (Small icons).
const HANDOFF_SHA256 = {
  word: 'bdc661c07052bef32d08abbc7e3a2943b979794772c66226894911c0944eda4b',
  excel: 'ce78c58128d87450ae9c5c35806634c6c5d9f8b6f7e77672815ec0b7c87d02b2',
  ppt: '3d3de24f4f97595af9269695715dcc6de244c10074d11e73a89fbfaa0525d2fb',
};
const sha = bytes => createHash('sha256').update(bytes).digest('hex');

test('file icons are the handoff\'s Small originals byte for byte, as data URIs', () => {
  for (const kind of ['word', 'excel', 'ppt']) {
    const prefix = 'data:image/svg+xml;base64,';
    assert.ok(FILE_ICONS[kind].startsWith(prefix));
    const bytes = Buffer.from(FILE_ICONS[kind].slice(prefix.length), 'base64');
    assert.equal(sha(bytes), HANDOFF_SHA256[kind], kind);
    assert.equal(FILE_ICON_SHA256[kind], HANDOFF_SHA256[kind]);
    assert.deepEqual(bytes, readFileSync(new URL(`${kind}-small.svg`, ICONS)), `${kind}: the same bytes as the copy in the repo`);
  }
});

// sha256 of icons/*-outline.svg and inspiration-bulb.svg in the handoff OfficeSDK-DSH-Inspiration-v2.0.
const HANDOFF_V2 = new URL('../src/client/icons/', import.meta.url);

test('outline icons (unselected switch state) are the handoff v2.0 files byte for byte, as data URIs', () => {
  for (const kind of ['word', 'excel', 'ppt']) {
    const prefix = 'data:image/svg+xml;base64,';
    const bytes = Buffer.from(OUTLINE_ICONS[kind].slice(prefix.length), 'base64');
    assert.ok(OUTLINE_ICONS[kind].startsWith(prefix));
    assert.deepEqual(bytes, readFileSync(new URL(`${kind}-outline.svg`, HANDOFF_V2)));
    assert.equal(OUTLINE_ICON_SHA256[kind], sha(bytes));
    assert.match(bytes.toString('utf8'), /currentColor/, 'drawn in the text colour');
  }
});

test('the bulb is the v2.0 one with the thin bolt: both paths of inspiration-bulb.svg, currentColor, 1.5 stroke', () => {
  const svg = readFileSync(new URL('inspiration-bulb.svg', ICONS), 'utf8');
  assert.equal(BULB_PATHS.length, 2);
  for (const d of BULB_PATHS) assert.ok(svg.includes(`d="${d}"`), d);
  assert.match(svg, /stroke="currentColor" stroke-width="1\.5"/);
});

test('fixed prompts per kind (§7), and the draft keeps what the user wrote', () => {
  assert.equal(promptOf('word'), '使用此文档技能帮助我创建文档：');
  assert.equal(promptOf('excel'), '使用此表格技能帮助我创建文档：');
  assert.equal(promptOf('ppt'), '使用此幻灯片技能帮助我创建文档：');
  assert.ok(promptOf('ppt', 'en').startsWith('Use this slides skill'));
  const ppt = PROMPTS.zh.ppt, word = PROMPTS.zh.word;
  assert.equal(composeDraft('', undefined, ppt), ppt);
  assert.equal(composeDraft('重庆三日游', undefined, ppt), `${ppt}重庆三日游`, 'the prompt goes first, no line break');
  assert.equal(composeDraft(`${ppt}重庆三日游`, ppt, word), `${word}重庆三日游`, 'another template replaces only the old prompt');
  const rewritten = `请帮我${ppt.slice(2)}重庆`;
  assert.equal(composeDraft(rewritten, ppt, word), `${word}${rewritten}`, 'a prompt the user changed is left alone');
  assert.equal(composeDraft('  \n重庆', undefined, ppt), `${ppt}重庆`);
  assert.equal(stripPrompt(`${ppt}重庆`, ppt), '重庆');
  assert.equal(stripPrompt(`重庆${ppt}`, ppt), `重庆${ppt}`, 'only at the start');
  assert.equal(stripPrompt(undefined, ppt), '');
  assert.equal(stripPrompt('x', undefined), 'x');
});

test('the mixed catalog: Word, Excel then PPT; PPT pages are the thumbnail then the previews; scenes are the top ten', () => {
  const styles = [
    { id: 'a', name: { zh: '甲' }, summary: { zh: '甲的摘要' }, thumb: 'data:a', previews: ['data:a2', 'data:a3', 'data:a'], scenarios: { zh: ['业务汇报', '品牌'] } },
    { id: 'b', name: { zh: '乙' }, thumb: null, previews: [], scenarios: { zh: ['旅行', '品牌'] } },
  ];
  const all = libraryTemplates(styles);
  assert.deepEqual(all.map(item => `${item.kind}:${item.id}`), ['word:meeting-minutes', 'word:gongwen', 'ppt:a', 'ppt:b']);
  const a = all.find(item => item.id === 'a');
  assert.deepEqual(a.pages, ['data:a', 'data:a2', 'data:a3'], 'thumbnail first, no repeats');
  assert.equal(a.summary, '甲的摘要');
  assert.deepEqual(all.find(item => item.id === 'b').pages, [], 'no pages: the card shows the missing face');
  for (const item of all.filter(entry => entry.kind !== 'ppt')) assert.equal(item.pages.length, 1, `${item.id}: one rendered page`);
  assert.deepEqual(categoriesOf(templatesOf('ppt', styles)), ['品牌', '业务汇报', '旅行']);
  const many = Array.from({ length: 14 }, (_, i) => ({ id: `p${i}`, categories: [`c${i}`] }));
  assert.equal(categoriesOf(many).length, 10, 'ten scenes at most');
  assert.deepEqual(categoriesOf(many).slice(0, 2), ['c0', 'c1'], 'ties keep the first-use order');
});

test('scopes: featured (default) is every built-in template of all kinds; mine has none until a service exists; a scene filters', () => {
  const all = libraryTemplates([{ id: 'a', name: { zh: '甲' }, thumb: 'data:a', scenarios: { zh: ['旅行'] } }]);
  assert.equal(scopeTemplates(all, SCOPE_FEATURED).length, all.length);
  assert.equal(scopeTemplates(all, undefined).length, all.length, 'featured by default');
  assert.deepEqual(scopeTemplates(all, SCOPE_MINE), []);
  assert.deepEqual(scopeTemplates(all, '旅行').map(item => item.id), ['a']);
  assert.deepEqual(scopeTemplates(all, '公文通知').map(item => item.id), ['gongwen']);
});

test('masonry: fixed 280px columns, as many as fit (never stretched, never more than the cards); each card into the shortest column', () => {
  assert.equal(CARD_WIDTH, 280);
  assert.equal(columnCount(1200, 20, 20), 4);
  assert.equal(columnCount(879, 20, 20), 2, '3 × 280 + 2 × 20 = 880 does not fit in 879');
  assert.equal(columnCount(880, 20, 20), 3);
  assert.equal(columnCount(2000, 20, 2), 2, 'two cards never fill the row');
  assert.equal(columnCount(100, 20, 5), 1, 'very narrow: one column (the area scrolls sideways)');
  assert.equal(columnCount(1000, 20, 0), 0);
  const items = [{ id: 'w', kind: 'word' }, { id: 'p1', kind: 'ppt' }, { id: 'p2', kind: 'ppt' }, { id: 'p3', kind: 'ppt' }, { id: 'x', kind: 'excel' }];
  const columns = masonry(items, 2);
  // Word 264×297/210+68 ≈ 441px; each PPT ≈ 216px: p1, p2 (433 < 441) and p3 go right, then column 1 is the shorter one.
  assert.deepEqual(columns.map(list => list.map(item => item.id)), [['w', 'x'], ['p1', 'p2', 'p3']]);
  assert.deepEqual(masonry(items, 0), []);
});

test('timing, summary lines, example page, page and mode keys', () => {
  assert.deepEqual([HOVER_DELAY_MS, HOVER_EVERY_MS, HOVER_PAGES], [1000, 3000, 3]);
  assert.equal(summaryLines(100, 20, false), 5);
  assert.equal(summaryLines(100, 20, true), 3, '24px kept for 更多');
  assert.equal(summaryLines(5, 20, true), 1, 'at least one');
  assert.equal(summaryLines(100, 0, false), 1);
  assert.equal(featurePage(['1', '2', '3']), '2');
  assert.equal(featurePage(['1']), '1');
  assert.equal(featurePage([]), undefined);
  assert.equal(pageKey('ArrowLeft', 0, 4), 0);
  assert.equal(pageKey('ArrowRight', 3, 4), 3);
  assert.equal(pageKey('Home', 2, 4), 0);
  assert.equal(pageKey('End', 0, 4), 3);
  assert.equal(pageKey('x', 0, 4), undefined);
  assert.equal(modeKey('ArrowRight', 2), 0, 'wraps');
  assert.equal(modeKey('ArrowUp', 0), 2);
  assert.equal(modeKey('ArrowDown', 0), 1);
  assert.equal(modeKey('Home', 2), 0);
  assert.equal(modeKey('End', 0), 2);
  assert.equal(modeKey('Enter', 0), undefined);
});

test('upload checks: one non-empty file of the kind\'s extensions; never reads it', () => {
  assert.deepEqual(checkUpload([], 'ppt'), { error: 'none' });
  assert.deepEqual(checkUpload(undefined, 'ppt'), { error: 'none' });
  assert.deepEqual(checkUpload([{ name: 'a.pptx', size: 1 }, { name: 'b.pptx', size: 1 }], 'ppt'), { error: 'multiple' });
  assert.deepEqual(checkUpload([{ name: 'a.pptx', size: 0 }], 'ppt'), { error: 'empty' });
  assert.deepEqual(checkUpload([{ name: 'a.xlsx', size: 3 }], 'ppt'), { error: 'type' });
  assert.deepEqual(checkUpload([{ name: 'noext', size: 3 }], 'word'), { error: 'type' });
  assert.deepEqual(checkUpload([{ name: '照片.JPEG', size: 3 }], 'excel'), { file: { name: '照片.JPEG', size: 3 } });
  assert.deepEqual(UPLOAD_EXTENSIONS.word, ['.docx', '.doc', '.png', '.jpg', '.jpeg', '.webp']);
});

test('the Host\'s template list keeps only what this copy can create from: the public build has Word gongwen and no Excel card (261113)', () => {
  const ids = list => list.map(item => `${item.kind}:${item.id}`);
  assert.deepEqual(ids(libraryTemplates([], 'zh', { word: ['gongwen'], excel: [] })), ['word:gongwen']);
  assert.deepEqual(ids(templatesOf('word', [], 'zh')), ['word:meeting-minutes', 'word:gongwen'], 'an older Host sends no list: every built-in');
  assert.deepEqual(ids(templatesOf('excel', [], 'zh', { word: ['gongwen'] })), [], 'no Excel template is built in');
});
