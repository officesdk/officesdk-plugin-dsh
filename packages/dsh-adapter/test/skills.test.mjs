import test from 'node:test';
import assert from 'node:assert/strict';
import { cpSync, mkdtempSync, readFileSync, readdirSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { BUNDLED_RANK, PROVIDER, SKILL_NAMES, apply, bundledSkills, createProvider, parseSkill } from '../src/skills/index.mjs';

const ROOT = fileURLToPath(new URL('../skills/', import.meta.url));
const manifest = JSON.parse(readFileSync(new URL('../package.json', import.meta.url), 'utf8'));

test('every bundled skill ships and lists with description, whenToUse and resources', () => {
  const warnings = [];
  const candidates = bundledSkills(ROOT, text => warnings.push(text));
  assert.deepEqual(warnings, []);
  assert.deepEqual(candidates.map(c => c.name), SKILL_NAMES);
  assert.deepEqual(SKILL_NAMES, ['shimo-pptx', 'shimo-docx']);
  for (const c of candidates) {
    assert.equal(c.rank, BUNDLED_RANK);
    assert.equal(c.provider, PROVIDER);
    assert.equal(c.source, 'bundled');
    assert.deepEqual(c.invocation, { modelInvocable: true, userInvocable: true });
    assert.match(c.description, /^(新建 (PPT|Word)|做或改 Excel 表格)/);
    assert.match(c.whenToUse, /加载/);
    assert.equal(c.path, join(ROOT, c.name, 'SKILL.md'));
    assert.equal(c.resourceBase.kind, 'directory');
    assert.ok(c.resourceBase.path.endsWith(`/skills/${c.name}`));
    assert.ok(!('metadata' in c));
  }
  // One PPT skill with both ways to build a deck, and when the offline one applies.
  const [pptx, docx] = candidates;
  assert.match(pptx.description, /实时生成/);
  assert.match(pptx.description, /presentation_create/);
  assert.match(pptx.description, /brand_deck\.py/);
  assert.match(pptx.description, /样式包/);
  for (const reason of [/表格或图表/, /批量/, /presentation_create/, /脚本出稿/]) assert.match(pptx.whenToUse, reason);
  assert.match(docx.description, /^新建 Word/);
  // 261109: Word and Excel are built live by default, red letterhead and page numbers included; a script is the fallback.
  assert.match(docx.description, /默认在右侧 OfficeSDK 编辑器里实时写/);
  assert.match(docx.description, /红头公文和页码也一样/);
  assert.doesNotMatch(docx.whenToUse, /用户想看着写/);
});

test('261109: the docx skill writes red letterhead live and keeps the script for what the editor cannot do, saying why', () => {
  const docx = readFileSync(join(ROOT, 'shimo-docx', 'SKILL.md'), 'utf8');
  const live = docx.slice(0, docx.indexOf('## 离线脚本方式'));
  assert.match(live, /`redhead`：通知、函、请示/);
  assert.match(live, /用 `setText` 改这两段/);
  assert.match(live, /不要因为是红头、要页码就改用离线脚本/);
  assert.match(live, /一句话告诉用户为什么没在右侧生成/);
  for (const reason of [/要\*\*普通表格\*\*/, /要图片/, /签发人/, /一次要出多份/]) assert.match(live, reason);
  assert.doesNotMatch(docx, /要红头（发文机关标志、红线、发文字号、版记）/, 'red letterhead is no longer a reason to go offline');
});

test('skill versions follow the plugin rule (shimo-pptx 0.0.11, shimo-docx 0.0.7) and the old names are gone from the shipped skills', () => {
  const versions = { 'shimo-pptx': '0.0.11', 'shimo-docx': '0.0.7' };
  assert.deepEqual(Object.keys(versions).toSorted(), [...SKILL_NAMES].toSorted());
  for (const name of SKILL_NAMES) assert.match(readFileSync(join(ROOT, name, 'SKILL.md'), 'utf8'), new RegExp(`\\n {2}version: ${versions[name].replaceAll('.', '\\.')}\\n`), name);
  // Plan the whole deck before page content is available.
  const pptx = readFileSync(join(ROOT, 'shimo-pptx', 'SKILL.md'), 'utf8');
  assert.match(pptx, /确定页数后立即新建/);
  assert.match(pptx, /completedSlideIds/);
  assert.match(pptx, /整批撤回/);
  assert.doesNotMatch(pptx, /一页一次调用|每页一次/);
  // The chat keeps a short result; the slides keep content only; city pictures go full width, not into 312×270.
  assert.match(pptx, /不要往聊天里写进度/);
  assert.match(pptx, /只回两三句/);
  assert.match(pptx, /不要贴 `sources\.md` 全文/);
  assert.match(pptx, /标题和副标题占位符/);
  assert.match(pptx, /每行 2–6 个汉字/);
  assert.match(pptx, /内容页的页标题只占一行/, 'two-line titles are for the large title boxes only');
  assert.match(pptx, /default 样式里城市、风景这类横图用全幅图文/, 'photo-statement is a default slot');
  assert.match(pptx, /限流/);
  assert.match(pptx, /--box 720x354\.6/);
  assert.match(pptx, /--anchor top\|bottom\|left\|right/);
  assert.doesNotMatch(pptx, /--box 312x270/);
  // The editor is called by the plugin's display name, OfficeSDK.
  assert.match(pptx, /右侧的 OfficeSDK 编辑器里实时生成/);
  assert.doesNotMatch(pptx, /石墨/);
  // shimo-docx writes live with the document tools, and says plainly there is no rollback.
  const docx = readFileSync(join(ROOT, 'shimo-docx', 'SKILL.md'), 'utf8');
  for (const tool of ['document_create', 'document_edit', 'document_read', 'document_save']) assert.match(docx, new RegExp(tool));
  assert.match(docx, /没有回滚/);
  assert.match(docx, /不要说已回滚或已撤销/);
  assert.match(docx, /## 离线脚本方式/, 'the offline way stays');
  const files = dir => readdirSync(dir, { withFileTypes: true }).flatMap(entry => entry.isDirectory() ? files(join(dir, entry.name)) : [join(dir, entry.name)]);
  for (const file of files(ROOT).filter(file => /\.(md|json|py)$/.test(file) && !file.endsWith('CHANGELOG.md'))) {
    assert.doesNotMatch(readFileSync(file, 'utf8'), /dsh-(pptx|docx)-(style|live)/, file);
  }
});

test('the bundled rank sorts after every project and user skill root', () => {
  // dsh-skill-filesystem: project .dsh 100, project .agents 200, custom 300, ~/.dsh 400, ~/.agents 500; packaged 600.
  assert.equal(BUNDLED_RANK, 600);
});

test('get() reads the current body and keeps the summary fields', async () => {
  const root = mkdtempSync(join(tmpdir(), 'officesdk-skills-'));
  cpSync(ROOT, root, { recursive: true });
  const provider = createProvider(bundledSkills(root));
  const [pptx] = await provider.list();
  const first = await provider.get(pptx, {});
  assert.equal(first.name, 'shimo-pptx');
  assert.equal(first.whenToUse, pptx.whenToUse);
  assert.ok(!('rank' in first) && !('locator' in first));
  assert.match(first.content, /^# 新建 PPT/);
  const file = join(root, 'shimo-pptx', 'SKILL.md');
  writeFileSync(file, readFileSync(file, 'utf8').replace('# 新建 PPT', '# 新建 PPT（改）'));
  assert.match((await provider.get(pptx, {})).content, /^# 新建 PPT（改）/);
});

test('a bad SKILL.md skips only that skill', () => {
  const root = mkdtempSync(join(tmpdir(), 'officesdk-skills-'));
  cpSync(ROOT, root, { recursive: true });
  writeFileSync(join(root, 'shimo-docx', 'SKILL.md'), '---\nname: Bad_Name\ndescription: x\n---\nbody');
  const warnings = [];
  assert.deepEqual(bundledSkills(root, text => warnings.push(text)).map(c => c.name), ['shimo-pptx']);
  assert.match(warnings[0], /shimo-docx/);
  writeFileSync(join(root, 'shimo-pptx', 'SKILL.md'), 'no frontmatter');
  assert.deepEqual(bundledSkills(root, () => {}), []);
});

test('apply registers one provider; nothing when no skill is valid', async () => {
  const registered = [];
  const ctx = { skills: { registerProvider: create => registered.push(create({ signal: new AbortController().signal, invalidate() {} })) } };
  apply(ctx, {}, { root: ROOT });
  assert.equal(registered.length, 1);
  assert.equal(registered[0].name, PROVIDER);
  assert.deepEqual((await registered[0].list()).map(c => c.name), SKILL_NAMES);
  apply(ctx, {}, { root: mkdtempSync(join(tmpdir(), 'officesdk-skills-empty-')) });
  assert.equal(registered.length, 1);
});

test('frontmatter parser: quoted, block and nested values', () => {
  const { fields, content } = parseSkill('---\nname: "a-b"\ndescription: |\n  line one\n  line two\nwhenToUse: >\n  folded\n  text\nmetadata:\n  version: 1.0.0\n---\n\n# Body\n');
  assert.deepEqual(fields, { name: 'a-b', description: 'line one\nline two', whenToUse: 'folded text' });
  assert.equal(content, '# Body');
  const odd = parseSkill(`\uFEFF---\nname: a-b # the id\ndescription: "say \\"hi\\"\\nnow"\nwhenToUse: 'it''s'\n---\nx`);
  assert.deepEqual(odd.fields, { name: 'a-b', description: 'say "hi"\nnow', whenToUse: "it's" });
  assert.throws(() => parseSkill('---\nname: a\ndescription: "bad \\q"\n---\n'));
});

test('the package ships the skills sub-plugin and its files', () => {
  assert.equal(manifest.exports['./skills'], './lib/skills/index.mjs');
  assert.equal(manifest.exports['./skills/locale/*.json'], './locale/skills/*.json');
  assert.ok(manifest.files.includes('lib/skills') && manifest.files.includes('skills'));
  const plugin = readFileSync(new URL('../plugin.yml', import.meta.url), 'utf8');
  assert.match(plugin, /- id: officesdk-skills\n\s+name: officesdk\/skills\n/);
  // Option A: enabled by default; the user turns it off on the plugins page.
  assert.doesNotMatch(plugin.split('officesdk-skills')[1], /disabled/);
  for (const lang of ['zh', 'en']) assert.ok(JSON.parse(readFileSync(new URL(`../locale/skills/${lang}.json`, import.meta.url), 'utf8')).meta.title);
});

test('shimo-pptx picks a style in three layers: named, else the workspace default (no style passed), else it recommends 2–4 through the tool; photos are warned about first', () => {
  const skill = readFileSync(join(ROOT, 'shimo-pptx', 'SKILL.md'), 'utf8');
  const section = skill.slice(skill.indexOf('## 2. 选样式'), skill.indexOf('## 3. 配图'));
  for (const rule of [/用户点名了/, /meta\.aliases/, /仅|只有用户主动要求比较风格/, /能满足页数的包/, /intent/, /previewPhase="prepare"/, /previewPhase="ready"/, /时间轴/]) assert.match(section,rule);
  assert.doesNotMatch(section,/先推荐再动手/);
});
