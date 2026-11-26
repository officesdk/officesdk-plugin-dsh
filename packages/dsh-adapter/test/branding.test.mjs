import test from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';

// The name users see (document tab header, DSH plugin list, component names) is OfficeSDK in both languages.
// Identifiers keep the shimo prefix (officesdk, shimo-pptx, /shimo, shimoPreview); only display text is checked here.
const NAME = 'OfficeSDK';
const OLD = [/石墨/, /Shimo Docs/i, /Shimo document/i, /\bshimo\s+\w/i];
const read = path => readFileSync(new URL(path, import.meta.url), 'utf8');
const LOCALES = ['../locale/zh.json', '../locale/en.json', '../locale/tools/zh.json', '../locale/tools/en.json', '../locale/skills/zh.json', '../locale/skills/en.json'];

test('plugin and component names in every locale start with OfficeSDK and drop the old name', () => {
  for (const path of LOCALES) {
    const { meta } = JSON.parse(read(path));
    assert.ok(meta.title === NAME || meta.title.startsWith(`${NAME} · `), `${path} title: ${meta.title}`);
    for (const text of [meta.title, meta.description]) for (const old of OLD) assert.doesNotMatch(text, old, path);
  }
  assert.equal(JSON.parse(read('../locale/zh.json')).meta.title, NAME);
  assert.equal(JSON.parse(read('../locale/en.json')).meta.title, NAME);
});

// The plugin page shows each component's meta.description: the Chinese package overview and concise component summaries.
// The skills component's text is display only; the model picks a skill by its SKILL.md description, which stays as it is.
const DESCRIPTIONS = {
  '../locale/en.json': 'Edit Word documents, Excel spreadsheets, and PowerPoint presentations directly in DeepSeek Harness. Open a file, change its content or formatting, and save it back to the original without opening a separate Office app. The editors currently offer 233 features across Word (81), Excel (64), and PowerPoint (88). Edit the file yourself or ask AI to revise a specific passage, cell range, or slide text box. Review each change where it was made.',
  '../locale/tools/zh.json': '让对话里的 AI 新建、读取、修改和保存 Word、Excel 和 PowerPoint 文件。你可以在右侧编辑器中查看写入过程和修改结果。使用时需同时开启 OfficeSDK 主插件。',
  '../locale/tools/en.json': 'Let the AI in your conversation create, read, edit, and save Word, Excel, and PowerPoint files. Follow live edits and review the results in the side editor. Enable the main OfficeSDK plugin to use these tools.',
  '../locale/skills/zh.json': '内置 Word 和 PPT 默认风格。Word 可起草会议纪要、工作周报、产品需求文档、项目计划、项目建议书和正式信函。PPT 可制作团队工作汇报、员工培训和经验分享。两者都可从空白页开始。风格随插件安装，你自己的同名 skill 优先。',
  '../locale/skills/en.json': 'Start Word documents and PowerPoint presentations with built-in styles. In Word, draft meeting minutes, weekly reports, product requirements, project plans, proposals, and formal letters. In PowerPoint, create team updates, training decks, and knowledge-sharing presentations. Both can start from a blank page or slide. The styles ship with the plugin, and your own skills of the same name take precedence.',
};

test('Chinese package overview describes all three editors, with short summaries for the other entries', () => {
  const { description } = JSON.parse(read('../locale/zh.json')).meta;
  assert.equal(description.split('\n\n').length, 1);
  assert.match(description, /^插件可直接编辑 Word 文稿、Excel 表格和 PowerPoint 幻灯片。/);
  assert.match(description, /233 项功能（持续更新中）.*Word 文稿 81 项、Excel 表格 64 项、PowerPoint 幻灯片 88 项/);
  assert.match(description, /每处修改都可以回到原文位置核对。$/);
  assert.doesNotMatch(description, /\*\*|<[^>]+>/);
  assert.deepEqual(Object.keys(DESCRIPTIONS).sort(), LOCALES.filter(path => path !== '../locale/zh.json').sort());
  for (const [path, text] of Object.entries(DESCRIPTIONS)) {
    const { description } = JSON.parse(read(path)).meta;
    assert.equal(description, text, path);
    // The engine downloads on first use (controller.mjs); the old "bundled, no network needed" was wrong.
  }
  for (const lang of ['zh', 'en']) {
    const skills = JSON.parse(read(`../locale/skills/${lang}.json`)).meta.description;
    assert.match(skills, /Word/);
    assert.match(skills, lang === 'zh' ? /PPT/ : /PowerPoint/);
    assert.match(skills, lang === 'zh' ? /同名 skill 优先/ : /same name take precedence/);
  }
  // The details dropped from the tools text live in the README.
  const readme = read('../../../README.md');
  assert.match(readme, /^### Confirming and undoing AI edits$/m);
  assert.match(readme, /asks for confirmation before applying changes/);
  assert.match(readme, /does not automatically undo edits/);
});

test('the document tab header is titled OfficeSDK in both languages', () => {
  const source = read('../src/client/plugin.mjs');
  const messages = source.slice(source.indexOf('const MESSAGES = {'), source.indexOf('\n};', source.indexOf('const MESSAGES = {')));
  assert.match(messages, /\ben: \{ title: 'OfficeSDK',/);
  assert.match(messages, /\bzh: \{ title: 'OfficeSDK',/);
  for (const old of OLD) assert.doesNotMatch(messages, old);
  assert.match(source, /title: t\?\.\('title'\) \|\| 'OfficeSDK'/);
  assert.doesNotMatch(source, /石墨|Shimo document|Shimo Docs/);
});

test('docs name the plugin and its skills component as OfficeSDK; the changelog keeps no old name', () => {
  const changelog = read('../CHANGELOG.md');
  assert.match(changelog, /^# OfficeSDK 更新日志\n/);
  for (const old of OLD) assert.doesNotMatch(changelog, old);
  assert.doesNotMatch(read('../whats-new.json'), /石墨|Shimo/i);
  for (const manifest of ['../package.json', '../../core/package.json']) assert.doesNotMatch(JSON.parse(read(manifest)).description, /石墨|\bShimo\b/, manifest);
  assert.match(read('../../../README.md'), /OfficeSDK · Default style skills/);
  assert.doesNotMatch(read('../../../README.md'), /石墨|\bShimo\b/);
});
