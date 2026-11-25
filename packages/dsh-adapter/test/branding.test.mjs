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
  '../locale/en.json': 'Preview and edit Word, Excel and PowerPoint files (DOCX/XLSX/PPTX) in a Harness side tab. The document engine downloads the first time you open a file; after that files open offline.',
  '../locale/tools/zh.json': '让对话里的 AI 新建、读取、修改、保存 PPTX、DOCX 和 XLSX，每一步都在「OfficeSDK」标签页里展示。需要同时开启「OfficeSDK」。',
  '../locale/tools/en.json': 'Lets the AI in the chat create, read, edit and save PPTX, DOCX and XLSX files, showing each step in the OfficeSDK tab. Requires OfficeSDK to be enabled.',
  '../locale/skills/zh.json': 'AI 新建文件时用的默认风格：shimo-pptx（PPT，蓝色品牌风格）、shimo-docx（Word，公文版式）。随插件自带，无需下载；你自己放的同名 skill 优先。',
  '../locale/skills/en.json': 'Default styles the AI uses for new files: shimo-pptx (PPT, blue house style) and shimo-docx (Word, official-document layout). They ship with the plugin, nothing to download; a same-named skill of your own takes precedence.',
};

test('Chinese package overview describes all three editors, with short summaries for the other entries', () => {
  const { description } = JSON.parse(read('../locale/zh.json')).meta;
  const paragraphs = description.split('\n\n');
  assert.equal(paragraphs.length, 3);
  assert.match(paragraphs[0], /^装上这款插件后，.*保存回原文件/);
  assert.match(paragraphs[1], /^现有编辑器能力\n.*233 项功能（持续更新中）.*Word 文稿 81 项、Excel 表格 64 项、PowerPoint 幻灯片 88 项/);
  assert.match(paragraphs[2], /每处修改都可以回到原文位置核对。$/);
  assert.doesNotMatch(description, /\*\*|<[^>]+>/);
  assert.deepEqual(Object.keys(DESCRIPTIONS).sort(), LOCALES.filter(path => path !== '../locale/zh.json').sort());
  for (const [path, text] of Object.entries(DESCRIPTIONS)) {
    const { description } = JSON.parse(read(path)).meta;
    assert.equal(description, text, path);
    assert.ok(description.split(/。|\.(?:\s|$)/).filter(Boolean).length <= 2, `${path}: at most two sentences`);
    // The engine downloads on first use (controller.mjs); the old "bundled, no network needed" was wrong.
    assert.doesNotMatch(description, /随包提供|不需要联网|bundled|no network needed/i, path);
  }
  for (const lang of ['zh', 'en']) {
    const skills = JSON.parse(read(`../locale/skills/${lang}.json`)).meta.description;
    for (const name of ['shimo-pptx', 'shimo-docx']) assert.match(skills, new RegExp(name), lang);
  }
  // The details dropped from the tools text live in the README.
  const readme = read('../../../README.md');
  assert.match(readme, /^## Model editing tools$/m);
  assert.match(readme, /only after you confirm/);
  assert.match(readme, /cannot be undone automatically/);
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
  assert.match(read('../../../README.md'), /「OfficeSDK · 默认风格 skill」 on the plugin page/);
  assert.doesNotMatch(read('../../../README.md'), /石墨|\bShimo\b/);
});
