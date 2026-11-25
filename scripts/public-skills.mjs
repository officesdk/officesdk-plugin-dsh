/**
 * The skills the public npm build ships (spec 261113-bundle-public-skills): complete inside the package, never
 * downloaded. The repository holds exactly these; the trim below keeps a build honest should anything else appear.
 *
 *   shimo-pptx   SKILL.md, references/, scripts/ and the `default` style pack (蓝色品牌风格 brand-ppt-12)
 *   shimo-docx   SKILL.md, references/, scripts/, assets/, the `gongwen` template and the examples that are not
 *                another template's own
 *
 * Only the build's temporary copy is trimmed; the repository's skill directories stay as they are.
 * No dependencies: CI runs this before anything is installed.
 */
import fs from 'node:fs/promises';
import path from 'node:path';

export const PUBLIC_SKILLS = Object.freeze({
  'shimo-pptx': Object.freeze({ styles: Object.freeze(['default']), drop: Object.freeze([]) }),
  'shimo-docx': Object.freeze({ templates: Object.freeze(['gongwen']), drop: Object.freeze(['examples/redhead-notice.json']) }),
});
export const PUBLIC_SKILL_NAMES = Object.freeze(Object.keys(PUBLIC_SKILLS));
/** Files directly under skills/ that are not a skill (the skills' own change log) and stay out of the public package. */
export const PUBLIC_DROPPED_FILES = Object.freeze(['CHANGELOG.md']);

/** Whether file `rel` (a path inside skill `name`, with `/`) ships in the public package. */
export function publicSkillFile(name, rel) {
  const rule = PUBLIC_SKILLS[name];
  if (!rule) return false;
  if (rule.drop.includes(rel)) return false;
  const parts = rel.split('/');
  if (parts[0] === 'styles' && parts.length > 1) return Boolean(rule.styles?.includes(parts[1]));
  if (parts[0] === 'templates' && parts.length === 2) {
    const stem = parts[1].replace(/\.[^.]+$/, '');
    return Boolean(rule.templates?.includes(stem));
  }
  return true;
}

async function walk(root, prefix = '') {
  const out = [];
  for (const entry of await fs.readdir(path.join(root, prefix), { withFileTypes: true })) {
    const rel = prefix ? `${prefix}/${entry.name}` : entry.name;
    if (entry.isDirectory()) out.push(...await walk(root, rel));
    else out.push(rel);
  }
  return out;
}

/**
 * Trims `skillsRoot` (a build's skills/ directory) to the public set in place: other skills, other style packs and
 * templates, the dropped files and directories left empty are removed. Returns `{ [name]: [kept files] }`; throws when
 * a public skill or one of its named style packs or templates is missing (a build must never ship less than promised).
 */
export async function prunePublicSkills(skillsRoot) {
  const kept = {};
  for (const entry of await fs.readdir(skillsRoot, { withFileTypes: true })) {
    if (!entry.isDirectory()) {
      if (PUBLIC_DROPPED_FILES.includes(entry.name)) await fs.rm(path.join(skillsRoot, entry.name), { force: true });
      continue;
    }
    if (!PUBLIC_SKILLS[entry.name]) await fs.rm(path.join(skillsRoot, entry.name), { recursive: true, force: true });
  }
  for (const name of PUBLIC_SKILL_NAMES) {
    const dir = path.join(skillsRoot, name);
    const files = await walk(dir).catch(() => { throw new Error(`public skills: ${name} is missing from the build`); });
    kept[name] = [];
    for (const rel of files) {
      if (publicSkillFile(name, rel)) kept[name].push(rel);
      else await fs.rm(path.join(dir, ...rel.split('/')), { force: true });
    }
    await removeEmptyDirs(dir);
    const rule = PUBLIC_SKILLS[name];
    for (const style of rule.styles ?? []) if (!kept[name].some(rel => rel.startsWith(`styles/${style}/`))) throw new Error(`public skills: ${name} style pack ${style} is missing`);
    for (const template of rule.templates ?? []) if (!kept[name].some(rel => rel.startsWith(`templates/${template}.`))) throw new Error(`public skills: ${name} template ${template} is missing`);
    if (!kept[name].includes('SKILL.md')) throw new Error(`public skills: ${name} has no SKILL.md`);
    kept[name].sort();
  }
  return kept;
}

async function removeEmptyDirs(dir) {
  for (const entry of await fs.readdir(dir, { withFileTypes: true })) {
    if (!entry.isDirectory()) continue;
    const child = path.join(dir, entry.name);
    await removeEmptyDirs(child);
    if (!(await fs.readdir(child)).length) await fs.rmdir(child);
  }
}

/** Display text of the skills component for the public build (locale/skills/<lang>.json). */
export const PUBLIC_SKILLS_LOCALE = Object.freeze({
  zh: { meta: { title: 'OfficeSDK · 默认风格 skill', description: 'AI 新建文件时用的默认风格：shimo-pptx（PPT，蓝色品牌风格）、shimo-docx（Word，公文版式）。随插件自带，无需下载；你自己放的同名 skill 优先。' } },
  en: { meta: { title: 'OfficeSDK · Default style skills', description: 'Default styles the AI uses for new files: shimo-pptx (PPT, blue house style) and shimo-docx (Word, official-document layout). They ship with the plugin, nothing to download; a same-named skill of your own takes precedence.' } },
});

/**
 * capabilities.json of the public build, from the internal one (scripts/capabilities.mjs) and the public catalog's
 * `entries` (`{ [name]: catalog entry }`): skills, style packs and templates as the package ships them; spreadsheets
 * can start from the built-in blank workbook independently of the optional spreadsheet skill.
 */
export function publicCapabilities(capabilities, entries) {
  const next = JSON.parse(JSON.stringify(capabilities));
  next.skills = Object.entries(entries).map(([name, entry]) => ({ name, version: entry.version })).sort((a, b) => a.name.localeCompare(b.name));
  for (const product of next.products ?? []) {
    if (product.id === 'presentation') {
      product.stylePacks = [...entries['shimo-pptx']?.stylePacks ?? []];
      for (const item of product.capabilities ?? []) if (item.id === 'style-packs') item.summary = '内置蓝色品牌风格（default）';
    } else if (product.id === 'document') product.templates = [...entries['shimo-docx']?.templates ?? []];
    else if (product.id === 'spreadsheet') {
      product.templates = [...new Set(['blank', ...(entries['shimo-xlsx']?.templates ?? [])])];
      product.aiTools = [...new Set(['spreadsheet_create', ...(product.aiTools ?? [])])];
    }
  }
  return next;
}
