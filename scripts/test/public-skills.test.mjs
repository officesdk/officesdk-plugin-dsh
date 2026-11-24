import test from 'node:test';
import assert from 'node:assert/strict';
import { cp, mkdtemp, readdir, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { PUBLIC_SKILL_NAMES, prunePublicSkills, publicCapabilities, publicSkillFile } from '../public-skills.mjs';
import { readTree, skillEntry } from '../../packages/dsh-adapter/src/skills/catalog.mjs';

const repo = join(dirname(fileURLToPath(import.meta.url)), '../..');
const skillsRoot = join(repo, 'packages/dsh-adapter/skills');

test('the public file rule: default style pack and gongwen template only, the redhead example dropped, common files kept', () => {
  assert.ok(publicSkillFile('shimo-pptx', 'SKILL.md') && publicSkillFile('shimo-pptx', 'scripts/brand_deck.py') && publicSkillFile('shimo-pptx', 'references/style.md'));
  assert.ok(publicSkillFile('shimo-pptx', 'styles/default/template.pptx'));
  assert.ok(!publicSkillFile('shimo-pptx', 'styles/other-style/template.pptx'));
  assert.ok(publicSkillFile('shimo-docx', 'templates/gongwen.docx') && publicSkillFile('shimo-docx', 'assets/style.json') && publicSkillFile('shimo-docx', 'examples/plain-report.json'));
  for (const rel of ['templates/basic.docx', 'templates/redhead.docx', 'examples/redhead-notice.json']) assert.ok(!publicSkillFile('shimo-docx', rel), rel);
  assert.ok(!publicSkillFile('shimo-xlsx', 'SKILL.md'));
});

test('prunePublicSkills on a copy of the real skills: two skills left, exactly the public files; the repository is untouched', async () => {
  const before = await readdir(join(skillsRoot, 'shimo-pptx/styles'));
  const work = await mkdtemp(join(tmpdir(), 'public-skills-'));
  try {
    await cp(skillsRoot, join(work, 'skills'), { recursive: true });
    const kept = await prunePublicSkills(join(work, 'skills'));
    assert.deepEqual(Object.keys(kept), PUBLIC_SKILL_NAMES);
    assert.deepEqual((await readdir(join(work, 'skills'))).sort(), ['shimo-docx', 'shimo-pptx'], 'no shimo-xlsx, no CHANGELOG.md');
    assert.deepEqual(await readdir(join(work, 'skills/shimo-pptx/styles')), ['default']);
    assert.deepEqual(await readdir(join(work, 'skills/shimo-docx/templates')), ['gongwen.docx']);
    assert.deepEqual(await readdir(join(work, 'skills/shimo-docx/examples')), ['plain-report.json']);
    for (const name of ['template.pptx', 'style.json', 'layouts.json', 'thumb.png']) assert.ok(kept['shimo-pptx'].includes(`styles/default/${name}`), name);
    // The catalog entry of what is left lists only that (the same generator CI uses).
    const pptx = skillEntry('shimo-pptx', readTree(join(work, 'skills/shimo-pptx')));
    const docx = skillEntry('shimo-docx', readTree(join(work, 'skills/shimo-docx')));
    assert.deepEqual(pptx.stylePacks, ['default']);
    assert.deepEqual(docx.templates, ['gongwen']);
    assert.deepEqual(Object.keys(pptx.files).sort(), kept['shimo-pptx']);
    const caps = publicCapabilities({ skills: [{ name: 'shimo-xlsx', version: '0.0.3' }], products: [
      { id: 'presentation', aiTools: ['presentation_create'], capabilities: [{ id: 'style-packs', supported: true, summary: '多套风格' }], stylePacks: ['default', 'other-style'] },
      { id: 'document', aiTools: ['document_create'], capabilities: [], templates: ['basic', 'gongwen', 'redhead'] },
      { id: 'spreadsheet', aiTools: ['spreadsheet_create', 'spreadsheet_read'], capabilities: [], templates: ['blank'] },
    ] }, { 'shimo-pptx': pptx, 'shimo-docx': docx });
    assert.deepEqual(caps.skills, [{ name: 'shimo-docx', version: docx.version }, { name: 'shimo-pptx', version: pptx.version }]);
    assert.deepEqual(caps.products.map(product => [product.id, product.stylePacks ?? product.templates, product.aiTools]), [
      ['presentation', ['default'], ['presentation_create']], ['document', ['gongwen'], ['document_create']], ['spreadsheet', [], ['spreadsheet_read']]]);
    assert.equal(caps.products[0].capabilities[0].summary, '内置蓝色品牌风格（default）');
  } finally { await rm(work, { recursive: true, force: true }); }
  assert.deepEqual(await readdir(join(skillsRoot, 'shimo-pptx/styles')), before);
});

test('prunePublicSkills refuses a build missing what goes public', async () => {
  const work = await mkdtemp(join(tmpdir(), 'public-skills-'));
  try {
    await cp(skillsRoot, join(work, 'skills'), { recursive: true });
    await rm(join(work, 'skills/shimo-docx/templates/gongwen.docx'));
    await assert.rejects(prunePublicSkills(join(work, 'skills')), /shimo-docx template gongwen is missing/);
  } finally { await rm(work, { recursive: true, force: true }); }
});
