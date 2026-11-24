import { existsSync, readFileSync, unwatchFile, watchFile } from 'node:fs';
import { readFile } from 'node:fs/promises';
import { join, sep } from 'node:path';
import { fileURLToPath } from 'node:url';
import { homedir } from 'node:os';
import { SkillStore } from './store.mjs';
import { readDownloadSources, selectProfile } from '../host/fetch/download-sources.mjs';

/**
 * Skills that ship inside the plugin package (shimo-pptx: new decks, live in the right-side editor or offline by
 * script, from style packs; shimo-docx: new Word documents, live or offline)
 * (`skills/<name>/SKILL.md`), registered with DSH's skill
 * registry the way DSH's own `@deepseek-ai/dsh-skill-office` registers its bundled skills. Nothing is written to
 * `~/.dsh/skills`: the skills come and go with the plugin, and a same-named skill the user keeps in a project or user
 * skill folder outranks these (lower rank, or the agent preset's nearer layer).
 */
export const name = 'officesdk-skills';
export const inject = ['skills'];

export const SKILL_NAMES = ['shimo-pptx', 'shimo-docx'];
/** `BUNDLED_SKILL_RANK` of `@deepseek-ai/dsh-skill` 0.2.x: packaged providers sort after every project and user root. */
export const BUNDLED_RANK = 600;
export const PROVIDER = 'officesdk';
const SKILL_NAME = /^[a-z0-9]+(?:-[a-z0-9]+)*$/;
const defaultRoot = fileURLToPath(new URL('../../skills/', import.meta.url));
const defaultPackageRoot = fileURLToPath(new URL('../../', import.meta.url));
/** How long a skill call waits for its resources before answering without them (spec 261110-universal-package §4.5). */
export const GET_BUDGET_MS = 15_000;
const REFRESH_DELAY_MS = 30_000;
const REFRESH_EVERY_MS = 6 * 3600_000;

/**
 * Top-level scalar fields of a SKILL.md frontmatter block: single-line values (optionally quoted) and `|` / `>`
 * block scalars. Nested mappings such as `metadata` are skipped; a leading BOM is ignored. A YAML subset, enough for
 * these skills without bundling a YAML parser: a double-quoted value that is not valid JSON makes the skill invalid.
 */
export function parseSkill(raw) {
  raw = raw.replace(/^\uFEFF/, '');
  const match = /^---\r?\n([\s\S]*?)\r?\n---(?:\r?\n|$)/u.exec(raw);
  if (!match) throw new Error('no YAML frontmatter');
  const fields = {};
  const lines = match[1].split(/\r?\n/);
  for (let i = 0; i < lines.length; i++) {
    const field = /^([A-Za-z][\w-]*):(?:[ \t]+(.*))?$/.exec(lines[i]);
    if (!field) continue;
    const [, key, rest = ''] = field;
    const value = rest.trim();
    if (value === '|' || value === '>' || /^[|>][+-]?$/.test(value)) {
      const block = [];
      while (i + 1 < lines.length && (lines[i + 1] === '' || /^[ \t]/.test(lines[i + 1]))) block.push(lines[++i]);
      const indent = Math.min(...block.filter(line => line.trim()).map(line => line.length - line.trimStart().length));
      const body = block.map(line => line.slice(Number.isFinite(indent) ? indent : 0));
      fields[key] = (value.startsWith('>') ? body.join(' ').replace(/\s+/g, ' ') : body.join('\n')).trim();
    } else if (value) {
      // Double quotes take JSON escapes (\n, \"), single quotes double ' to escape it, plain values end at ` #`.
      if (/^".*"$/.test(value)) fields[key] = JSON.parse(value);
      else if (/^'.*'$/.test(value)) fields[key] = value.slice(1, -1).replaceAll("''", "'");
      else fields[key] = value.replace(/\s+#.*$/, '');
    }
  }
  return { fields, content: raw.slice(match[0].length).trim() };
}

/** One registry candidate per readable, valid bundled skill; a bad one is reported and skipped, never fatal. */
export function bundledSkills(root = defaultRoot, warn = () => {}) {
  const candidates = [];
  for (const skillName of SKILL_NAMES) {
    const directory = join(root, skillName);
    const locator = join(directory, 'SKILL.md');
    try {
      const { fields } = parseSkill(readFileSync(locator, 'utf8'));
      if (fields.name !== skillName || !SKILL_NAME.test(fields.name)) throw new Error(`name must be "${skillName}"`);
      if (typeof fields.description !== 'string' || !fields.description) throw new Error('description is required');
      candidates.push({
        name: skillName,
        description: fields.description,
        ...fields.whenToUse && { whenToUse: fields.whenToUse },
        invocation: { modelInvocable: true, userInvocable: true },
        provider: PROVIDER,
        source: 'bundled',
        rank: BUNDLED_RANK,
        path: locator,
        resourceBase: { kind: 'directory', path: directory.split(sep).join('/') },
        locator,
      });
    } catch (error) {
      warn(`skipped bundled skill ${skillName}: ${error?.message ?? error}`);
    }
  }
  return candidates;
}

export function createProvider(candidates) {
  return {
    name: PROVIDER,
    list: () => Promise.resolve(candidates),
    async get(candidate, options = {}) {
      const { rank: _rank, locator, ...summary } = candidate;
      // Read the current file on every load, as DSH's own providers do.
      const { content } = parseSkill(await readFile(locator, { encoding: 'utf8', signal: options.signal }));
      return { ...summary, content };
    },
  };
}

const slash = path => path.split(sep).join('/');
const unavailableNote = (name, reason) => `> 注意：${name} 的样式包、模板和脚本${reason ? `下载失败（${reason}）` : '还在后台下载'}，完成前只能用实时工具里不依赖它们的部分；可在插件设置的「Skills」里重试。\n> Note: ${name}'s style packs, templates and scripts are ${reason ? `not downloaded (${reason})` : 'still downloading'}; until then use only what does not need them.\n\n`;

/**
 * The provider of the universal package: each enabled skill at its effective version (description, body and
 * resourceBase from one version, spec §4.5). `get` fetches that version's resources when they are not here yet,
 * waiting at most GET_BUDGET_MS; the rest of the download goes on and `invalidate()` shows it when done.
 */
export function createStoreProvider(store, { invalidate = () => {}, log = () => {}, budgetMs = GET_BUDGET_MS } = {}) {
  const candidate = name => {
    const effective = store.effective(name);
    const { fields } = parseSkill(store.skillText(name, effective));
    if (fields.name !== name || typeof fields.description !== 'string' || !fields.description) throw new Error(`${name}: SKILL.md has no name or description`);
    const locator = effective.dir ? join(effective.dir, 'SKILL.md') : join(store.packageDir(name), 'SKILL.md');
    return { name, description: fields.description, ...fields.whenToUse && { whenToUse: fields.whenToUse }, invocation: { modelInvocable: true, userInvocable: true },
      provider: PROVIDER, source: 'bundled', rank: BUNDLED_RANK, path: locator, version: effective.version,
      ...effective.dir && { resourceBase: { kind: 'directory', path: slash(effective.dir) } }, locator };
  };
  // A download `get()` started tries the first source only; if it (or anything joined to it) fails, one more run with
  // every source follows, so a slow or broken git archive still falls back to Generic Packages.
  const background = name => {
    void store.ensure(name).catch(() => store.ensure(name)).then(() => invalidate(), error => log(`skill ${name}: ${error?.message ?? error}`));
  };
  return {
    name: PROVIDER,
    async list() {
      const out = [];
      for (const name of store.names) {
        if (!store.enabled(name)) continue;
        try {
          const item = candidate(name);
          out.push(item);
          const effective = store.effective(name);
          if (!store.dev && effective.selected && effective.selected !== effective.version) background(name);
        } catch (error) { log(`skipped skill ${name}: ${error?.message ?? error}`); }
      }
      return out;
    },
    async get(found, options = {}) {
      const { rank: _rank, locator, version, ...summary } = found;
      if (summary.resourceBase || store.dev) {
        const { content } = parseSkill(await readFile(locator, { encoding: 'utf8', signal: options.signal }));
        return { ...summary, content };
      }
      try {
        const got = await store.ensure(found.name, { version, budgetMs, firstSourceOnly: true, signal: options.signal });
        invalidate();
        const { content } = parseSkill(await readFile(join(got.dir, 'SKILL.md'), { encoding: 'utf8', signal: options.signal }));
        return { ...summary, resourceBase: { kind: 'directory', path: slash(got.dir) }, content };
      } catch (error) {
        if (options.signal?.aborted) throw error;
        // Answer with the light layer now; the download goes on (or is retried) in the background.
        background(found.name);
        const reason = error?.code === 'still-downloading' ? undefined : error?.message?.slice(0, 160);
        const { content } = parseSkill(await readFile(locator, { encoding: 'utf8', signal: options.signal }));
        return { ...summary, content: unavailableNote(found.name, reason) + content };
      }
    },
  };
}

export function apply(ctx, _config, environment = {}) {
  const logger = ctx.logger?.('officesdk/skills');
  const packageRoot = environment.packageRoot ?? defaultPackageRoot;
  // The repository and packages from before 0.1.0 carry complete skill directories and no catalog: as before.
  if (!environment.store && !existsSync(join(packageRoot, 'skills-catalog.json'))) {
    const candidates = bundledSkills(environment.root ?? defaultRoot, text => logger?.warn?.(text));
    if (!candidates.length) return;
    const provider = createProvider(candidates);
    // The registry ties the registration to this plugin's scope: disabling the component or removing the plugin drops it.
    ctx.skills.registerProvider(() => provider);
    return;
  }
  const store = environment.store ?? skillStoreFor({ packageRoot, dataRoot: environment.dataRoot ?? skillsDataRoot(), log: text => logger?.warn?.(text) });
  ctx.skills.registerProvider(control => {
    const provider = createStoreProvider(store, { invalidate: control.invalidate, log: text => logger?.warn?.(text) });
    // Settings and the remote catalog are shared by every DSH process on this machine: a change anywhere shows here.
    const watched = [store.settingsFile, store.cacheFile];
    const onChange = () => control.invalidate();
    for (const file of watched) watchFile(file, { interval: 5000, persistent: false }, onChange);
    const refresh = () => { void store.refresh().then(changed => { if (changed) control.invalidate(); }, error => logger?.debug?.(`skill catalog: ${error?.message}`)); };
    const first = setTimeout(refresh, environment.refreshDelayMs ?? REFRESH_DELAY_MS);
    const every = setInterval(refresh, REFRESH_EVERY_MS);
    first.unref?.(); every.unref?.();
    store.startLease?.();
    void store.prune?.().catch(() => {});
    control.signal.addEventListener('abort', () => { clearTimeout(first); clearInterval(every); for (const file of watched) unwatchFile(file, onChange); void store.dispose?.(); }, { once: true });
    return provider;
  });
}

/**
 * The Host's data directory (`<DSH home>/plugins/officesdk`). DSH home is `$DSH_HOME`, else `~/.dsh` (dsh-home-paths'
 * rule; not imported, so this module loads without DSH's packages in CI).
 */
const skillsDataRoot = () => join(process.env.DSH_HOME?.trim() || join(homedir(), '.dsh'), 'plugins', 'officesdk');

/**
 * The skill store of the skills component: the same download profile (download-sources.json, spec 261111-npm-publish-prep
 * §6) as the Host's store.
 */
export function skillStoreFor({ packageRoot, dataRoot, log = () => {} }) {
  return new SkillStore({ packageRoot, dataRoot, names: SKILL_NAMES,
    pluginVersion: JSON.parse(readFileSync(join(packageRoot, 'package.json'), 'utf8')).version,
    sources: selectProfile(readDownloadSources(packageRoot, (level, text) => log(text))).skills, log });
}
