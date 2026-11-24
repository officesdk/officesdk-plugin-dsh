#!/usr/bin/env node
// Mechanical layering check for packages/: dependency direction and host neutrality.
import { existsSync, readFileSync, readdirSync, statSync } from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '../..');
const RULES = {
  core: { allow: [], forbid: [/^@deepseek-ai\//, /^@officesdk\/(?!core)/] },
  'kit-host': { allow: [], forbid: [/^@deepseek-ai\//, /^@officesdk\//] },
  'dsh-adapter': { allow: ['core', 'kit-host'], forbid: [] },
};
const problems = [];
const files = dir => readdirSync(dir).flatMap(name => {
  const file = path.join(dir, name);
  if (['node_modules', 'dist'].includes(name)) return [];
  return statSync(file).isDirectory() ? files(file) : /\.(mjs|js|ts)$/.test(name) ? [file] : [];
});
for (const [name, rule] of Object.entries(RULES)) {
  const src = path.join(root, 'packages', name, 'src');
  if (!existsSync(src)) continue;
  for (const file of files(src)) {
    const text = readFileSync(file, 'utf8');
    const browser = file.includes(`${path.sep}browser${path.sep}`);
    for (const [, specifier] of text.matchAll(/(?:import|export)\s[^'"]*?from\s*['"]([^'"]+)['"]|import\(\s*['"]([^'"]+)['"]\s*\)/g)) {
      if (!specifier) continue;
      const where = path.relative(root, file);
      if (rule.forbid.some(pattern => pattern.test(specifier))) problems.push(`${where}: forbidden import ${specifier}`);
      if (browser && specifier.startsWith('node:')) problems.push(`${where}: browser code imports ${specifier}`);
      if (specifier.startsWith('.')) {
        const target = path.relative(path.join(root, 'packages'), path.resolve(path.dirname(file), specifier)).split(path.sep)[0];
        if (target !== name && !rule.allow.includes(target) && !target.startsWith('..')) problems.push(`${where}: ${name} may not depend on ${target}`);
        if (target.startsWith('..')) problems.push(`${where}: imports outside packages/ (${specifier})`);
      }
    }
  }
}
if (problems.length) { console.error(problems.join('\n')); process.exit(1); }
console.log('check-deps: ok');
