#!/usr/bin/env node
/**
 * Checks, in a built package, that a tool error reaches the model as "<code>: <message>".
 *
 * The release bundles Core once into lib/host and once into lib/tools, so the Host's service throws a SuiteError of
 * another class than the one lib/tools knows. Unit tests run from source with one Core and cannot see this; this
 * check loads the two bundles as shipped, makes the Host's own code throw (its exported loadStyle with a style that
 * does not exist) behind presentation_create, and reads the error the tool hands back to the model.
 *
 *   node scripts/verify-tool-errors.mjs <package root containing lib/host and lib/tools>
 *
 * DSH's @deepseek-ai packages are peers a test install omits; verify-tool-errors-hooks.mjs stubs them.
 */
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { register } from 'node:module';
import { pathToFileURL } from 'node:url';

const root = process.argv[2];
if (!root) throw new Error('usage: node scripts/verify-tool-errors.mjs <package root>');
const bundles = { host: path.resolve(root, 'lib/host/index.mjs'), tools: path.resolve(root, 'lib/tools/index.mjs') };
const IDENTIFIER = /^[A-Za-z_$][\w$]*$/;

const exports = {};
for (const [label, file] of Object.entries(bundles)) {
  const text = await fs.readFile(file, 'utf8');
  // The check proves something only while each bundle carries its own SuiteError class (esbuild keepNames writes
  // `SuiteError=class`). If a build ever shares Core, this check has to be rewritten rather than pass silently.
  const classes = text.match(/\bSuiteError=class\b|\bclass SuiteError\b/g)?.length ?? 0;
  if (classes !== 1) throw new Error(`lib/${label}/index.mjs has ${classes} SuiteError classes, expected 1; the bundle layout changed, so update scripts/verify-tool-errors.mjs`);
  // Static imports start a statement; a bare `import` inside a string literal (`"import"`) does not.
  for (const [, clause, specifier] of text.matchAll(/(?:^|[;}\n])\s*import\s*([^"';()]*?)\s*(?:from\s*)?["']([^"']+)["']/g)) {
    if (specifier.startsWith('node:') || specifier.startsWith('.') || specifier.startsWith('/')) continue;
    if (!specifier.startsWith('@deepseek-ai/')) throw new Error(`lib/${label}/index.mjs imports ${specifier}; only node: and @deepseek-ai/ packages can be stubbed, so update scripts/verify-tool-errors-hooks.mjs`);
    const names = new Set(exports[specifier] ?? []);
    const named = /\{([^}]*)\}/.exec(clause)?.[1] ?? '';
    const head = clause.replace(/\{[^}]*\}/, '').replace(/,/g, ' ').trim();
    if (head && !head.startsWith('*')) names.add('default');
    for (const part of named.split(',').map(entry => entry.trim()).filter(Boolean)) names.add(part.split(/\s+as\s+/)[0]);
    for (const name of names) if (name !== 'default' && !IDENTIFIER.test(name)) throw new Error(`lib/${label}/index.mjs imports an unexpected name ${JSON.stringify(name)} from ${specifier}`);
    exports[specifier] = [...names];
  }
}
if (!exports['@deepseek-ai/dsh-tools']?.includes('defineTool')) throw new Error('lib/tools/index.mjs no longer imports defineTool from @deepseek-ai/dsh-tools; update scripts/verify-tool-errors.mjs');
register(pathToFileURL(path.join(import.meta.dirname, 'verify-tool-errors-hooks.mjs')), { data: { exports } });

const host = await import(pathToFileURL(bundles.host));
const tools = await import(pathToFileURL(bundles.tools));
if (typeof host.loadStyle !== 'function') throw new Error('lib/host/index.mjs no longer exports loadStyle; pick another Host function that throws a SuiteError');

const registered = new Map();
const missing = path.join(await fs.mkdtemp(path.join(os.tmpdir(), 'officesdk-tool-errors-')), 'no-styles');
try {
  // The tools' service is the Host's: its failure is a SuiteError made by lib/host's copy of Core.
  tools.apply({ tools: { register: tool => registered.set(tool.name, tool) }, shimoPreview: { aiCreate: async () => host.loadStyle(missing, 'no-such-style') } });
  const create = registered.get('presentation_create');
  if (!create) throw new Error('lib/tools registered no presentation_create tool');
  let message;
  try { await create.execute({ path: 'check.pptx', style: 'no-such-style' }, { agent: { session: { id: 'verify-tool-errors' } }, signal: new AbortController().signal }); }
  catch (error) { message = String(error?.message); }
  if (message === undefined) throw new Error('presentation_create did not fail for an unknown style');
  if (!message.startsWith('invalid-edit: ')) throw new Error(`a Host SuiteError reached the model without its code prefix: ${JSON.stringify(message)} (expected "invalid-edit: ...")`);
  console.log(`Verified tool error prefix in the package: ${JSON.stringify(message)}`);
} finally {
  await fs.rm(path.dirname(missing), { recursive: true, force: true });
}
