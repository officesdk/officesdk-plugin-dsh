import { readFileSync, readdirSync } from 'node:fs';
import { lstat } from 'node:fs/promises';
import { join } from 'node:path';
import { SuiteError } from '../../../core/src/index.mjs';

/**
 * Word documents the model writes live (spec 261105 §3.1). Templates are the bundled shimo-docx skill's
 * `templates/<name>.docx` (scripts/build_live_templates.py): page setup lives in the template because the Writer
 * JS API has none (office-kit K2). Reading, editing and saving use the Host's format-neutral aiRead / aiEdit / aiSave.
 */
export const DEFAULT_TEMPLATE = 'gongwen';
/** Workbooks start from shimo-xlsx's `templates/<name>.xlsx` (scripts/build_template.py). */
export const DEFAULT_SHEET_TEMPLATE = 'blank';
const TEMPLATE_NAME = /^[a-z0-9]+(?:-[a-z0-9]+)*$/;

export function templateNames(root, extension = '.docx') {
  try { return readdirSync(root, { withFileTypes: true }).filter(entry => entry.isFile() && entry.name.endsWith(extension) && TEMPLATE_NAME.test(entry.name.slice(0, -extension.length))).map(entry => entry.name.slice(0, -extension.length)).sort(); }
  catch { return []; }
}

export function loadTemplate(root, name = DEFAULT_TEMPLATE, extension = '.docx') {
  const available = templateNames(root, extension);
  if (typeof name !== 'string' || !available.includes(name)) throw new SuiteError('invalid-edit', `Unknown template ${JSON.stringify(name)}; available: ${available.join(', ') || 'none'}.`);
  try { return readFileSync(join(root, `${name}${extension}`)); }
  catch (error) { throw new SuiteError('invalid-edit', `Template ${name} cannot be read (${error?.name ?? 'error'}).`); }
}

/**
 * AI tool: a new .docx from a template, written where the Session may write (outside the workspace only with
 * `approve`), opened beside the conversation and read. Like presentation_create, the new file is the model's to
 * write: its edits apply without a card each time until the user presses Stop or the budget runs out.
 */
export function createDocument(service, sessionId, args, options) {
  return createFromTemplate(service, sessionId, { ...args, template: args?.template ?? DEFAULT_TEMPLATE }, options, { extension: '.docx', root: service.docTemplateRoot, tool: 'document_create', noun: 'document' });
}

/** AI tool: a new .xlsx from a shimo-xlsx template, generated live the same way (spreadsheet_create). */
export function createSpreadsheet(service, sessionId, args, options) {
  return createFromTemplate(service, sessionId, { ...args, template: args?.template ?? DEFAULT_SHEET_TEMPLATE }, options, { extension: '.xlsx', root: service.sheetTemplateRoot, tool: 'spreadsheet_create', noun: 'spreadsheet' });
}

/** Word template names the shimo-docx skill text may still use; the build ships gongwen alone, so these fall back to it. */
const FALLBACK_WORD_TEMPLATES = Object.freeze(['redhead', 'basic']);

async function createFromTemplate(service, sessionId, { path, template }, { approve, signal } = {}, { extension, root, tool, noun }) {
  const scope = await service.scopeOf(sessionId);
  if (typeof path !== 'string' || !path.toLowerCase().endsWith(extension)) throw new SuiteError('invalid-path', `${tool} makes ${extension} files; give a path ending in ${extension}.`);
  // The universal package downloads the skill's templates on first use (spec 261110-universal-package §4.5).
  if (service.skillRoot) root = await service.skillRoot(extension === '.docx' ? 'shimo-docx' : 'shimo-xlsx', sessionId, 'templates', { signal }) ?? root;
  // A built-in Word template this copy does not ship (the public build has gongwen only, spec 261113-bundle-public-skills):
  // the document starts from the default one and the result says so, so the model can tell the user. A misspelt name still fails.
  let fallback;
  if (extension === '.docx' && typeof template === 'string' && template !== DEFAULT_TEMPLATE && FALLBACK_WORD_TEMPLATES.includes(template)) {
    const available = templateNames(root, extension);
    if (!available.includes(template) && available.includes(DEFAULT_TEMPLATE)) {
      fallback = { requested: template, used: DEFAULT_TEMPLATE, reason: `This build of the plugin does not include the "${template}" template; the document was created from "${DEFAULT_TEMPLATE}" (the same official-document layout${template === 'redhead' ? ', without the red letterhead: write the issuing authority and document number as ordinary paragraphs if needed' : ''}). Tell the user.` };
      template = DEFAULT_TEMPLATE;
    }
  }
  const data = loadTemplate(root, template, extension);
  let created;
  try { created = await service.writer.create(scope, path, data, signal ?? AbortSignal.timeout(30_000)); }
  catch (error) {
    if (error?.code !== 'sandbox-denied' || !error.details?.outside) throw error;
    const outcome = await approve?.(`Create ${path} (outside the workspace) for a new ${noun}`);
    if (outcome !== 'allowed-once') throw new SuiteError('sandbox-denied', `The user did not allow creating ${path} outside the workspace (${outcome ?? 'no approval channel'}). Create it inside the workspace instead.`, { outside: true });
    created = await service.writer.create(scope, path, data, signal ?? AbortSignal.timeout(30_000), { approved: true });
  }
  const real = created.absolutePath;
  const identity = await lstat(real).catch(() => undefined);
  const windowed = Boolean(identity?.isFile() && await service.realPathOf(scope, real) === real);
  if (windowed) service.generations.set(`${sessionId}\0${real}`, { steps: 0, imageBytes: 0, reserved: { steps: 0, imageBytes: 0 }, stopped: false, dev: identity.dev, ino: identity.ino });
  const read = await service.agents.request(sessionId, { kind: 'read', path }, signal);
  return { created: true, template, ...fallback && { templateFallback: fallback }, ...read, generating: windowed };
}

/** What approval and error texts call a file the AI saves. */
export const nounOf = path => /\.docx$/i.test(path ?? '') ? 'document' : /\.xlsx$/i.test(path ?? '') ? 'spreadsheet' : 'presentation';
