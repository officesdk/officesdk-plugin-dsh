import { finishProjectTable } from './project-plan-stage.mjs';
import { blankExistingDeck } from './existing-deck.mjs';
import { existsSync, readFileSync, readdirSync } from 'node:fs';
import { lstat, realpath } from 'node:fs/promises';
import { createHash, randomUUID } from 'node:crypto';
import { fileURLToPath } from 'node:url';
import { basename, join, resolve as resolvePath } from 'node:path';
import { userInfo } from 'node:os';
import z from '@deepseek-ai/schemastery';
import { Remote, RemoteError, TypertRemoteService } from '@deepseek-ai/dsh-typert-protocol';
import { withFileLock, writeFileAtomic } from '@deepseek-ai/dsh-atomic-write';
import { writableRoots } from '@deepseek-ai/dsh-sandbox';
import { dshHomePath } from '@deepseek-ai/dsh-home-paths';
import { DEFAULT_LIMITS, NEW_SLIDES_MESSAGE, PPTX_LIMITS, SuiteError, checkRewriteText, checkRewriteValues, countNewSlides, SuiteService, createHttpHandler, parallelEditMessage, selectionMessage, validateSelectionRequest } from '../../../core/src/index.mjs';
import { KitDocuments, KitProcess, createKitHttpHandler } from '../../../kit-host/src/index.mjs';
import { AgentHub } from './agent.mjs';
import { createDocument, createSpreadsheet, nounOf } from './document.mjs';
import { createAuthorizer } from './authorizer.mjs';
import { fromKit, principalKey, toRemote } from './errors.mjs';
import { createSource } from './source.mjs';
import { liveLayouts, readLayouts } from './style-layouts.mjs';
import { createWriter } from './writer.mjs';
import { PluginUpdater, redactCredentials, updateWithKit } from './updater.mjs';
import { OfficeKitRuntime } from './fetch/runtime.mjs';
import { SkillStore } from '../skills/store.mjs';
import { SKILL_NAMES } from '../skills/index.mjs';
import { readManifest } from './fetch/office-kit.mjs';
import { readDownloadSources, selectProfile } from './fetch/download-sources.mjs';
import { NO_CONVERTER, resolveKit, unsupportedError } from './kit-platform.mjs';
import { WhatsNew } from './whats-new.mjs';
import { ReviewStore } from './review-store.mjs';
import { StyleMemory, styleCard, styleCards } from './style-catalog.mjs';
import { StyleSuggestions, suggestionContext } from './style-suggest.mjs';
import { BUILTIN_TEMPLATES, PICK_KINDS, TemplatePicks, pickContext } from './template-pick.mjs';
import { PARALLEL_EDIT_COMMAND } from '../parallel-message.mjs';
import { generateSelectionReplacement } from './parallel-rewrite.mjs';
import { SelectionStore } from './selection.mjs';

const packageRoot = fileURLToPath(new URL('../../', import.meta.url));
const METHODS = ['open', 'engineStatus', 'skillsStatus', 'skillSetEnabled', 'skillPrefetch', 'skillClearCache', 'skillsCheckUpdates', 'release', 'renew', 'access', 'rebind', 'save', 'exportTo', 'probe', 'browse', 'resolveDiskChange', 'listDrafts', 'discardDraft', 'exportDraft', 'agentNext', 'agentAlive', 'agentReply', 'agentStop', 'updateStatus', 'updatePlugin', 'listStyles', 'rememberStyle', 'selectTemplate', 'templatePick', 'styleSuggestion', 'dismissStyleSuggestion', 'whatsNew', 'whatsNewSeen', 'selectionCreate', 'selectionUpdate', 'selectionParallel'];
const initializers = [];
for (const name of METHODS) Remote(undefined, {
  kind: 'method', name, private: false, static: false,
  addInitializer: initializer => initializers.push(initializer),
});
const scalar = Object.fromEntries(Object.entries(DEFAULT_LIMITS).filter(([, value]) => typeof value === 'number'));
export const Config = z.object(Object.fromEntries(Object.entries(scalar).map(([key, value]) => [key, z.number().step(1).min(1).max(Number.MAX_SAFE_INTEGER - 1).default(value)])));

// Only acknowledged batches advance a fixed template's restored sorter.
const slideSnapshot = generation => generation?.fixedSlides && !generation.stopped
  ? { done: [...generation.fixedSlides.done], partial: [...generation.fixedSlides.partial] } : undefined;

const readJson = path => JSON.parse(readFileSync(path, 'utf8'));
/**
 * Style packs of the bundled shimo-pptx skill (`skills/shimo-pptx/styles/<name>/`): `style.json` and the
 * `template.pptx` its scripts/build_template.py makes from it. presentation_create starts from a pack's template;
 * adding a style is adding a pack, the Host lists whatever packs are there.
 */
export const DEFAULT_STYLE = 'default';
const STYLE_CACHE_MS = 10_000;
const STYLE_NAME = /^[a-z0-9]+(?:-[a-z0-9]+)*$/;
/**
 * A pack is a folder with all three — loadStyle needs layouts.json too, so a pack without it is not offered (the picker
 * would otherwise list and remember a style presentation_create then rejects). scripts/brand_deck.py uses the same rule.
 */
const PACK_FILES = ['style.json', 'template.pptx', 'layouts.json'];
export function styleNames(root) {
  try { return readdirSync(root, { withFileTypes: true }).filter(entry => entry.isDirectory() && STYLE_NAME.test(entry.name) && PACK_FILES.every(file => existsSync(join(root, entry.name, file)))).map(entry => entry.name).sort(); }
  catch { return []; }
}
/**
 * The pack a user-facing value names: its pack name, else its display name (Chinese or English) or an alias, ignoring
 * case and surrounding spaces. The inspiration library writes display names into the message, so the model may pass
 * one. Undefined when nothing matches.
 */
export function resolveStyle(root, value, ids = styleNames(root)) {
  if (typeof value !== 'string' || !value.trim()) return undefined;
  const wanted = value.trim();
  if (ids.includes(wanted)) return wanted;
  const key = wanted.toLowerCase();
  for (const id of ids) {
    let card;
    try { card = styleCard(root, id, { thumb: false }); } catch { continue; }
    if ([card.name?.zh, card.name?.en, ...(card.aliases ?? [])].some(name => typeof name === 'string' && name.trim().toLowerCase() === key)) return id;
  }
  return undefined;
}
/** The packs for an error message: `id（中文名）` each, so a model that passed a name can see which id is meant. */
export function styleList(root, ids = styleNames(root)) {
  if (!ids.length) return 'none';
  return ids.map(id => { try { const name = styleCard(root, id, { thumb: false }).name?.zh; return name && name !== id ? `${id}（${name}）` : id; } catch { return id; } }).join(', ');
}
/** A pack's template bytes and the values the model needs to build pages in it (colours, fonts, sizes in pt). */
export function loadStyle(root, name = DEFAULT_STYLE) {
  const available = styleNames(root);
  if (typeof name !== 'string' || !available.includes(name)) throw new SuiteError('invalid-edit', `Unknown style ${JSON.stringify(name)}; available: ${styleList(root, available)}.`);
  let style, template;
  try { style = readJson(join(root, name, 'style.json')); template = readFileSync(join(root, name, 'template.pptx')); }
  catch (error) { throw new SuiteError('invalid-edit', `Style ${name} cannot be read (${error?.name ?? 'error'}); its style.json or template.pptx is broken.`); }
  const palette = style?.palettes?.[style.default_palette];
  if (!palette) throw new SuiteError('invalid-edit', `Style ${name} has no default palette.`);
  // The live layouts and slot geometry (pt) come from the pack's layouts.json, the same file the offline renderer draws.
  let layouts;
  try { layouts = liveLayouts(readLayouts(root, name)); }
  catch (error) { throw new SuiteError('invalid-edit', `Style ${name} has a broken layouts.json: ${error.message}`); }
  const { primary, onPrimary, ink, paper, secondary, lineOnPaper } = palette;
  return {
    template,
    summary: { name, palette: { primary, onPrimary, ink, paper, secondary, lineOnPaper }, fonts: { cjk: style.fonts?.cjk, latin: style.fonts?.latin }, ...(style.template?.sizes_pt !== undefined && { sizesPt:style.template.sizes_pt }), layouts, ...(style.fixed_deck && {fixedDeck:style.fixed_deck}) },
  };
}
/**
 * A file presentation_create made is the model's own while it generates it: its edits apply without a card per
 * batch, within this budget per file. Past it, or once the user stops it, each batch asks again.
 */
export const GENERATION_LIMITS = Object.freeze({ maxSteps: 400, maxImageBytes: 50 * 1024 * 1024 });
const IMAGE_SIGNATURES = [['image/png', [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]], ['image/jpeg', [0xff, 0xd8, 0xff]]];
const imageTypeOf = bytes => IMAGE_SIGNATURES.find(([, magic]) => magic.every((value, index) => bytes[index] === value))?.[0];
// `preview` is the Client's fallback when the Kit editor cannot mount (spec §13.4); this package ships no converter, so it reports Core as unavailable.
const MODES = new Set(['view', 'edit', 'preview']);
const KIT_FORMATS = /\.(docx|xlsx|pptx)$/i;
/**
 * The author the editors record (Writer comments, Sheet edit mode), per DSH principal: the owner is this
 * machine's account, as Word names authors; any other principal gets a stable anonymous identity.
 */
function editorUser(principal) {
  if (principal === 'owner') { let name = 'DSH'; try { name = userInfo().username || name; } catch {} return { uid: `local:${name}`, name }; }
  return { uid: `guest:${createHash('sha256').update(String(principal)).digest('hex').slice(0, 16)}`, name: 'Guest' };
}

/**
 * The Kit page's static roots. `resources` (office-kit's SDKs) is read per request: with the universal package it is
 * the downloaded editor resources once the engine is ready, the package's own copy (offline package) before that.
 */
function staticRootsOf(packageRoot, officeKit) {
  const bundled = join(packageRoot, 'resources', 'office-kit');
  return { page: join(packageRoot, 'resources', 'kit-page'), editor: join(packageRoot, 'resources', 'editor'), fonts: join(packageRoot, 'resources', 'host-fonts'),
    get resources() { return officeKit?.editorDir ?? bundled; } };
}

/** Host-owned service; reads go through WorkspaceFiles, writes through the Session's sandbox policy. */
export default class ShimoPreview extends TypertRemoteService {
  static inject = ['workspaceFiles', 'webServer', 'connection', 'typert'];
  static Config = Config;
  constructor(ctx, config = {}, environment = {}) {
    super(ctx, 'shimoPreview');
    for (const initialize of initializers) initialize.call(this);
    // Office documents open through office-kit's Kit only: the package ships no converter of its own, so Core converts
    // nothing (tests may inject one to exercise Core's service).
    const converter = environment.converter ?? NO_CONVERTER;
    // ~/.dsh/plugins/officesdk.
    const dataRoot = environment.dataRoot ?? join(dshHomePath('plugins'), 'officesdk');
    this.dataRoot = dataRoot;
    this.updater = environment.updater ?? new PluginUpdater({ packageRoot, dataRoot });
    this.whatsNewNotes = environment.whatsNew ?? new WhatsNew({ packageRoot, file: join(dataRoot, 'whats-new-seen.json'), atomic: environment.atomic ?? { withFileLock, writeFileAtomic } });
    this.writer = environment.writer ?? createWriter({ ctx, workspaceFiles: ctx.workspaceFiles, dataRoot, writableRoots: environment.writableRoots ?? writableRoots, atomic: environment.atomic ?? { withFileLock, writeFileAtomic } });
    // Write capability is decided per call; a composition without sandbox/sessions/local fs still previews.
    const writer = this.writer;
    const source = createSource(ctx.workspaceFiles);
    this.suite = new SuiteService({
      source, writer, authorizer: createAuthorizer(ctx), converter,
      runtimeRoot: environment.runtimeRoot ?? join(packageRoot, 'resources', 'runtime'), tempRoot: environment.tempRoot, dataRoot, routePrefix: '/shimo', limits: config,
    });
    Object.defineProperty(this.suite, 'hostCapabilities', { get: () => writer.canWrite() ? {} : { edit: false, save: false } });
    const admit = req => {
      const admission = ctx.connection.admit(req);
      return 'rejection' in admission ? admission : { peer: this.peerOf(admission.peer) };
    };
    const allowedOrigin = (origin, req) => environment.allowedOrigin?.(origin, req) ?? sameOrigin(origin, req);
    const coreHandler = createHttpHandler({ service: this.suite, admit, allowedOrigin });
    // DOCX, XLSX and PPTX go through office-kit's Kit (spec §12, §13).
    const kitRoot = environment.kitRoot ?? join(packageRoot, 'resources', 'kit');
    const logger = ctx.logger?.('officesdk/kit');
    // The universal package (spec 261110-universal-package) carries kit-manifest.json instead of a Kit: the engine is
    // downloaded into the data directory when first needed (an offline package carries it and uses it in place).
    const manifest = environment.kitDocuments || environment.kitProcess || environment.kitRoot ? undefined : environment.kitManifest ?? readManifest(packageRoot);
    // Which sources each download tries (download-sources.json).
    this.sources = environment.sources ?? selectProfile(readDownloadSources(packageRoot, (level, text) => logger?.[level]?.(text)));
    if (manifest) {
      this.officeKit = environment.officeKit ?? new OfficeKitRuntime({ manifest, dataRoot, packageRoot, host: environment.host, profile: this.sources,
        mirror: process.env.OFFICESDK_KIT_MIRROR || undefined, npm: process.env.OFFICESDK_EDITOR_NPM || undefined, log: (level, text) => logger?.[level]?.(text) });
    }
    // This machine's Kit executable (`.exe` on Windows); a missing or other platform's Kit is never launched.
    const located = environment.kitDocuments || environment.kitProcess ? {}
      : this.officeKit ? (this.officeKit.unsupported ? { unsupported: this.officeKit.unsupported } : { binary: () => this.officeKit.binary() })
      : resolveKit(kitRoot, environment.host);
    this.kitUnsupported = located.unsupported;
    // Leases tell the Client which path rules apply (`C:\\…` on win32); the Client never guesses the platform.
    this.platform = environment.host?.platform ?? process.platform;
    if (located.unsupported) logger?.warn?.(`document service unavailable on ${located.unsupported.platform}: ${located.unsupported.reason}${located.unsupported.file ? ` (${located.unsupported.file})` : ''}`);
    let kitHandler;
    if (!located.unsupported) {
      this.kit = environment.kitProcess ?? new KitProcess({ binary: located.binary, work: join(dataRoot, 'kit') });
      this.kit.on?.('log', text => logger?.debug?.(text.trimEnd()));
      this.kit.on?.('down', ({ error }) => logger?.warn?.(`document service stopped: ${error?.message}`));
      // A Kit that started and opened its work directory moves the downgrade record up (§4.3); a failed start never does.
      if (this.officeKit) this.kit.on?.('ready', () => { void this.officeKit.markStarted(); });
      this.kitDocs = environment.kitDocuments ?? new KitDocuments({
        kit: this.kit, exportRoot: join(dataRoot, 'kit-exports'),
        authorize: (peer, scope, signal, { action, path }) => this.suite.authorize(peer, scope, signal, action, path ?? '', 0),
        stat: (scope, path, signal) => source.stat(scope, path, signal),
        writeAccess: (scope, realPath) => writer.accessFor(scope, realPath),
        warn: text => logger?.warn?.(text),
      });
      kitHandler = createKitHttpHandler({ documents: this.kitDocs, admit, allowedOrigin, prefix: '/shimo',
        reviews:new ReviewStore(join(dataRoot,'pending-reviews'),environment.atomic ?? {withFileLock,writeFileAtomic}),
        staticRoots: environment.kitStaticRoots ?? staticRootsOf(packageRoot, this.officeKit),
        meta: view => ({ v: 1, id: view.id, suite: view.suite, mode: view.writer ? 'edit' : 'view', name: basename(view.path), diskChanged: view.diskChanged, user: editorUser(view.principal) }) });
      // Start the engine download soon after start-up, and drop other versions not used for 30 days.
      this.officeKit?.prefetch();
      void this.officeKit?.prune({ kitWork: join(dataRoot, 'kit') }).catch(() => {});
      this.kitSweeper = setInterval(() => { void this.kitDocs.expire().catch(() => {}); }, 60_000);
      this.kitSweeper.unref?.();
    }
    const handler = async (req, res) => { if (kitHandler && await kitHandler(req, res)) return; return coreHandler(req, res); };
    this.agents = new AgentHub({ platform: this.platform, ...environment.agentHub });
    // Skill resources (style packs, templates) download on first use in the universal package (spec 261110-universal-package
    // §4.5); the roots below then follow the skill's effective version. The repository and tests keep fixed directories.
    this.skills = environment.skills ?? (environment.styleRoot || !existsSync(join(packageRoot, 'skills-catalog.json')) ? undefined
      : new SkillStore({ packageRoot, dataRoot, names: SKILL_NAMES, pluginVersion: this.updater.current, sources: this.sources.skills, log: text => ctx.logger?.('officesdk/skills')?.warn?.(text) }));
    const skillDir = (name, fallback) => () => (this.skills ? join(this.skills.resolveNow(name) ?? join(packageRoot, 'skills', name), fallback) : undefined);
    if (this.skills) {
      for (const [field, name, sub] of [['styleRoot', 'shimo-pptx', 'styles'], ['docTemplateRoot', 'shimo-docx', 'templates'], ['sheetTemplateRoot', 'shimo-xlsx', 'templates']]) {
        Object.defineProperty(this, field, { get: skillDir(name, sub), configurable: true });
      }
    } else this.styleRoot = environment.styleRoot ?? join(packageRoot, 'skills', 'shimo-pptx', 'styles');
    const styleLogger = ctx.logger?.('officesdk/styles');
    this.styleMemory = environment.styleMemory ?? new StyleMemory({ file: join(dataRoot, 'style-defaults.json'), atomic: environment.atomic ?? { withFileLock, writeFileAtomic }, warn: text => styleLogger?.warn?.(text) });
    this.styleCache = undefined;
    if (!this.skills) {
      this.docTemplateRoot = environment.docTemplateRoot ?? join(packageRoot, 'skills', 'shimo-docx', 'templates');
      this.sheetTemplateRoot = environment.sheetTemplateRoot ?? join(packageRoot, 'skills', 'shimo-xlsx', 'templates');
    }
    this.styleSuggestions = environment.styleSuggestions ?? new StyleSuggestions();
    this.templatePicks = environment.templatePicks ?? new TemplatePicks();
    /** Selections the user asked to rewrite, until selection_rewrite uses them (spec 261109-selection-ai-edit). */
    this.selections = environment.selections ?? new SelectionStore();
    /** Generation windows: `${sessionId}\0${realPath}` → { steps, imageBytes, stopped }. */
    this.generations = new Map();
    const unregister = ctx.webServer.register({ kind: 'prefix', path: '/shimo', handler });
    this.skills?.startLease?.();
    ctx.effect(() => async () => { await this.skills?.dispose?.(); this.officeKit?.dispose(); unregister(); clearInterval(this.kitSweeper); this.agents.dispose(); await this.kitDocs?.dispose(); await this.suite.dispose(); });
  }

  /** Core sees a stable principal key for drafts and the original peer for remoteAccess. */
  peerOf(remote) {
    if (!remote?.id) throw new SuiteError('forbidden', 'Documents require an authenticated Remote invocation.');
    return { id: remote.id, principal: principalKey(remote.principal), signal: remote.signal, remote };
  }
  remotePeer() { return this.peerOf(this.ctx.invocation?.peer); }
  async remote(operation, signal) {
    try { return await operation(this.remotePeer()); }
    catch (error) { throw toRemote(error, signal, RemoteError); }
  }

  /** Parameter names are the installed Typert source-mode lookup contract; keep them unmangled. */
  async open(workspaceFileScope, path, revision, mode, signal) {
    return this.remote(async peer => {
      const selected = mode ?? 'view';
      if (!MODES.has(selected)) throw new SuiteError('unsupported-operation', 'Mode must be view or edit.');
      // No Kit can open it here: say so plainly instead of a generic failure (spec §4.1).
      if (!this.kitDocs && KIT_FORMATS.test(path ?? '')) throw unsupportedError(this.kitUnsupported ?? { platform: 'unknown', reason: 'missing', file: 'offline-kit-runtime' });
      if (selected === 'preview') return this.suite.open(workspaceFileScope, path, revision, signal, peer, { mode: 'view' });
      if (this.kitDocs && KIT_FORMATS.test(path ?? '')) {
        try {
          const view = await this.kitDocs.open(workspaceFileScope, path, peer, { mode: selected, signal });
          const handle = this.kitHandle(view, revision, await this.kitAccess(view));
          if (view.suite === 'presentation' && view.mode === 'edit') {
            const progress = slideSnapshot(await this.generationOf(workspaceFileScope.sessionId, workspaceFileScope, path));
            if (progress) handle.slideProgress = progress;
          }
          return handle;
        }
        catch (error) { throw fromKit(error); }
      }
      // Core edits presentations only (§11); its DOCX/XLSX stay the read-only preview they were.
      return this.suite.open(workspaceFileScope, path, revision, signal, peer, { mode: /\.pptx$/i.test(path ?? '') ? selected : 'view' });
    }, signal);
  }
  /** A Kit view as the Client's lease: same fields, `backend: 'kit'`, and the editor page as entry. */
  kitHandle(view, revision, access) {
    const editing = view.mode === 'edit';
    return { id: view.id, backend: 'kit', format: { document: 'docx', sheet: 'xlsx', presentation: 'pptx' }[view.suite], mode: view.mode, revision, version: view.version, editRevision: 0, unsaved: view.unsaved, diskChanged: view.diskChanged, access,
      capabilities: { import: true, preview: true, edit: editing, save: editing, export: true }, expiresInMs: this.kitDocs.viewTtlMs, platform: this.platform,
      entry: `/shimo/kit/static/page/index.html?view=${encodeURIComponent(view.id)}` };
  }
  /** Whether the Session may write an editing view's file ('allowed' | 'outside' | 'denied'); unknown when it cannot tell. */
  async kitAccess(view) {
    // `open` answers with a summary; the scope and real path live on the view Kit documents keep.
    const kept = this.kitDocs.views.get(view.id);
    if (view.mode !== 'edit' || !kept) return undefined;
    try { return await this.writer.accessFor(kept.scope, kept.realPath); } catch { return undefined; }
  }
  isKit(id) { return Boolean(this.kitDocs?.views.has(id)); }
  /** `autosave: false` when the tab falls back to a preview (spec §13.4): the Kit view closes without saving. */
  async release(workspaceFileScope, id, autosave, signal) {
    return this.remote(peer => this.isKit(id) ? this.kitDocs.close(id, peer, { autosave: autosave !== false }) : this.suite.release(workspaceFileScope, id, signal, peer), signal);
  }
  async renew(workspaceFileScope, id, signal) {
    return this.remote(peer => {
      if (!this.isKit(id)) return this.suite.renew(workspaceFileScope, id, signal, peer);
      try { this.kitDocs.view(id, peer); } catch (error) { throw fromKit(error); }
      return { expiresInMs: this.kitDocs.viewTtlMs };
    }, signal);
  }
  /**
   * Whether the Session may write an editing Kit view's file now ('allowed' | 'outside' | 'denied'): the conversation's
   * permission can change while the tab is open. `{}` for other documents or when it cannot tell. Writes nothing.
   */
  async access(workspaceFileScope, id, signal) {
    return this.remote(async peer => {
      if (!this.isKit(id)) return {};
      let view;
      try { view = this.kitDocs.view(id, peer); } catch (error) { throw fromKit(error); }
      // Kit keeps `writer` on its own view; only the editing tab may write.
      const access = await this.kitAccess({ id: view.id, mode: view.writer ? 'edit' : 'view' });
      return access ? { access } : {};
    }, signal);
  }
  async rebind(workspaceFileScope, id, revision, diverged, signal) {
    return this.remote(async peer => {
      if (!this.isKit(id)) return this.suite.rebind(workspaceFileScope, id, revision, { diverged: diverged === true }, signal, peer);
      // The file is still what this view opened or saved, or the tab reopens (source-changed).
      try { return { revision, state: 'active', ...(await this.kitDocs.checkDisk(id, peer, signal)) }; }
      catch (error) { throw fromKit(error); }
    }, signal);
  }
  /**
   * `approved`: the owner confirmed in the panel that this file may be written outside the workspace.
   * `overwrite`: the user chose to replace a file that changed on disk (Kit views only).
   */
  async save(workspaceFileScope, id, editRevision, approved, overwrite, signal) {
    return this.remote(async peer => {
      if (approved === true && peer.principal !== 'owner') throw new SuiteError('forbidden', 'Only the owner can approve writing outside the workspace.');
      if (this.isKit(id)) {
        try { return { ...(await this.kitDocs.save(id, peer, { approved: approved === true, overwrite: overwrite === true, signal })), editRevision }; }
        catch (error) { throw fromKit(error); }
      }
      return this.suite.save(workspaceFileScope, id, { editRevision, approved: approved === true }, signal, peer);
    }, signal);
  }
  async exportTo(workspaceFileScope, id, editRevision, target, signal) {
    return this.remote(async peer => {
      // Save As: the owner's choice of place allows writing outside the workspace, as approving a save does.
      if (target?.approved === true && peer.principal !== 'owner') throw new SuiteError('forbidden', 'Only the owner can approve writing outside the workspace.');
      if (!this.isKit(id)) return this.suite.exportTo(workspaceFileScope, id, { editRevision, target }, signal, peer);
      try {
        if (target === 'download') {
          const { token, name } = await this.kitDocs.exportDownload(id, peer, { signal });
          return { url: `/shimo/kit/v/${encodeURIComponent(id)}/download/${token}`, name };
        }
        const { file } = await this.kitDocs.export(id, peer, { target, signal });
        return { path: file };
      } catch (error) { throw fromKit(error); }
    }, signal);
  }
  /**
   * A picked file or folder before the tab opens or writes it: canonical path, real path, and for a folder whether
   * this Session may write there. Same read authorization as `open`; no file content.
   */
  async probe(workspaceFileScope, path, kind, signal) {
    return this.remote(async peer => {
      if (!this.kitDocs) throw new SuiteError('unsupported-operation', 'Picking files needs the document service.');
      try { return await this.kitDocs.probe(workspaceFileScope, path, peer, { kind: kind === 'folder' ? 'folder' : 'file', signal }); }
      catch (error) { throw fromKit(error); }
    }, signal);
  }
  /** One workspace folder's subfolders and Office files, for the picker used where the desktop's own is missing. */
  async browse(workspaceFileScope, path, signal) {
    return this.remote(async peer => {
      if (!this.kitDocs) throw new SuiteError('unsupported-operation', 'Browsing needs the document service.');
      try { return await this.kitDocs.browse(workspaceFileScope, path ?? '', peer, { signal }); }
      catch (error) { throw fromKit(error); }
    }, signal);
  }
  /** The disk copy changed while edits were kept: `useDisk` replaces them (Kit `import`), otherwise keep editing. */
  async resolveDiskChange(workspaceFileScope, id, useDisk, signal) {
    return this.remote(async peer => {
      if (!this.isKit(id)) throw new SuiteError('unsupported-operation', 'This document has no disk change to resolve.');
      try { return await this.kitDocs.resolveDiskChange(id, peer, { useDisk: useDisk === true, signal }); }
      catch (error) { throw fromKit(error); }
    }, signal);
  }
  /**
   * Whether the install branch has a newer release; `installed` is a newer copy waiting for a DSH restart.
   * `canUpdate`: whether this caller may install it (owner only, as updatePlugin checks).
   */
  /** Skills settings (spec 261110-universal-package §4.5): one row per skill; only the owner may change anything. */
  async skillsStatus(workspaceFileScope, signal) {
    return this.remote(async peer => ({ skills: this.skills?.status() ?? [], managed: Boolean(this.skills), canChange: peer.principal === 'owner' }), signal);
  }
  skillOwner(peer, name) {
    if (peer.principal !== 'owner') throw new SuiteError('forbidden', 'Only the owner changes skill settings.');
    if (!this.skills) throw new SuiteError('unsupported-operation', 'This copy carries its skills complete; there is nothing to download or clear.');
    if (name !== undefined && !this.skillNames().includes(name)) throw new SuiteError('invalid-edit', `Unknown skill ${JSON.stringify(name)}.`);
  }
  /** The skills this copy has: every one in the repository, those its catalog lists in a package (the public build ships two). */
  skillNames() { return this.skills?.names ?? SKILL_NAMES; }
  /** Whether this copy includes the named skill. */
  skillIncluded(name) { return this.skillNames().includes(name); }
  /**
   * Built-in Word and Excel template ids this copy can create from, by pick kind, as its skill catalog lists them; undefined
   * without a packaged catalog (the repository: every built-in template). A skill the package lacks lists none.
   */
  templateIds() {
    if (!this.skills) return undefined;
    const ids = name => (this.skillIncluded(name) ? [...this.skills.effective(name).entry?.templates ?? []] : []);
    return { word: ids('shimo-docx'), excel: ids('shimo-xlsx') };
  }
  async skillSetEnabled(workspaceFileScope, name, enabled, signal) {
    return this.remote(async peer => {
      if (peer.principal !== 'owner') throw new SuiteError('forbidden', 'Only the owner changes skill settings.');
      if (!this.skillNames().includes(name)) throw new SuiteError('invalid-edit', `Unknown skill ${JSON.stringify(name)}.`);
      const store = this.skills ?? new SkillStore({ packageRoot, dataRoot: this.dataRoot, names: SKILL_NAMES, pluginVersion: this.updater.current, sources: this.sources?.skills });
      await store.setEnabled(name, enabled === true);
      return { skills: (this.skills ?? store).status() };
    }, signal);
  }
  /** Starts the download of one skill (or all, `name` null) in the background; the panel polls skillsStatus. */
  async skillPrefetch(workspaceFileScope, name, signal) {
    return this.remote(async peer => {
      this.skillOwner(peer, name ?? undefined);
      for (const skill of name ? [name] : this.skills.names) void this.skills.ensure(skill).catch(error => this.ctx.logger?.('officesdk/skills')?.warn?.(`${skill}: ${error?.message}`));
      return { skills: this.skills.status() };
    }, signal);
  }
  async skillClearCache(workspaceFileScope, name, signal) {
    return this.remote(async peer => {
      this.skillOwner(peer, name);
      const removed = await this.skills.clear(name);
      return { removed: removed.length, skills: this.skills.status() };
    }, signal);
  }
  async skillsCheckUpdates(workspaceFileScope, signal) {
    return this.remote(async peer => {
      this.skillOwner(peer);
      let changed = false;
      try { changed = await this.skills.refresh(); }
      catch (error) { throw new SuiteError('unavailable', `The skill catalog could not be read (${redactCredentials(error?.message)}).`); }
      if (changed) for (const skill of this.skills.names) void this.skills.ensure(skill).catch(() => {});
      return { changed, skills: this.skills.status() };
    }, signal);
  }
  /** The document engine download, for the Client's placeholder while a document waits (spec 261110-universal-package §4.2 step 7). */
  async engineStatus(workspaceFileScope, signal) {
    return this.remote(async () => this.officeKit?.status() ?? { state: 'ready' }, signal);
  }
  async updateStatus(workspaceFileScope, refresh, signal) {
    return this.remote(async peer => ({ ...await this.updater.status({ refresh: refresh === true }), canUpdate: peer.principal === 'owner' }), signal);
  }
  /**
   * Owner only: installs the npm channel's newest release into the profile through `dsh plugin add`; HMR or a restart
   * loads it. Refused while any Kit view in any window has unsaved changes; on Windows Kit stops for it (updateWithKit).
   */
  async updatePlugin(workspaceFileScope, version, signal) {
    return this.remote(async peer => {
      if (peer.principal !== 'owner') throw new SuiteError('forbidden', 'Only the owner can update the plugin.');
      const offered = this.updater.checked?.latest;
      if (!offered || (version !== undefined && version !== offered)) throw new SuiteError('unavailable', 'The available version changed. Review the update details again.');
      // Older clients omit version. New clients pass the exact version whose notes they displayed.
      try { return await updateWithKit({ updater: this.updater, views: this.kitDocs?.views.values() ?? [], kit: this.kit, platform: this.platform, target: offered }); }
      catch (error) {
        if (error?.code === 'unsupported-operation' || error?.code === 'unsaved-changes') throw new SuiteError(error.code, error.message);
        throw new SuiteError('unavailable', `The plugin could not be updated: ${redactCredentials(error?.stderr || error?.message).trim().split('\n').slice(-2).join(' ')}`);
      }
    }, signal);
  }
  /**
   * 「本次更新」 for the owner: `{ current, previous?, entries: [{ version, highlights }] }` when this plugin version has
   * notes not shown yet on this machine, else null. Other principals never see it.
   */
  async whatsNew(workspaceFileScope, signal) {
    return this.remote(async peer => peer.principal === 'owner' ? await this.whatsNewNotes.pending() : null, signal);
  }
  /** Owner only: the card for `version` (this plugin's own) appeared, so it is not shown again. */
  async whatsNewSeen(workspaceFileScope, version, signal) {
    return this.remote(async peer => {
      if (peer.principal !== 'owner') throw new SuiteError('forbidden', 'Only the owner can dismiss the update notes.');
      try { await this.whatsNewNotes.markSeen(version); }
      catch (error) { throw new SuiteError('unavailable', `The update notes could not be marked as seen (${error?.code ?? error?.message ?? 'error'}).`); }
      return { seen: true };
    }, signal);
  }
  /**
   * A skill's resource directory for a tool call: downloaded first when needed and pinned to the session (spec
   * 261110-universal-package §4.5). Without the universal package's skills (the repository, tests): the fixed roots.
   */
  async skillRoot(name, sessionId, sub, { signal } = {}) {
    const fixed = { 'shimo-pptx': this.styleRoot, 'shimo-docx': this.docTemplateRoot, 'shimo-xlsx': this.sheetTemplateRoot }[name];
    if (!this.skills) return fixed;
    if (!this.skills.names.includes(name)) throw new SuiteError('unsupported-operation', `此版本插件未包含 ${name}，没有它的样式包或模板。This build of the plugin does not include ${name}; its style packs and templates are not available.`);
    try { return join(await this.skills.resolve(name, sessionId, { signal }), sub); }
    catch (error) { throw new SuiteError('unavailable', `${name} 的样式包/模板还没下载好（${error?.message ?? error}），请稍后重试。The ${name} resources are not downloaded yet; try again shortly.`); }
  }
  /** Style packs and templates the input box shows: fetched first when the package ships the light layer only. */
  async skillsForUi(names, signal) {
    if (!this.skills) return;
    // Not pinned (only a session's tool calls pin): the picker follows the effective version as it moves.
    await Promise.all(names.map(name => (this.skills.effective(name).ready ? undefined : this.skills.ensure(name, { version: this.skills.effective(name).selected, budgetMs: 120_000, signal }))?.catch(error => this.ctx.logger?.('officesdk/skills')?.warn?.(`${name}: ${error?.message}`))));
    this.styleCache = undefined;
  }
  /** Style packs there are now; cached briefly because the model's runtime context asks on every step. */
  styleIds({ refresh = false } = {}) {
    const now = Date.now();
    if (refresh || !this.styleCache || now - this.styleCache.at > STYLE_CACHE_MS) this.styleCache = { at: now, ids: styleNames(this.styleRoot) };
    return this.styleCache.ids;
  }
  /**
   * The style a presentation in `path` (workspace root or a folder in it) gets when the model names none: the one the
   * user remembered for that workspace while its pack still exists, else default.
   */
  defaultStyleFor(path, options) {
    const remembered = this.styleMemory.rememberedFor(path);
    return remembered && this.styleIds(options).includes(remembered) ? { style: remembered, source: 'workspace' } : { style: DEFAULT_STYLE, source: 'default' };
  }
  /**
   * Runtime context for the model: the workspace's remembered style, only when there is one besides default, and the
   * recommendation the user is answering in this Session (style-suggest.mjs); else ''.
   */
  styleContext(cwd, sessionId) {
    const answered = sessionId ? suggestionContext(this.styleSuggestions.forStep(sessionId), id => this.styleName(id)) : '';
    const workspace = this.workspaceStyleContext(cwd);
    // The inspiration library's pick comes last: it is the user naming a template, so it outranks the default line.
    const picked = sessionId ? pickContext(this.templatePicks?.forStep(sessionId), pick => this.templateName(pick)) : '';
    return [workspace, answered, picked].filter(Boolean).join('\n');
  }
  /** A picked template's display name: a pack's for PPT, the built-in table's for Word and Excel. */
  templateName(pick) {
    if (pick.kind === 'ppt') return this.styleName(pick.id);
    const found = BUILTIN_TEMPLATES[pick.kind]?.find(item => item.id === pick.id);
    return found ? found.name.zh + (found.name.en !== found.name.zh ? ` / ${found.name.en}` : '') : '';
  }
  /** A pack's Chinese display name (the English one too when it differs), or '' when it cannot be read. */
  styleName(id) {
    try { const card = styleCard(this.styleRoot, id, { thumb: false }); return card.name.zh + (card.name.en !== card.name.zh ? ` / ${card.name.en}` : ''); }
    catch { return ''; }
  }
  workspaceStyleContext(cwd) {
    const { style, source } = this.defaultStyleFor(cwd);
    if (source !== 'workspace' || style === DEFAULT_STYLE) return '';
    // Only the name: no thumbnail is read on this per-step path.
    let card;
    try { card = styleCard(this.styleRoot, style, { thumb: false }); } catch {}
    const title = card ? ` (${card.name.zh}${card.name.en !== card.name.zh ? ` / ${card.name.en}` : ''})` : '';
    return `Workspace default PPT style (shimo-pptx): ${style}${title}. When the user names no style, use it: leave style out of presentation_create, and pass --style ${style} to brand_deck.py.`;
  }
  /** The input box's style picker: every pack's card, and this workspace's default. */
  async listStyles(workspaceFileScope, signal) {
    return this.remote(async peer => {
      await this.suite.authorize(peer, workspaceFileScope, signal ?? peer.signal, 'open', '', 0);
      await this.skillsForUi(['shimo-pptx'], signal);
      const ids = this.styleIds({ refresh: true });
      const styles = styleCards(this.styleRoot, ids, text => this.ctx.logger?.('officesdk/styles')?.warn?.(text));
      const templates = this.templateIds();
      return { styles, ...templates && { templates }, ...this.styleChoice(workspaceFileScope), canRemember: peer.principal === 'owner' };
    }, signal);
  }
  /**
   * `rememberedFrom`: the enclosing folder whose choice this workspace inherits (a workspace inside another one), or
   * null when the choice is this workspace's own; only its own choice can be cleared here.
   */
  styleChoice(scope) {
    const found = this.styleMemory.lookup(scope.workspaceRoot);
    return { defaultStyle: this.defaultStyleFor(scope.workspaceRoot).style, remembered: found?.style ?? null, rememberedFrom: found && !found.own ? found.root : null };
  }
  /** Owner only: remember `style` as this workspace's default, or forget it with null. */
  async rememberStyle(workspaceFileScope, style, signal) {
    return this.remote(async peer => {
      if (peer.principal !== 'owner') throw new SuiteError('forbidden', 'Only the owner can change the workspace default style.');
      await this.suite.authorize(peer, workspaceFileScope, signal ?? peer.signal, 'open', '', 0);
      if (style !== null) {
        const ids = this.styleIds({ refresh: true });
        const resolved = resolveStyle(this.styleRoot, style, ids);
        if (!resolved) throw new SuiteError('invalid-edit', `Unknown style ${JSON.stringify(style)}; available: ${styleList(this.styleRoot, ids)}.`);
        style = resolved;
      }
      try { await this.styleMemory.set(workspaceFileScope.workspaceRoot, style); }
      catch (error) { throw new SuiteError('unavailable', `The workspace default style could not be saved (${error?.code ?? error?.name ?? 'error'}).`); }
      return this.styleChoice(workspaceFileScope);
    }, signal);
  }
  /**
   * The inspiration library's pick for this Session (`id` null clears it): `kind` is word, excel or ppt; a PPT id is a
   * pack (a display name or alias is resolved), a Word or Excel id a built-in template. Any peer that may open the
   * Session may set it: it only shapes that Session's next request. Returns the pick now recorded, or null.
   */
  async selectTemplate(workspaceFileScope, kind, id, signal) {
    return this.remote(async peer => {
      await this.suite.authorize(peer, workspaceFileScope, signal ?? peer.signal, 'open', '', 0);
      if (id === null || id === undefined) return this.templatePicks.set(workspaceFileScope.sessionId, null);
      if (!PICK_KINDS.includes(kind)) throw new SuiteError('invalid-edit', `Unknown template kind ${JSON.stringify(kind)}; use ${PICK_KINDS.join(', ')}.`);
      let resolved;
      if (kind === 'ppt') {
        const ids = this.styleIds({ refresh: true });
        resolved = resolveStyle(this.styleRoot, id, ids);
        if (!resolved) throw new SuiteError('invalid-edit', `Unknown style ${JSON.stringify(id)}; available: ${styleList(this.styleRoot, ids)}.`);
      } else {
        // Only what this copy can create from (the public build: Word gongwen, no Excel template).
        const allowed = this.templateIds()?.[kind];
        const available = BUILTIN_TEMPLATES[kind].filter(item => !allowed || allowed.includes(item.id));
        resolved = available.find(item => item.id === id)?.id;
        if (!resolved) throw new SuiteError('invalid-edit', `Unknown ${kind} template ${JSON.stringify(id)}; available: ${available.map(item => item.id).join(', ') || 'none'}.`);
      }
      return this.templatePicks.set(workspaceFileScope.sessionId, { kind, id: resolved });
    }, signal);
  }
  /** The Session's current pick (the Client restores its tag from it on load), or null. */
  async templatePick(workspaceFileScope, signal) {
    return this.remote(async peer => {
      await this.suite.authorize(peer, workspaceFileScope, signal ?? peer.signal, 'open', '', 0);
      return this.templatePicks.get(workspaceFileScope.sessionId);
    }, signal);
  }
  /** The recommendation strip: this Session's open recommendation as cards with reasons (marking it shown), or null. */
  async styleSuggestion(workspaceFileScope, signal) {
    return this.remote(async peer => {
      await this.suite.authorize(peer, workspaceFileScope, signal ?? peer.signal, 'open', '', 0);
      const suggestion = this.styleSuggestions.show(workspaceFileScope.sessionId);
      if (!suggestion) return null;
      await this.skillsForUi(['shimo-pptx'], signal);
      const ids = this.styleIds({ refresh: true });
      const warn = text => this.ctx.logger?.('officesdk/styles')?.warn?.(text);
      const cards = styleCards(this.styleRoot, suggestion.candidates.map(item => item.style).filter(id => ids.includes(id)), warn);
      const reasons = new Map(suggestion.candidates.map(item => [item.style, item.reason]));
      const fallback = this.defaultStyleFor(workspaceFileScope.workspaceRoot).style;
      let defaultCard = null;
      try { defaultCard = styleCard(this.styleRoot, fallback, { thumb: false }); } catch {}
      return { id: suggestion.id, topic: suggestion.topic, candidates: cards.map(card => ({ ...card, reason: reasons.get(card.id) ?? '' })), defaultStyle: defaultCard ?? { id: fallback, name: { zh: fallback, en: fallback } } };
    }, signal);
  }
  /** The user picked, closed the strip, or a new turn began: end that recommendation (only if it is still current). */
  async dismissStyleSuggestion(workspaceFileScope, id, signal) {
    return this.remote(async peer => {
      await this.suite.authorize(peer, workspaceFileScope, signal ?? peer.signal, 'open', '', 0);
      return { dismissed: this.styleSuggestions.dismiss(workspaceFileScope.sessionId, typeof id === 'string' ? id : undefined) };
    }, signal);
  }
  /** AI tool: record the model's recommended packs for the strip above the input box. */
  async aiSuggestStyles(sessionId, { candidates, topic } = {}) {
    const scope = await this.scopeOf(sessionId);
    if (this.skills) await this.skillRoot('shimo-pptx', sessionId, 'styles');
    const ids = this.styleIds({ refresh: true });
    let proposed;
    const named = Array.isArray(candidates) ? candidates.map(item => (item && typeof item.style === 'string' ? { ...item, style: resolveStyle(this.styleRoot, item.style, ids) ?? item.style } : item)) : candidates;
    try { proposed = this.styleSuggestions.propose(sessionId, { candidates: named, topic }, ids); }
    catch (error) { throw new SuiteError('invalid-edit', error.message); }
    const { suggestion, dropped } = proposed;
    const fallback = this.defaultStyleFor(scope.workspaceRoot).style;
    return {
      suggestionId: suggestion.id, shown: true,
      candidates: suggestion.candidates.map((item, index) => ({ index: index + 1, style: item.style, name: this.styleName(item.style), reason: item.reason })),
      ...(dropped.length ? { dropped } : {}),
      defaultStyle: fallback,
      next: 'The user now sees these as cards above the input box, with "use the default" and "see all styles". End your turn here with one short sentence asking them to pick (do not list the styles again). Their choice arrives as their next message.',
    };
  }
  /**
   * The user selected text (or cells) in this tab's document and wrote what to change: record the selection for
   * selection_rewrite and return the message that carries the request into the conversation. The Session is the tab's
   * (its workspaceFileScope), never one the page names.
   */
  async selectionCreate(workspaceFileScope, path, selectionRequest, selectionLanguage, signal) {
    return this.remote(async peer => {
      await this.suite.authorize(peer, workspaceFileScope, signal ?? peer.signal, 'edit', '', 0);
      if (typeof path !== 'string' || !path) throw new SuiteError('invalid-path', 'path is required.');
      const request = validateSelectionRequest(selectionRequest);
      if (/\.([A-Za-z0-9]+)$/.exec(path)?.[1]?.toLowerCase() !== request.kind) throw new SuiteError('invalid-edit', 'The selection does not match the document type.');
      const record = this.selections.create(workspaceFileScope.sessionId, { path, realPath: await this.realPathOf(workspaceFileScope, path), name: basename(path), request, lang: selectionLanguage === 'en' ? 'en' : 'zh' });
      record.tracked = selectionRequest.trackRange===true;
      return { id: record.id, message: selectionMessage(record) };
    }, signal);
  }
  async selectionParallel(workspaceFileScope, id, signal) {
    return this.remote(async peer => {
      const sessionId = workspaceFileScope.sessionId;
      const stop = AbortSignal.any([signal, peer.signal].filter(Boolean));
      await this.suite.authorize(peer, workspaceFileScope, stop, 'edit', '', 0);
      const record = this.selections.get(sessionId, id);
      if (record.state !== 'pending' || record.generating) throw new SuiteError('busy', '这处修改已经提交。');
      const lifetime = new AbortController();
      record.generating = lifetime;
      const workerSignal = AbortSignal.any([stop, lifetime.signal]);
      try {
        const agent = await this.ctx.typert.lookups.get('agent')?.resolve(sessionId);
        if (!agent?.session?.append) throw new SuiteError('unavailable', '无法连接当前对话，请重新打开会话后再发送。');
        // One editor command, one durable transcript row. Log-only command events leave the
        // main turn's model history and inbox intact while the selection worker runs.
        const commandId = 'cmd-officesdk-' + id;
        agent.session.append('command/run', { commandId, name: PARALLEL_EDIT_COMMAND, args: parallelEditMessage(record), source: { kind: 'user' } });
        try {
          const replacement = await generateSelectionReplacement(agent, record, workerSignal);
          if (!replacement) {
            this.selections.report(sessionId, id, 'cancelled');
            agent.session.append('command/done', { commandId, kind:'success' });
            return { recorded:true, skipped:true };
          }
          workerSignal.throwIfAborted();
          await this.aiRewriteSelection(sessionId, {selection:id, ...replacement}, workerSignal);
        } catch (error) {
          const text = error.message || String(error);
          agent.session.append('command/done', { commandId, kind:'error', text });
          return { recorded:true, error:text };
        }
        agent.session.append('command/done', { commandId, kind:'success' });
        return { recorded:true };
      } finally { record.generating = false; }
    }, signal);
  }
  /** The user cancelled a waiting selection, or kept / undid a rewritten one, in the tab. */
  async selectionUpdate(workspaceFileScope, id, selectionState, signal) {
    return this.remote(async peer => {
      await this.suite.authorize(peer, workspaceFileScope, signal ?? peer.signal, 'edit', '', 0);
      return this.selections.report(workspaceFileScope.sessionId, id, selectionState);
    }, signal);
  }
  /**
   * AI tool selection_rewrite: replace the user's selection with `text` (Word) or `values` (spreadsheet), in the editor,
   * without an approval card (the user asked for it there), once per selection.
   */
  async aiRewriteSelection(sessionId, { selection, text, values } = {}, signal) {
    await this.scopeOf(sessionId);
    const record = this.selections.take(sessionId, selection);
    const { kind, anchor, original, cells } = record.request;
    let payload;
    try {
      if (kind !== 'xlsx') {
        if (values !== undefined) throw new SuiteError('invalid-edit', 'This selection is a text range: give text, not values.');
        payload = { text: checkRewriteText(text) };
      } else {
        if (values === undefined) throw new SuiteError('invalid-edit', `This selection is ${anchor.rowCount}×${anchor.columnCount} spreadsheet cells: give values (rows of cell values), not text.`);
        payload = { values: checkRewriteValues(values, anchor) };
      }
    } catch (error) { record.state = 'pending'; throw error; }
    try {
      const result = await this.whileAiWrites(sessionId, record.path, () => this.agents.request(sessionId, { kind: 'rewrite', path: record.path, selection: { id: record.id, kind, anchor, original, ...(cells && { cells }) }, ...payload }, signal));
      record.state = 'used';
      return { ...result, selection: record.id, path: record.path, unsaved: !result.pendingReview };
    } catch (error) {
      if (error?.code === 'source-changed') {
        record.state = 'stale';
        throw new SuiteError('selection-stale', `The selected text changed since the user selected it, so nothing was replaced. Ask them to select it again and send the request from the editor. (${error.message})`);
      }
      // Stopped while writing: what was written stays (no rollback), so the selection cannot be applied again.
      record.state = error?.code === 'cancelled' ? 'used' : 'pending';
      throw error;
    }
  }
  async listDrafts(workspaceFileScope, signal) { return this.remote(peer => this.suite.listDrafts(workspaceFileScope, signal, peer), signal); }
  async discardDraft(workspaceFileScope, draftId, signal) { return this.remote(peer => this.suite.discardDraft(workspaceFileScope, draftId, signal, peer), signal); }
  async exportDraft(workspaceFileScope, draftId, target, signal) { return this.remote(peer => this.suite.exportDraft(workspaceFileScope, draftId, { target }, signal, peer), signal); }

  /** Client long poll for AI document commands addressed to this Session. */
  async agentNext(workspaceFileScope, confirmDelivery, signal) {
    return this.remote(async peer => {
      await this.suite.authorize(peer, workspaceFileScope, signal ?? peer.signal, 'edit', '', 0);
      return this.agents.next(workspaceFileScope.sessionId, AbortSignal.any([signal, peer.signal].filter(Boolean)), confirmDelivery === true);
    }, signal);
  }
  async agentAlive(workspaceFileScope, requestId, deliveryToken, signal) {
    return this.remote(async peer => {
      await this.suite.authorize(peer, workspaceFileScope, signal ?? peer.signal, 'edit', '', 0);
      return { alive: this.agents.alive(workspaceFileScope.sessionId, requestId, deliveryToken) };
    }, signal);
  }
  async agentReply(workspaceFileScope, requestId, ok, result, code, message, signal) {
    return this.remote(async peer => {
      await this.suite.authorize(peer, workspaceFileScope, signal ?? peer.signal, 'edit', '', 0);
      return { accepted: this.agents.reply(workspaceFileScope.sessionId, requestId, { ok: ok === true, result, code, message }) };
    }, signal);
  }

  /** The tab's Stop: the model's generation of this file asks for each batch again from now on. */
  async agentStop(workspaceFileScope, path, signal) {
    return this.remote(async peer => {
      await this.suite.authorize(peer, workspaceFileScope, signal ?? peer.signal, 'edit', '', 0);
      const generation = await this.generationOf(workspaceFileScope.sessionId, workspaceFileScope, path);
      if (generation) generation.stopped = true;
      return { stopped: Boolean(generation) };
    }, signal);
  }

  /** The real path a Session path names, for matching generation windows; undefined if it cannot be resolved. */
  async realPathOf(scope, path) {
    try { return await realpath(resolvePath(scope.workspaceRoot, path)); } catch { return undefined; }
  }
  /**
   * The window for the file `path` names, if this Session created that very file: the same real path and the same
   * file identity (device and inode) as right after presentation_create wrote it. A file swapped in later does not inherit it.
   * Kit saves in place and keeps the inode; a save that renames a new file over it (Core's writer.replace) ends the
   * window, so later batches ask again: it fails closed.
   */
  async generationOf(sessionId, scope, path) {
    const real = await this.realPathOf(scope, path);
    const key = real && `${sessionId}\0${real}`;
    const generation = key && this.generations.get(key);
    if (!generation) return undefined;
    const now = await lstat(real).catch(() => undefined);
    if (!now || now.dev !== generation.dev || now.ino !== generation.ino) { this.generations.delete(key); return undefined; }
    return generation;
  }

  /**
   * AI tool: a new presentation from a style pack's template, written where the Session may write (outside the
   * workspace only with `approve`), then opened beside the conversation. The file is the model's to generate:
   * its edits apply without a card per batch until the user stops it or the budget runs out.
   */
  async aiCreate(sessionId, { path, style, slideCount, slideLayouts, reportingData }, { approve, signal } = {}) {
    const scope = await this.scopeOf(sessionId);
    if (typeof path !== 'string' || !/\.pptx$/i.test(path)) throw new SuiteError('invalid-path', 'presentation_create makes .pptx files; give a path ending in .pptx.');
    if (slideCount !== undefined && (!Number.isInteger(slideCount) || slideCount < 1 || slideCount > PPTX_LIMITS.maxPlannedSlides)) {
      throw new SuiteError('invalid-edit', `slideCount must be between 1 and ${PPTX_LIMITS.maxPlannedSlides}.`);
    }
    if (slideLayouts !== undefined && (!Array.isArray(slideLayouts) || slideLayouts.length !== slideCount || slideLayouts.some(role => !['cover','content','title','closing','blank'].includes(role)))) {
      throw new SuiteError('invalid-edit', 'Provide one valid slide layout role for each planned slide.');
    }
    // No style named: the workspace's remembered one (the input box's PPT style picker), else default.
    // A named style may be a display name or alias (the inspiration library writes names); unknown ones fail in loadStyle.
    let styleRoot = await this.skillRoot('shimo-pptx', sessionId, 'styles', { signal });
    // A style the session's pinned version lacks but the newer effective one has (a skill update mid-session, §4.5 J1).
    if (this.skills && style !== undefined && style !== null && !resolveStyle(styleRoot, style)) {
      const current = this.skills.repin('shimo-pptx', sessionId);
      if (current && resolveStyle(join(current, 'styles'), style)) { styleRoot = join(current, 'styles'); this.ctx.logger?.('officesdk/skills')?.info?.(`shimo-pptx: moved session to ${current} for style ${style}`); }
    }
    const chosen = style === undefined || style === null ? this.defaultStyleFor(scope.workspaceRoot, { refresh: true }) : { style: resolveStyle(styleRoot, style) ?? style, source: 'given' };
    const { template, summary } = loadStyle(styleRoot, chosen.style);
    const fixed=summary.fixedDeck;
    if (fixed && slideCount !== undefined && slideCount !== fixed.slideCount) throw new SuiteError('invalid-edit', `This original template has ${fixed.slideCount} pages; use that page count.`);
    if (reportingData !== undefined && (chosen.style !== 'slide-002' || !reportingData || typeof reportingData !== 'object' || Array.isArray(reportingData) || Object.keys(reportingData).some(key => !['chart', 'bars', 'statuses'].includes(key)))) {
      throw new SuiteError('invalid-edit', 'reportingData accepts chart, bars and statuses for slide-002 only.');
    }
    const data = fixed ? await blankExistingDeck({ stylesRoot: styleRoot, style: chosen.style, dataRoot: this.dataRoot, signal, reportingData }) : template;
    if (fixed && fixed.textSlots === 'text-slots.json') summary.textSlots=readJson(join(styleRoot,chosen.style,'text-slots.json'));
    let created;
    try { created = await this.writer.create(scope, path, data, signal ?? AbortSignal.timeout(30_000)); }
    catch (error) {
      if (error?.code !== 'sandbox-denied' || !error.details?.outside) throw error;
      const outcome = await approve?.(`Create ${path} (outside the workspace) for a new presentation`);
      if (outcome !== 'allowed-once') throw new SuiteError('sandbox-denied', `The user did not allow creating ${path} outside the workspace (${outcome ?? 'no approval channel'}). Create it inside the workspace instead.`, { outside: true });
      created = await this.writer.create(scope, path, data, signal ?? AbortSignal.timeout(30_000), { approved: true });
    }
    // The writer resolved the target before writing; the window is that path, as long as it stays this regular file.
    const real = created.absolutePath;
    const identity = await lstat(real).catch(() => undefined);
    const windowed = Boolean(identity?.isFile() && await this.realPathOf(scope, real) === real);
    if (windowed) this.generations.set(`${sessionId}\0${real}`, { ...(fixed && { fixedSlides: { count: fixed.slideCount, done: new Set(), partial: new Set() } }), steps: slideCount ?? 0, imageBytes: 0, reserved: { steps: 0, imageBytes: 0 }, stopped: false, dev: identity.dev, ino: identity.ino });
    // A style is settled once a deck exists: the recommendation strip has nothing left to ask.
    this.styleSuggestions.dismiss(sessionId);
    const read = await this.agents.request(sessionId, { kind: 'read', path, ...(fixed ? {initialBlank:'ppt-template'} : slideCount ? {slideCount, slideLayouts} : {}) }, signal);
    return { created: true, style: summary, styleSource: chosen.source, ...read, generating: windowed };
  }

  /** AI tool: a new Word document from a shimo-docx template, generated live like presentation_create (host/document.mjs). */
  async aiCreateDocument(sessionId, args, options) { return createDocument(this, sessionId, args, options); }
  async aiFinishProjectTable(sessionId, args, signal) { return finishProjectTable(this, sessionId, args, signal); }
  async aiFinishMeetingTable(sessionId, args, signal) { return finishProjectTable(this, sessionId, {path:args.path,milestones:args.decisions}, signal, 'meeting'); }
  async aiFinishProposalTable(sessionId, args, signal) { return finishProjectTable(this, sessionId, {path:args.path,milestones:args.rows}, signal, 'proposal'); }
  /** AI tool: a new workbook from a shimo-xlsx template, generated live the same way. */
  async aiCreateSpreadsheet(sessionId, args, options) { return createSpreadsheet(this, sessionId, args, options); }

  /**
   * The model names images by path; the Host reads each one the Session can read and passes its bytes on. Sizes are
   * checked for the whole batch before any file is read, and files are read one at a time. A whole-file read is
   * capped by WorkspaceFiles (`maxFileBytes`); a ranged read cannot be used, its window is at most `maxBytes` (2 MB).
   */
  async loadImages(scope, ops, signal) {
    const list = Array.isArray(ops) ? ops : [];
    if (list.length > PPTX_LIMITS.maxOps) throw new SuiteError('invalid-edit', `At most ${PPTX_LIMITS.maxOps} operations per edit.`);
    const max = PPTX_LIMITS.maxImageBytes, mb = bytes => bytes / 1024 / 1024;
    const images = [];
    let total = 0;
    for (const [index, op] of list.entries()) {
      if (!op || typeof op !== 'object' || op.op !== 'addImage') continue;
      if (typeof op.path !== 'string' || !op.path) throw new SuiteError('invalid-edit', `Operation ${index + 1}: addImage needs the path of a PNG or JPEG file.`, { index });
      let info;
      try { info = await this.suite.source.stat(scope, op.path, signal); }
      catch (error) { throw new SuiteError('invalid-edit', `Operation ${index + 1}: cannot read ${op.path}: ${error?.message ?? error}`, { index }); }
      const size = Number(info?.size ?? info?.bytes ?? 0);
      if (size > max) throw new SuiteError('invalid-edit', `Operation ${index + 1}: the image is larger than ${mb(max)} MB.`, { index });
      total += size;
      if (total > PPTX_LIMITS.maxBatchImageBytes) throw new SuiteError('invalid-edit', `Operation ${index + 1}: images in one edit add up to more than ${mb(PPTX_LIMITS.maxBatchImageBytes)} MB; split the edit.`, { index });
      images.push(index);
    }
    const loaded = [...list];
    let imageBytes = 0;
    for (const index of images) {
      const { imageBase64: _bytes, imageType: _type, ...rest } = list[index];
      let bytes;
      try {
        const read = await this.suite.source.readBytes(scope, rest.path, {}, signal);
        bytes = new Uint8Array(read?.data ?? read);
      } catch (error) { throw new SuiteError('invalid-edit', `Operation ${index + 1}: cannot read ${rest.path}: ${error?.message ?? error}`, { index }); }
      // A file that grew after the size check is still capped.
      if (bytes.byteLength > max) throw new SuiteError('invalid-edit', `Operation ${index + 1}: the image is larger than ${mb(max)} MB.`, { index });
      const imageType = imageTypeOf(bytes);
      if (!imageType) throw new SuiteError('invalid-edit', `Operation ${index + 1}: ${rest.path} is not a PNG or JPEG image.`, { index });
      imageBytes += bytes.byteLength;
      if (imageBytes > PPTX_LIMITS.maxBatchImageBytes) throw new SuiteError('invalid-edit', `Operation ${index + 1}: images in one edit add up to more than ${mb(PPTX_LIMITS.maxBatchImageBytes)} MB; split the edit.`, { index });
      loaded[index] = { ...rest, imageType, imageBase64: Buffer.from(bytes).toString('base64') };
    }
    return { ops: loaded, imageBytes };
  }

  /** The Session's file scope, as WorkspaceFiles and the writer expect it. */
  async scopeOf(sessionId) {
    const scope = await this.ctx.typert.lookups.get('workspaceFileScope')?.resolve(sessionId);
    if (!scope || scope.sessionId !== sessionId) throw new SuiteError('forbidden', 'The Session is no longer available.');
    return scope;
  }

  /** AI tool: open the document beside the conversation and describe it. `callId` lets the Client label the chat row. */
  async aiRead(sessionId, path, signal, { callId, editTarget } = {}) {
    await this.scopeOf(sessionId);
    return this.agents.request(sessionId, { kind: 'read', path, ...(callId && { callId }), ...(editTarget && { editTarget }) }, signal);
  }

  /** The editing Kit view of `path` in this Session, if the document is open through Kit. */
  kitViewFor(sessionId, path) {
    for (const view of this.kitDocs?.views.values() ?? []) {
      if (view.closed || view.stale || !view.writer || view.scope.sessionId !== sessionId) continue;
      if (resolvePath(view.scope.workspaceRoot, path) === resolvePath(view.scope.workspaceRoot, view.path) || resolvePath(view.scope.workspaceRoot, path) === view.realPath) return view;
    }
    return undefined;
  }

  /** While the AI writes into a Kit view, saving and exporting it wait (spec §12.5). */
  async whileAiWrites(sessionId, path, run) {
    const view = this.kitViewFor(sessionId, path);
    const done = view ? this.kitDocs.aiBegin(view.id) : () => {};
    try { return await run(); } finally { done(); }
  }

  /** AI tool: show the change for confirmation, apply it in the editor. The file is untouched until a save. */
  async aiEdit(sessionId, { path, ops, summary, callId, readToken, completedSlideIds }, signal) {
    const scope = await this.scopeOf(sessionId);
    const editId = randomUUID();
    // Too many new slides fails before images are read or the editor is asked; the editor checks again.
    if (countNewSlides(ops) > PPTX_LIMITS.maxNewSlides) throw new SuiteError('invalid-edit', NEW_SLIDES_MESSAGE);
    const loaded = await this.loadImages(scope, ops, signal);
    // A file the model created and is still generating applies without a card, within its budget.
    const generation = await this.generationOf(sessionId, scope, path);
    if (generation?.tableFrozen) throw new SuiteError('busy', 'The project table is being committed; wait until it completes.');
    const steps = Array.isArray(ops) ? ops.length : 0;
    // The budget counts batches still running too, so parallel edits cannot overshoot it.
    const used = generation && { steps: generation.steps + generation.reserved.steps, imageBytes: generation.imageBytes + generation.reserved.imageBytes };
    const auto = Boolean(generation && !generation.stopped && used.steps + steps <= GENERATION_LIMITS.maxSteps && used.imageBytes + loaded.imageBytes <= GENERATION_LIMITS.maxImageBytes);
    if (auto) { generation.reserved.steps += steps; generation.reserved.imageBytes += loaded.imageBytes; }
    let result;
    try { result = await this.whileAiWrites(sessionId, path, () => this.agents.request(sessionId, { kind: 'edit', path, ops: loaded.ops, summary, editId, ...(completedSlideIds !== undefined && {completedSlideIds}), ...(callId && { callId }), ...(auto && { auto: true }), ...(auto && generation.fixedSlides && { slideProgress: slideSnapshot(generation) }), ...(typeof readToken === 'string' && { readToken }) }, signal)); }
    catch (error) {
      // Stopped from the tab while this batch ran: the window is over (a PPT batch is undone; Word keeps the steps it finished).
      if (generation && error?.code === 'cancelled') generation.stopped = true;
      throw error;
    } finally {
      if (auto) { generation.reserved.steps -= steps; generation.reserved.imageBytes -= loaded.imageBytes; }
    }
    if (generation && auto && !result?.declined) { generation.steps += steps; generation.imageBytes += loaded.imageBytes; }
    if (auto && generation.fixedSlides && result?.applied > 0 && !result.declined && !result.pendingReview) {
      const pages = generation.fixedSlides, finished = new Set(completedSlideIds ?? result.slideIds);
      for (const id of result.slideIds ?? []) {
        pages.done.delete(id);
        pages.partial.delete(id);
        (finished.has(id) ? pages.done : pages.partial).add(id);
      }
      if (pages.done.size === pages.count) delete generation.fixedSlides;
    }
    if (result?.declined) return { declined: true };
    if (result?.pendingReview) return { ...result, editId, unsaved: false };
    // Selections still waiting on this file no longer hold what the user selected.
    const real = await this.realPathOf(scope, path);
    if (real) this.selections.staleFor(sessionId, real);
    return { ...result, editId, unsaved: true, ...(generation && { generating: auto && !generation.stopped }) };
  }

  /** AI tool: undo the most recent AI edit inside the editor. */
  async aiUndo(sessionId, { path, editId }, signal) {
    await this.scopeOf(sessionId);
    return this.whileAiWrites(sessionId, path, () => this.agents.request(sessionId, { kind: 'undo', path, editId }, signal));
  }

  /** AI save of a Kit view: the page hands over what it holds, then Kit writes the authorized real path. */
  async aiSaveKit(sessionId, scope, path, view, { approve, signal }) {
    const flushed = await this.agents.request(sessionId, { kind: 'flush', path }, signal);
    if (!view.unsaved) return { saved: false, unchanged: true, path: view.realPath };
    let approved = false;
    const access = await this.writer.accessFor(scope, view.realPath);
    if (access === 'denied') throw new SuiteError('sandbox-denied', `This Session cannot write ${view.realPath} (read only). The changes stay in the editor; switch the Session to a writable mode, then save.`);
    if (access === 'outside') {
      const outcome = await approve?.(`Write ${view.realPath} (outside the workspace) with the ${nounOf(path)} edits`);
      if (outcome !== 'allowed-once') throw new SuiteError('sandbox-denied', `The user did not allow writing ${view.realPath} outside the workspace (${outcome ?? 'no approval channel'}). The changes stay in the editor; they can save from the panel.`, { outside: true });
      approved = true;
    }
    let saved;
    try { saved = await this.kitDocs.save(view.id, view.peer, { approved, signal }); }
    catch (error) { throw fromKit(error); }
    this.agents.notify(sessionId, { kind: 'saved', path, version: saved.version, editRevision: flushed?.editRevision });
    return { saved: true, path: view.realPath };
  }

  /**
   * AI tool: write the editor's changes back to the file. A target outside the writable roots needs
   * `approve(reason)` to return 'allowed-once' (DSH approval, asked inside the tool's turn).
   */
  async aiSave(sessionId, path, { approve, signal } = {}) {
    const scope = await this.scopeOf(sessionId);
    const kitView = this.kitViewFor(sessionId, path);
    if (kitView) {
      const result = await this.aiSaveKit(sessionId, scope, path, kitView, { approve, signal });
      if ((result.saved || result.unchanged) && kitView.suite === 'presentation') {
        const generation = await this.generationOf(sessionId, scope, path);
        if (generation) delete generation.fixedSlides;
      }
      return result;
    }
    const session = await this.suite.find(scope, path, { principal: 'owner' }, signal);
    if (!session) throw new SuiteError('not-found', nounOf(path) === 'presentation' ? 'That presentation is not open, so there is nothing to save. Use presentation_read first.' : `That ${nounOf(path)} is not open, so there is nothing to save. Read it first.`);
    if (!this.suite.dirty(session)) return { saved: false, unchanged: true, version: session.baseVersion, path: session.absolutePath };
    let approved = false;
    const access = await this.suite.writeAccess(session);
    if (access === 'denied') throw new SuiteError('sandbox-denied', `This Session cannot write ${session.absolutePath} (read only). The changes stay in the editor; switch the Session to a writable mode, then save.`);
    if (access === 'outside') {
      const outcome = await approve?.(`Write ${session.absolutePath} (outside the workspace) with the ${nounOf(path)} edits`);
      if (outcome !== 'allowed-once') throw new SuiteError('sandbox-denied', `The user did not allow writing ${session.absolutePath} outside the workspace (${outcome ?? 'no approval channel'}). The changes stay in the editor; they can save from the panel.`, { outside: true });
      approved = true;
    }
    const written = await this.suite.saveSession(session, signal ?? AbortSignal.timeout(300_000), { approved, peer: session.peer, system: !session.peer?.id });
    this.agents.notify(sessionId, { kind: 'saved', path, version: written.version, editRevision: written.editRevision });
    return { saved: true, ...written };
  }
}

/** Same-origin writes only: the Origin header must name the host this request reached. */
export function sameOrigin(origin, req) {
  const host = req.headers.host;
  if (!host) return false;
  try {
    const url = new URL(origin);
    // Browsers never let a web page claim the desktop app scheme as its Origin.
    if (url.protocol === 'dsh-app:') return true;
    return url.host === host && (url.protocol === 'http:' || url.protocol === 'https:');
  } catch { return false; }
}
