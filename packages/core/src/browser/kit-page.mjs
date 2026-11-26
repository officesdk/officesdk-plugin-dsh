import { createFrameChannel, quietModals, serveAgent } from './frame.mjs';

export { quietModals };
import { installSheetMediaBridge, sheetMediaReference } from './kit-sheet-assets.mjs';
import { createSelectionFeature, createSelectionUi } from './selection-ui.mjs';
import { installEditorChrome } from './editor-chrome.mjs';
import { createPendingReviews } from './pending-reviews.mjs';

/**
 * The editor page for a Kit view (spec §12.4, §13.3): office-kit's editor package mounts the suite, every local
 * change is numbered and sent to the Host in order, and the Host forwards it to Kit. The page never sees
 * Kit, its work directory or file passwords. When the Host refuses a change (a gap, a fault, a restarted
 * Kit) the page rebuilds the editor from the Host's current content.
 *
 * URL: <prefix>/kit/static/page/index.html?view=<id>&revision=<n>&parentOrigin=<origin>[&lang=<en|zh>]
 */
/** What the read-only refusal calls the open file. */
const SUITE_NOUN = { presentation: 'presentation', document: 'document', sheet: 'spreadsheet' };
/** Input that means the user is working in the editor; a change before any of it (or an AI edit) is the editor's own. */
export const ACT_EVENTS = Object.freeze(['keydown', 'beforeinput', 'input', 'compositionstart', 'paste', 'cut', 'drop', 'pointerdown', 'mousedown', 'touchstart', 'contextmenu']);
const RUNTIME_SCRIPTS = ['shared/react.js', 'shared/react-dom.js', 'shared/react-aliases.js'];

/** What each suite's editor needs beyond the common options (office-kit editor package contracts). */
export function suiteMountOptions({ suite, mode, user, language, nextChangeSetRevision }) {
  const editing = mode === 'edit';
  // Sheet is only accepted in zh-CN and brings its own React. Writer names the user in both modes; Sheet takes
  // a user only for editing (its read-only mode refuses edit options).
  return {
    language: suite === 'sheet' ? 'zh-CN' : language,
    ...(suite !== 'sheet' && { runtimeScripts: RUNTIME_SCRIPTS }),
    ...(editing && suite !== 'sheet' && { nextChangeSetRevision }),
    ...((suite === 'document' || (suite === 'sheet' && editing)) && user && { user: { uid: user.uid, name: user.name } }),
    // Sheet stays on its current (new) UI: no switch to the old one, and no feedback entry (no feedbackUrl). office-kit's
    // editor page takes `ui.versionSwitch` since 0.3.0-rc.1 (office-kit request K18); 0.2.4 and earlier refuse `ui`.
    ...(suite === 'sheet' && { ui: { versionSwitch: { disabled: true } } }),
  };
}

/**
 * DOCX opens with the whole first page visible, as Core's preview does. The Writer SDK has no zoom API in this mode, so the page
 * runs the SDK's own status-bar command, found by its language-independent hooks. Missing → default zoom.
 */
export async function fitWholeWriterPage(document, { timeoutMs = 5000, wait = ms => new Promise(resolve => setTimeout(resolve, ms)) } = {}) {
  for (let waited = 0; waited <= timeoutMs; waited += 100) {
    const zoom = document.querySelector('.editor-statusbar__zoom-value');
    if (zoom) {
      zoom.click();
      for (let tries = 0; tries < 20; tries++) {
        const item = document.querySelector('[data-menu-binding-token="root:fit-whole-page"]');
        if (item) { item.click(); return true; }
        await wait(50);
      }
      zoom.click();
      break;
    }
    await wait(100);
  }
  console.warn('officesdk: whole page zoom unavailable; opening at the default zoom');
  return false;
}

/** The media kind each suite's Kit store expects for an inserted file (sheet takes none). */
export function attachmentKind(editorModule, suite, bucket, mime) {
  if (suite === 'presentation') return editorModule.presentationResourceCategory(bucket, mime);
  if (suite === 'document') return bucket === 'images' ? 'image' : 'embedding';
  return undefined;
}

export class PageError extends Error {
  constructor(code, message, { rebuild = false } = {}) { super(message); this.name = 'PageError'; this.code = code; this.rebuild = rebuild; }
}

async function call(fetchImpl, url, init) {
  const response = await fetchImpl(url, { credentials: 'same-origin', ...init });
  if (response.ok) return response;
  const body = await response.json().catch(() => ({}));
  throw new PageError(body.error ?? 'unavailable', body.message ?? `Request failed: ${response.status}`, { rebuild: body.rebuild === true || response.status === 404 });
}

/**
 * Sends numbered changes one at a time. A refused change stops the queue and reports `onRebuild`; the
 * page then starts a new sequence from 0 after the Host's content reload.
 */
export function createChangeQueue({ send, onAccepted = () => {}, onRebuild = () => {} }) {
  let seq = 0, chain = Promise.resolve(), stopped = false, generation = 0;
  return {
    push(change, tag) {
      if (stopped) return;
      const mine = generation, number = seq++;
      chain = chain.then(async () => {
        if (stopped || mine !== generation) return;
        try { await send(number, change, tag); onAccepted(number, tag); }
        catch (error) { if (!stopped) { stopped = true; onRebuild(error); } }
      });
    },
    /** Resolves when every change pushed so far has its Host result. */
    idle: () => chain,
    get sent() { return seq; },
    get stopped() { return stopped; },
    restart() { generation++; seq = 0; stopped = false; chain = Promise.resolve(); },
  };
}

/**
 * The page's interface language: the `lang` the tab opened the frame with (the DSH interface language), else
 * `<html lang>`, else Chinese. Anything not English is Chinese, as the page's labels only come in those two.
 */
export function pageLocale(params, document) {
  const lang = params?.get?.('lang') || document?.documentElement?.lang || 'zh';
  return lang.toLowerCase().startsWith('en') ? 'en' : 'zh';
}

export async function startKitPage({ window, document, fetch: fetchImpl = window.fetch.bind(window), loadEditor, createAgent }) {
  quietModals(window);
  const params = new URL(window.location.href).searchParams;
  const viewId = params.get('view');
  const revision = Number(params.get('revision') ?? 0);
  const locale = pageLocale(params, document);
  if (!viewId) throw new Error('The page needs a view');
  const base = new URL(`../../v/${encodeURIComponent(viewId)}/`, window.location.href);
  const staticRoot = new URL('../', window.location.href);
  const lifetime = new AbortController();
  const channel = createFrameChannel({ window, id: viewId, revision, parentOrigin: params.get('parentOrigin'), signal: lifetime.signal });
  try {
    const editorModule = await loadEditor(new URL('editor/bundle/offline-editor.js', staticRoot).href);
    const meta = await (await call(fetchImpl, new URL('meta', base))).json();
    const reviewUrl = new URL('reviews', base);
    const reviews = meta.reviewHistory && createPendingReviews(
      await (await call(fetchImpl, reviewUrl)).json(),
      async patch => (await call(fetchImpl, reviewUrl, {method:'POST', headers:{'content-type':'application/json'}, body:JSON.stringify(patch)})).json(),
      reviewSaveFailed => channel.notify('document-ui', {reviewSaveFailed}),
    ) || undefined;
    if (meta.suite === 'sheet') {
      // The Sheet SDK reads pictures by their raw reference and needs the icon faces the resources stripped.
      installSheetMediaBridge(window, async (reference, purpose) => {
        const response = await fetchImpl(new URL(`media/${encodeURIComponent(reference)}?purpose=${purpose}`, base), { credentials: 'same-origin' });
        return { ok: response.ok, status: response.status, contentType: response.headers.get('content-type') ?? undefined, bytes: response.ok ? await response.arrayBuffer() : undefined };
      });
      const fonts = document.createElement('link');
      fonts.rel = 'stylesheet';
      fonts.href = new URL('fonts/sheet-host-fonts.css', staticRoot).href;
      document.head.append(fonts);
    }
    const resourcesBase = new URL('resources/', staticRoot).href;
    const catalog = editorModule.parseResourceCatalog(await (await call(fetchImpl, new URL('catalog.json', resourcesBase))).json());
    const root = document.getElementById('editor-root');
    const generationMode = params.get('generation');
    const initiallyBlank = ['project','meeting','proposal','template','letter'].includes(generationMode) && meta.suite === 'document' && meta.mode === 'edit';
    const initiallyBlankDeck = generationMode === 'ppt-template' && meta.suite === 'presentation' && meta.mode === 'edit';
    if (initiallyBlank || initiallyBlankDeck) root.style.visibility = 'hidden';
    // The editor only reads resources under `assets.baseUrl`, so each reference becomes one path segment there.
    // The editor's asset proxy resolves every URL the SDK requests, including one it already resolved and
    // extended (Sheet adds `?sm_xform=` for thumbnails): a URL under this view's media/ stays as it is.
    const mediaBase = new URL('media/', base);
    const mediaUrl = reference => {
      if (typeof reference === 'string' && reference.startsWith(mediaBase.href)) return reference;
      const kitReference = meta.suite === 'sheet' ? sheetMediaReference(reference) ?? reference : reference;
      return new URL(encodeURIComponent(kitReference), mediaBase).href;
    };

    const fontProvider = {
      async getSystemCatalog(options) {
        try { return editorModule.parseSystemFontCatalog(await (await call(fetchImpl, new URL('../../fonts/catalog', base), { signal: options?.signal })).json()); }
        catch (error) { options?.signal?.throwIfAborted(); return { status: 'unavailable', reason: 'query-failed' }; }
      },
      async readResource(resourceId, options) {
        const response = await call(fetchImpl, new URL(`../../fonts/file?resourceId=${encodeURIComponent(resourceId)}`, base), { signal: options?.signal });
        return new Uint8Array(await response.arrayBuffer());
      },
    };
    const attachmentStore = {
      async store({ bucket, name, mime, file, signal }) {
        const extension = /\.([a-z0-9]{1,8})$/i.exec(name ?? '')?.[1]?.toLowerCase();
        if (!extension) throw new Error('This file type cannot be inserted.');
        const kind = attachmentKind(editorModule, meta.suite, bucket, mime);
        const response = await call(fetchImpl, new URL(`media?extension=${extension}${kind ? `&kind=${kind}` : ''}`, base), { method: 'POST', body: await file.arrayBuffer(), signal, headers: { 'content-type': 'application/octet-stream' } });
        return { resourceRef: (await response.json()).reference };
      },
      async cancel(resourceRef) { await call(fetchImpl, new URL(`media?reference=${encodeURIComponent(resourceRef)}`, base), { method: 'DELETE' }); },
    };

    let editor, agent, selection, chrome, rebuilding = null, dirtyNotified = false, printing = false, historyOff, lastHistory;
    // Until someone acts on the document — the user's input in the page, or an AI or toolbar operation — a change is
    // the editor's own load-time housekeeping (the Sheet SDK clears its calc-csr data after the first render). Kit
    // takes it to stay in step with the editor, but it is no edit: nothing turns unsaved and nothing is written back.
    let acted = false;
    const act = event => { if (event?.isTrusted !== false) acted = true; };
    for (const type of ACT_EVENTS) window.addEventListener?.(type, act, { capture: true, signal: lifetime.signal });
    // A press in the editor (its ribbon keeps focus, so the parent sees no blur) closes the P area's open menus and
    // the update bubble: the parent cannot listen inside this cross-origin page, so the page says so. Nothing is held.
    window.addEventListener?.('pointerdown', event => { if (event?.isTrusted !== false) channel.notify('press'); }, { capture: true, signal: lifetime.signal });
    // PPT's editor history, for the toolbar's undo/redo (Writer and Sheet do not expose theirs). The availability
    // getters are the editor's own, read after each state change; without them nothing is announced.
    const presentationInternals = () => { try { return meta.suite === 'presentation' ? editor?.api?.presentation?.editorInternals?.() : undefined; } catch { return undefined; } };
    const pushHistory = () => {
      const document = presentationInternals()?.getDocument?.();
      if (!document || !('canUndo' in document)) return;
      const next = { canUndo: meta.mode === 'edit' && document.canUndo === true, canRedo: meta.mode === 'edit' && document.canRedo === true };
      if (lastHistory?.canUndo === next.canUndo && lastHistory?.canRedo === next.canRedo) return;
      lastHistory = next;
      channel.notify('history', next);
    };
    const watchHistory = () => {
      historyOff?.();
      historyOff = undefined;
      lastHistory = undefined;
      const internals = presentationInternals();
      if (typeof internals?.on !== 'function') return;
      const offs = ['state-change', 'commit'].map(event => internals.on(event, pushHistory));
      historyOff = () => { for (const off of offs) off?.(); };
      pushHistory();
    };
    const queue = createChangeQueue({
      send: (seq, change, tag) => call(fetchImpl, new URL('change', base), { method: 'POST', headers: { 'content-type': 'application/json' }, body: JSON.stringify({ seq, format: change.format, body: change.body, ...(tag?.inert && { inert: true }) }) }),
      // The editor's own load-time change is in Kit now, but the document is no more edited than before.
      onAccepted: (number, tag) => { if (tag?.inert) return; void selection?.contentChanged(); channel.notify('dirty', { dirty: true, editRevision: queue.sent, ...(tag?.manual && { manual: true }) }); dirtyNotified = true; },
      onRebuild: error => { void rebuild(error); },
    });

    async function mount() {
      const content = await call(fetchImpl, new URL('content', base));
      const nextChangeSetRevision = Number(content.headers.get('x-next-change-set-revision'));
      const text = await content.text();
      const { runtimeScripts, ...suiteOptions } = suiteMountOptions({ suite: meta.suite, mode: meta.mode, user: meta.user, nextChangeSetRevision,
        language: locale === 'en' ? 'en-US' : 'zh-CN' });
      const created = await editorModule.createOfflineEditor({
        container: root, suite: meta.suite, content: text, mode: meta.mode === 'edit' ? 'edit' : 'readonly', ...suiteOptions,
        file: { name: meta.name }, direction: 'ltr', fontProvider,
        resources: { baseUrl: resourcesBase, catalog, ...(runtimeScripts && { runtimeScripts }), signal: lifetime.signal },
        assets: { baseUrl: mediaBase.href, resolveReadUrl: mediaUrl },
        ...(meta.mode === 'edit' && { attachment: { store: attachmentStore } }),
      });
      // Whether a change is the user's is decided when the editor makes it, not when the Host accepts it.
      created.on('change', change => {
        const applying = agent?.applying === true;
        if (applying) acted = true;
        const inert = !acted;
        const manual = !applying && !inert;
        if (manual) agent?.handOff?.();
        queue.push(change, { manual, ...(inert && { inert: true }) });
      });
      // A fault while the editor is still mounting fails the mount itself (start() rejects); only a mounted editor rebuilds.
      created.on('transferFault', () => { if (editor === created) void rebuild(new PageError('unavailable', 'A change could not be delivered.', { rebuild: true })); });
      created.on('saveIntent', () => channel.notify('save-request', { editRevision: queue.sent }));
      await created.start();
      editor = created;
      if (meta.mode === 'edit' && window.MutationObserver) {
        chrome?.dispose();
        chrome = installEditorChrome({ root, window, document, suite: meta.suite, locale, signal: lifetime.signal, onWebsite:()=>channel.notify('open-website'), onGeometry:()=>selection?.layoutChanged(), onStatusInset: dockBottom => channel.notify('document-ui', { dockBottom }) });
      }
      if (meta.suite === 'document') {
        // Limit the font correction to our meeting template; it survives opening an already saved file.
        try {
          await editor.api.writer.Word.run(async ctx => {
            const styles = ctx.document.getStyles(); styles.load('items/nameLocal'); await ctx.sync();
            if (Array.from(styles.items).some(item => String(item.nameLocal) === 'OfficeSDK Meeting Font'))
              root.setAttribute('data-meeting-font-fix', '');
          });
        } catch (error) { console.warn('officesdk: meeting font detection unavailable', error); }
        void fitWholeWriterPage(document);
      }
      watchHistory();
      agent = createAgent?.({ editor, document, channel, view: base, locale });
      if (meta.mode === 'edit' && document.body && agent && (meta.suite === 'presentation' || typeof agent.readSelection === 'function')) {
        selection = createSelectionFeature({ ui: createSelectionUi(document, { locale }), agent, window, document, reviews,
          allowSelection: typeof agent.readSelection === 'function', notify: (event, payload) => channel.notify(event, payload) });
      }
    }

    /** The Host refused a change or Kit restarted: reload from the Host and mount a new editor. */
    function rebuild(cause) {
      rebuilding ??= (async () => {
        try {
          selection?.dispose();
          selection = undefined;
          chrome?.dispose();
          chrome = undefined;
          agent?.dispose?.();
          agent = undefined;
          historyOff?.();
          historyOff = undefined;
          await editor?.destroy().catch(() => {});
          editor = undefined;
          root.replaceChildren();
          // A change still on its way must land (or fail) before the Host restarts the sequence at 0.
          await Promise.race([queue.idle(), new Promise(resolve => setTimeout(resolve, 5000))]);
          queue.restart();
          await mount();
          // Untouched, the reloaded editor holds no edit of anyone's: the document stays clean.
          if (acted || dirtyNotified) channel.notify('dirty', { dirty: true, editRevision: 0, manual: true, recovered: true });
        } catch (error) {
          // The view itself is gone (Kit restarted): the parent reopens the document.
          channel.notify('failed', { message: error?.message ?? 'The editor could not recover.', rebuild: true });
        } finally { rebuilding = null; }
      })();
      return rebuilding;
    }

    const delivered = async running => {
      const result = await running;
      await queue.idle();
      if (queue.stopped) throw new PageError('stage-failed', 'The edit did not reach the document service; the editor is reloading.');
      // The edit has already reached Kit. Report storage errors in the review
      // sidebar without making the model repeat an applied document mutation.
      await reviews?.flush().catch(() => {});
      return { ...result, editRevision: queue.sent };
    };

    // Model calls can overlap. Mutations of a single live document are validated and committed in order of arrival.
    let writes=Promise.resolve();
    const commit=(signal,run)=>{
      const result=writes.then(()=>{
        if(signal?.aborted) throw new PageError('cancelled','This edit was cancelled.');
        return run();
      });
      writes=result.catch(()=>{});
      return result;
    };
    serveAgent(channel, {
      // Save and export first take what the editor still holds: settle input, then every change reaches the Host.
      // `settle: false` (autosave) only waits for what was already sent: an IME composition or a cell being
      // edited stays with the user.
      async flush(args) {
        if (rebuilding) await rebuilding;
        if (!editor) throw new PageError('no-editor', 'The editor is not ready.');
        if (meta.mode === 'edit' && args?.settle !== false) await editor.prepare();
        await queue.idle();
        await reviews?.flush();
        if (queue.stopped) throw new PageError('stage-failed', 'The latest changes did not reach the document service; the editor is reloading.');
        return { editRevision: queue.sent };
      },
      // The page's own model, to compare with what Kit composed from the forwarded changes.
      async snapshot() {
        if (!editor) throw new PageError('no-editor', 'The editor is not ready.');
        const api = meta.suite === 'presentation' ? editor.api.presentation : meta.suite === 'document' ? editor.api.writer : undefined;
        if (!api?.exportSnapshot) throw new PageError('unsupported-operation', 'This editor has no model snapshot.');
        return { content: await api.exportSnapshot() };
      },
      async inspect(args, abort) {
        if (!agent) throw new PageError('unsupported-operation', 'This editor has no AI support.');
        if (args?.editTarget && meta.suite === 'document') {
          if (!selection) throw new PageError('read-only', 'Open this document for editing.');
          return selection.inspectEditTarget(args.editTarget, abort);
        }
        return agent.inspect(args);
      },
      // An AI edit or undo returns once its changes reached the Host, with the page's change count.
      apply: (args, abort) => {
        acted = true;
        return meta.mode !== 'edit' ? Promise.reject(new PageError('read-only', `This ${SUITE_NOUN[meta.suite] ?? 'document'} is open read only.`))
          : agent ? commit(abort,()=>delivered(agent.apply(args, abort, undefined, selection?.takeChatPreview()))) : Promise.reject(new PageError('unsupported-operation', 'This editor has no AI support.'));
      },
      // Selection rewrite: the conversation's state, the answer to a request, and the model's replacement.
      composer: async args => selection ? selection.composer(args) : { ok: false },
      documentUi: async args => {
        if (args?.beginTableStage && meta.mode === 'edit' && agent?.beginTableStage) { await agent.beginTableStage(); return {ok:true}; }
        if (args?.fillProjectTable && meta.mode === 'edit' && agent?.fillProjectTable) { await agent.fillProjectTable(args.fillProjectTable, args.tableKind); return {ok:true}; }
        if (args?.freezeTableStage && meta.mode === 'edit' && agent?.freezeTableStage) { await agent.freezeTableStage(); return {ok:true}; }
        if (args?.endTableStage && agent?.endTableStage) { agent.endTableStage(); return {ok:true}; }
        if (args?.beginCreation && meta.mode === 'edit' && agent?.beginCreation) {
          const start = agent.beginCreation(args.beginCreation);
          await selection?.refresh?.();
          await start;
          return {ok:true};
        }
        if (args?.endGeneration) {
          if (meta.suite === 'presentation' || agent?.creating) await commit(undefined, () => agent?.endGeneration?.());
          return {ok:true};
        }
        if (args?.retryReviews) await reviews?.flush();
        if (typeof args?.focusMode === 'boolean') chrome?.focusMode(args.focusMode);
        return selection?.hydrate(args) ?? { ok: false };
      },
      changeDetail: async (args, abort) => {
        if (meta.mode !== 'edit' || !selection) throw new PageError('read-only', 'Open this document for editing.');
        if (args.action === 'remove') { acted = true; return delivered(selection.changeDetail(args, abort)); }
        if (args.action === 'accept') {
          const result=await selection.changeDetail(args, abort);
          await reviews?.flush().catch(() => {});
          return result;
        }
        return selection.changeDetail(args, abort);
      },
      readSelection: async () => {
        const selected = meta.mode === 'edit' ? await agent?.readSelection?.() : undefined;
        return selected ? { original: selected.original, unsupported: selected.unsupported } : null;
      },
      submitChatRegion: args => selection?.submitChatRegion(args) ?? {handled:false},
      selectionStatus: async args => selection ? selection.status(args) : { ok: false },
      rewrite: (args, abort) => {
        acted = true;
        return meta.mode !== 'edit' ? Promise.reject(new PageError('read-only', `This ${SUITE_NOUN[meta.suite] ?? 'document'} is open read only.`))
          : selection && typeof agent?.rewriteSelection === 'function' ? commit(abort,()=>delivered(selection.rewrite(args, abort)))
            : typeof agent?.rewriteSelection === 'function' ? delivered(agent.rewriteSelection(args, abort).then(outcome => outcome.reply))
              : Promise.reject(new PageError('unsupported-operation', 'This editor cannot rewrite a selection.'));
      },
      undo: args => {
        acted = true;
        return meta.mode !== 'edit' ? Promise.reject(new PageError('read-only', `This ${SUITE_NOUN[meta.suite] ?? 'document'} is open read only.`))
          : agent ? delivered(agent.undo(args)) : Promise.reject(new PageError('unsupported-operation', 'This editor has no AI support.'));
      },
      // The toolbar's undo/redo: the editor's own history, never while an AI edit is writing (its rollback needs a
      // stable baseline). Returns once the resulting changes reached the Host.
      history: args => {
        acted = true;
        if (meta.mode !== 'edit') return Promise.reject(new PageError('read-only', 'This document is open read only.'));
        if (agent?.applying) return Promise.reject(new PageError('busy', 'An AI edit is still running.'));
        const internals = presentationInternals();
        const action = args?.action;
        if (!['undo', 'redo'].includes(action) || typeof internals?.[action] !== 'function') return Promise.reject(new PageError('unsupported-operation', 'This editor cannot do that.'));
        return delivered(Promise.resolve().then(() => { internals[action](); return {}; }));
      },
      // The browser's print dialog for the page, opened after replying: it blocks the page until it closes. One at a
      // time. Writer keeps a print-ready copy of every page; Sheet and PPT print only through their own (disabled)
      // print flows, so the panel offers Print for DOCX alone.
      print: async () => {
        if (!editor) throw new PageError('no-editor', 'The editor is not ready.');
        if (printing) return { ignored: true };
        // Done when the dialog closes (afterprint), whether or not print() blocked; a minute at most.
        printing = true;
        let timer;
        const done = () => { printing = false; clearTimeout(timer); window.removeEventListener('afterprint', done); };
        window.addEventListener('afterprint', done, { once: true });
        timer = setTimeout(done, 60_000);
        setTimeout(() => { try { window.print(); } catch (error) { console.error('print failed', error); done(); } });
        return {};
      },
    }, { signal: lifetime.signal });

    await mount();
    if (initiallyBlank) {
      await agent.beginCreation('page');
      if (['project','meeting','proposal'].includes(generationMode)) { await agent.beginTableStage(); if(generationMode === 'meeting') agent.beginMeetingCreation(); if(generationMode==='proposal') agent.beginProposalCreation(); }
      else agent.beginTemplateCreation(generationMode);
      root.style.visibility = '';
    }
    if (initiallyBlankDeck) { await agent.beginExistingTemplate(JSON.parse(params.get('slideProgress') || '{}')); root.style.visibility = ''; }
    channel.notify('ready', { mode: meta.mode, ...(meta.diskChanged && { diskChanged: true }) });
    // Unloading tears the editor down with the page; destroying it during pagehide only trips the SDK's own disposal.
    window.addEventListener('pagehide', () => lifetime.abort(), { once: true });
    return { get editor() { return editor; }, queue, rebuild, channel, get dirty() { return dirtyNotified; } };
  } catch (error) {
    // No editor for a live view (bundle, resources, meta or mount failed): close the view without autosaving so
    // the tab can fall back to a read-only preview. A gone view (rebuild) is the parent's to reopen.
    let unsaved;
    if (!error?.rebuild) unsaved = await call(fetchImpl, new URL('abandon', base), { method: 'POST' }).then(response => response.json(), () => undefined).then(result => result?.unsaved);
    channel.notify('failed', { message: error?.message ?? 'The editor could not open.', ...(error?.rebuild && { rebuild: true }), ...(unsaved && { unsaved: true }) });
    throw error;
  }
}
