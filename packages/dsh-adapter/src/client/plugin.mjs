import {openOfficeWebsite} from '../../../core/src/browser/office-brand.mjs';
import { registerDeckChat } from './deck-chat.mjs';
import { createEditSurfaces } from './edit-surfaces.mjs';
import { DocumentController } from './controller.mjs';
import { failureKind, fileAddress, parseSource } from './address.mjs';
import { baseName, isAbsolute, joinPath, parentOf, segments } from './paths.mjs';
import { AGENT_FORMATS, AgentBridge, ControllerRegistry, deckOpener } from './agent.mjs';
import { createEditorFrame } from './editor-frame.mjs';
import { createToolViewStore, registerToolViews } from './tool-views.mjs';
import { registerStylePicker } from './style-picker.mjs';
import { registerStyleFlows } from './style-flows.mjs';
import { registerDocumentContext } from './document-context.mjs';
import { createWhatsNewWatch, fill, whatsNewView } from './whats-new.mjs';
import { createComposerRegistry, handleSelectionRequest, pushComposerState, registerSelectionComposer } from './selection-bridge.mjs';
import { anyUnsaved, cancelKey, copyName, createToolbarPrefs, createUpdateWatch, editButtons, menuKey, modeButtons, nameProblem, pickedPath, switchPlan, offeredModules, placeMenu, rovingKey, toolbarBusy, updateEntry, updatePanelView } from './toolbar.mjs';
import { mountStyles } from './styles.mjs';
import { DEMO_STATES, accountGate, accountHeld, accountMenuKey, createAccountSource, demoState, heldButtons } from './account.mjs';

export const inject = ['documentPreviews', 'slots', 'locale', 'connection'];
const NS = 'shimoPreview';
// One literal on purpose: a trimmed test package could narrow it to the suites it ships.
const EXTENSIONS = 'docx,xlsx,pptx'.split(',');
const MESSAGES = {
  en: { title: 'OfficeSDK', lang: 'en', loading: 'Loading document…', edit: 'Edit', view: 'Read only', save: 'Save', saving: 'Saving…', export: 'Download', drafts: 'Drafts',
    diverged: 'The file changed outside this editor. Discard your edits and reload, or keep them as a copy.', discard: 'Discard and reload', keepCopy: 'Download my version',
    conflict: 'The file changed before saving. Your edits were kept as a draft.', retry: 'Try again', noDrafts: 'No drafts', unsaved: 'Unsaved changes',
    proposalTitle: 'AI wants to change this presentation', apply: 'Apply', decline: 'Discard', danger: 'This edit deletes content or changes many slides. Check it before applying.',
    focusMode: 'Focus mode', exitFocus: 'Exit focus mode', generating: 'AI is generating this file', stopGenerating: 'Stop', generationStopped: 'Stopped: AI edits ask first again',
    unsavedBadge: 'Unsaved', unsavedApprovalBadge: 'Unsaved (editing not allowed yet)', unstagedBadge: 'Latest change not handed to DSH yet — retrying', saveOutside: 'Allow editing the original file? Once allowed, changes in this tab are saved to it automatically.', saveConfirm: 'Allow editing', authorize: 'Allow editing', authorizeTitle: 'Changes in this tab will be saved to the original file automatically', authorized: 'Editing allowed', authorizedTitle: 'Changes are saved to the original file automatically', readOnlyTitle: 'This Session is read-only; switch it to a writable mode to edit the original file', authorizeHint: 'You are changing this file, but editing the original file is not allowed yet, so nothing is saved. Once allowed, changes are saved to it automatically.', readOnlyHint: 'This Session is read-only: changes here are not saved to the original file.', later: 'Later', updateTo: 'Update to {version}', updating: 'Updating…', modeEdit: 'Edit mode', modeView: 'Read mode', modeSelected: '{mode} · selected', modeSwitch: '{mode} · click to switch', modeBlocked: 'Some changes are not saved yet, so the document stays in edit mode.', viewBlockedTitle: 'Unsaved changes: allow editing so they are saved, then switch to read mode', updateTitle: 'Plugin update', updateNew: 'New version', updateAvailable: 'New version v{version} found — plugin update', updateCurrent: 'Current version', updateLatest: 'New version', updateNow: 'Update', updateUnsaved: 'Some documents have unsaved changes. Save them before updating.', updateKeeps: 'Open documents reload with the new version.', updateInstalled: 'Updated to {version}', updateInstalledTitle: 'Restart DSH if the new version does not load by itself', updateFailed: 'Update failed', updatingLabel: 'Updating plugin', updateReadyLabel: 'Plugin update ready', updateProduct: 'OfficeSDK v{version}', updatePreparing: 'Updating', updatePreparingNote: 'Preparing the update…', updateGotIt: 'Got it', updateReadyRestart: 'The update is ready. Save your files, then restart DSH to use it.', updateReadyLoaded: 'The update is ready; the new version is loading.', updateNotes: 'What’s new', updateNotesFallback: 'The full notes appear in “What’s new” once the update is installed. Open documents reload with the new version.', toolbar: 'Quick access', customize: 'Customize Quick Access Toolbar', module_undo: 'Undo', module_redo: 'Redo', module_save: 'Save', module_saveAs: 'Save As', module_folder: 'Open local folder', module_print: 'Print', module_layout: 'Ribbon layout', module_mode: 'Edit / view mode', module_update: 'Plugin update', module_account: 'Account', busyTitle: 'Wait until the current save finishes', saveViewTitle: 'Nothing to save in read mode', accountConnected: 'Connected account', accountConnect: 'Connect account', accountReconnect: 'Reconnect', accountUsage: 'Usage', accountCredits: '{n} credits', accountUsageUnavailable: 'Unavailable — retry', accountWebsite: 'Visit officesdk.com', accountBlocked: 'The account is not connected: editing and saving wait until it is.', module_accountPreview: 'Account (preview)', accountDemo: 'Demo data', accountDemoStates: 'Demo state', demo_connected: 'Connected', demo_disconnected: 'Not connected', demo_expired: 'Expired', demo_zero: 'No credits left', demo_unknown: 'Usage unavailable', saveAsTitle: 'Save As', saveAsName: 'Save as', saveAsFormat: 'Format', saveAsWhere: 'Where', saveAsChoose: 'Choose…', saveAsOutside: 'This place is outside the workspace; saving allows writing here.', saveAsSave: 'Save', copySuffix: ' copy', nameEmpty: 'Enter a file name.', nameCharacters: 'A file name cannot contain / \\ : * ? " < > |', nameDot: 'A file name cannot start with a dot.', nameSame: 'That is this document itself; use Save instead.', targetBusy: 'Another program has that file open. Close it there, then save again.', replacePrompt: '“{name}” already exists. Replace it?', replace: 'Replace', saveAsFailed: 'Not saved: ', saveAsAction: 'Save As…', openTitle: 'Open', openAction: 'Open', openFailed: 'Cannot open: ', openUnsupported: 'Only Word, Excel and PowerPoint files open here.', folderUp: 'Up', workspace: 'Workspace', emptyFolder: 'No documents here', keptEditing: 'New edits kept arriving while saving; try again.', saveAsKept: 'The copy is saved; the original kept its unsaved edits in the editor record: ', savedLabel: 'Saved', autosaveFailed: 'Not saved', cancel: 'Cancel', saveFailed: 'Not saved: ', diskChanged: 'The file changed on disk since your unsaved edits.', useDisk: 'Use the disk version', keepEdits: 'Keep my edits', overwritePrompt: 'Another program changed the file. Replace it with this version?', overwriteConfirm: 'Replace', whatsNewTag: 'What’s new', whatsNewFresh: 'What’s new in v{current}', whatsNewRange: 'Updated from v{previous} to v{current}', whatsNewMore: 'Show all ({n})', whatsNewLess: 'Show less', whatsNewOk: 'Got it', engineDownloading: 'Downloading the document engine (first use) {received}/{total} MB…', engineFallback: ' The npm registry is unreachable; downloading from another source.', engineTesting: ' (this platform is in test)', skillsMenu: 'Skills…', skillsTitle: 'Skills', skillsDownload: 'Download', skillsClear: 'Clear cache', skillsCheck: 'Check for updates', skillsDownloadAll: 'Download all', skillsDone: 'Done', skillState_ready: 'downloaded', skillState_missing: 'downloads on first use', skillState_update: 'update available', skillState_downloading: 'downloading…', skillState_failed: 'download failed', skillState_bundled: 'included',
    fallbackPreview: 'The editor is unavailable; showing a read-only preview.', fallbackUnsaved: 'The editor is unavailable; showing a read-only preview of the file on disk. Your unsaved edits are still kept in the edit record.',
    op_setText: 'Text', op_setFont: 'Font', op_setAlignment: 'Alignment', op_setFill: 'Fill', op_setGeometry: 'Position/size', op_addTextBox: 'New text box',
    op_deleteShape: 'Delete shape', op_addSlide: 'New slide', op_deleteSlide: 'Delete slide', op_moveSlide: 'Move slide',
    proposalTitleDocument: 'AI wants to change this document', dangerDocument: 'This edit deletes paragraphs, and AI edits to Word cannot be undone automatically. Check it before applying.',
    op_insertParagraph: 'New paragraph', op_replaceText: 'Replace text', op_formatParagraph: 'Format', op_deleteParagraph: 'Delete paragraph', op_insertTable: 'New table', op_setCell: 'Cell', op_insertPageBreak: 'Page break',
    proposalTitleSpreadsheet: 'AI wants to change this spreadsheet', dangerSpreadsheet: 'This edit clears cells or deletes rows or columns, and AI edits to spreadsheets cannot be undone automatically. Check it before applying.',
    op_setValues: 'Values', op_setFormula: 'Formula', op_formatRange: 'Format', op_clearRange: 'Clear', op_addSheet: 'New sheet', op_renameSheet: 'Rename sheet', op_activateSheet: 'Switch sheet', op_insertRows: 'Insert rows', op_deleteRows: 'Delete rows', op_insertColumns: 'Insert columns', op_deleteColumns: 'Delete columns', op_setColumnWidth: 'Column width', op_freezeRows: 'Freeze rows', op_addChart: 'New chart' },
  zh: { title: 'OfficeSDK', lang: 'zh', loading: '正在加载文档…', edit: '编辑', view: '只读', save: '保存', saving: '正在保存…', export: '下载', drafts: '草稿',
    diverged: '文件已在外部修改。可以放弃修改并重新加载，或保留你的版本。', discard: '放弃修改并重新加载', keepCopy: '下载我的版本',
    conflict: '保存前文件已变化，你的修改已保存为草稿。', retry: '重试', noDrafts: '没有草稿', unsaved: '有未保存的修改',
    proposalTitle: 'AI 想修改这份演示文稿', apply: '应用', decline: '放弃', danger: '这次修改会删除内容或改动大量页面，请确认后再应用。',
    focusMode: '专注模式', exitFocus: '退出专注模式', generating: 'AI 生成中', stopGenerating: '停止', generationStopped: '已停止，之后 AI 的修改会先问你',
    unsavedBadge: '未保存', unsavedApprovalBadge: '未保存（未授权修改）', unstagedBadge: '最新修改还没交给 DSH，正在重试', saveOutside: '允许修改原文件吗？允许后，这个标签里的修改会自动保存到原文件。', saveConfirm: '允许修改', authorize: '授权修改', authorizeTitle: '授权后，这个标签里的修改会自动保存到原文件', authorized: '已授权修改', authorizedTitle: '修改会自动保存到原文件', readOnlyTitle: '当前会话是只读的，切换为可写模式后才能修改原文件', authorizeHint: '你已经修改了这个文件，但还没有授权修改原文件，修改不会保存。授权后会自动保存到原文件。', readOnlyHint: '当前会话是只读的，这里的修改不会保存到原文件。', later: '稍后', updateTo: '更新到 {version}', updating: '更新中…', modeEdit: '编辑模式', modeView: '阅读模式', modeSelected: '{mode} · 已选中', modeSwitch: '{mode} · 点击切换', modeBlocked: '还有修改没保存，暂不能切到阅读。', viewBlockedTitle: '有未保存修改，授权保存后才能切到阅读', updateTitle: '插件更新', updateNew: '有新版本', updateAvailable: '发现新版本 v{version}，插件更新', updateCurrent: '当前版本', updateLatest: '新版本', updateNow: '更新', updateUnsaved: '仍有未保存修改，请先保存后更新。', updateKeeps: '更新后打开的文档会用新版本重新加载。', updateInstalled: '已更新到 {version}', updateInstalledTitle: '如未自动生效，请重启 DSH', updateFailed: '更新失败', updatingLabel: '正在更新插件', updateReadyLabel: '插件更新已准备就绪', updateProduct: 'OfficeSDK v{version}', updatePreparing: '正在更新', updatePreparingNote: '正在准备更新…', updateGotIt: '知道了', updateReadyRestart: '更新已准备就绪，保存文件后重启 DSH 生效。', updateReadyLoaded: '更新已准备就绪，新版本正在加载。', updateNotes: '更新内容', updateNotesFallback: '完整的更新说明会在更新完成后的「本次更新」中显示；打开的文档会用新版本重新加载。', toolbar: '快速访问', customize: '自定义快速访问工具栏', module_undo: '撤销', module_redo: '重做', module_save: '保存', module_saveAs: '另存为', module_folder: '打开本地文件夹', module_print: '打印', module_layout: '功能区模式', module_mode: '编辑 / 查看模式', module_update: '插件更新', module_account: '个人中心', busyTitle: '请等当前保存完成', saveViewTitle: '阅读模式下没有要保存的修改', accountConnected: '已连接账户', accountConnect: '连接账户', accountReconnect: '重新连接', accountUsage: '使用情况', accountCredits: '{n} credits', accountUsageUnavailable: '暂不可用，点击重试', accountWebsite: '访问 officesdk.com', accountBlocked: '账户未连接，暂不能编辑和保存。', module_accountPreview: '个人中心（预览）', accountDemo: '演示数据', accountDemoStates: '演示状态', demo_connected: '已连接', demo_disconnected: '未连接', demo_expired: '已失效', demo_zero: '额度为 0', demo_unknown: '用量暂不可用', saveAsTitle: '另存为', saveAsName: '保存为', saveAsFormat: '文件格式', saveAsWhere: '位置', saveAsChoose: '选择…', saveAsOutside: '位置在工作区外，保存即允许写入此位置。', saveAsSave: '保存', copySuffix: ' - 副本', nameEmpty: '请输入文件名。', nameCharacters: '文件名不能包含 / \\ : * ? " < > |', nameDot: '文件名不能以 . 开头。', nameSame: '这是当前文件本身，请直接保存。', targetBusy: '该文件正被其他程序占用，请先在那个程序中关闭它，再重新保存。', replacePrompt: '「{name}」已存在，要替换吗？', replace: '替换', saveAsFailed: '未能另存：', saveAsAction: '另存为…', openTitle: '打开', openAction: '打开', openFailed: '无法打开：', openUnsupported: '只能打开 Word、Excel、PowerPoint 文件。', folderUp: '上一级', workspace: '工作区', emptyFolder: '这里没有可打开的文档', keptEditing: '保存时仍有新的修改，请再试一次。', saveAsKept: '副本已保存；原文件未写入，但未保存的修改仍留在编辑记录中：', savedLabel: '已保存', autosaveFailed: '未能自动保存', cancel: '取消', saveFailed: '未保存：', diskChanged: '磁盘上的文件在你未保存期间被修改过。', useDisk: '使用磁盘版本', keepEdits: '保留我的修改', overwritePrompt: '文件已被其他程序修改，用当前内容覆盖吗？', overwriteConfirm: '覆盖', whatsNewTag: '本次更新', whatsNewFresh: 'v{current} 更新内容', whatsNewRange: '已从 v{previous} 更新到 v{current}', whatsNewMore: '展开全部（共 {n} 条）', whatsNewLess: '收起', whatsNewOk: '知道了', engineDownloading: '首次使用，正在下载文档引擎 {received}/{total} MB…', engineFallback: 'npm 源连不上，已改用其他来源下载。', engineTesting: '（此平台测试中）', skillsMenu: 'Skills 管理…', skillsTitle: 'Skills', skillsDownload: '预下载', skillsClear: '清除缓存', skillsCheck: '检查更新', skillsDownloadAll: '全部预下载', skillsDone: '完成', skillState_ready: '已下载', skillState_missing: '首次使用时下载', skillState_update: '有新版本', skillState_downloading: '下载中…', skillState_failed: '下载失败', skillState_bundled: '随包',
    fallbackPreview: '编辑器不可用，已切换为只读预览。', fallbackUnsaved: '编辑器不可用，已切换为只读预览。未保存的修改仍保留在编辑记录中，预览显示的是磁盘版本。',
    op_setText: '文字', op_setFont: '字体', op_setAlignment: '对齐', op_setFill: '填充', op_setGeometry: '位置/大小', op_addTextBox: '新增文本框',
    op_deleteShape: '删除形状', op_addSlide: '新增页面', op_deleteSlide: '删除页面', op_moveSlide: '移动页面',
    proposalTitleDocument: 'AI 想修改这份文档', dangerDocument: '这次修改会删除段落，Word 的 AI 修改不能自动撤销，请确认后再应用。',
    op_insertParagraph: '新增段落', op_replaceText: '替换文字', op_formatParagraph: '格式', op_deleteParagraph: '删除段落', op_insertTable: '新增表格', op_setCell: '单元格', op_insertPageBreak: '分页',
    proposalTitleSpreadsheet: 'AI 想修改这份表格', dangerSpreadsheet: '这次修改会清空单元格或删除行列，表格的 AI 修改不能自动撤销，请确认后再应用。',
    op_setValues: '写入数值', op_setFormula: '公式', op_formatRange: '格式', op_clearRange: '清空', op_addSheet: '新建工作表', op_renameSheet: '重命名工作表', op_activateSheet: '切换工作表', op_insertRows: '插入行', op_deleteRows: '删除行', op_insertColumns: '插入列', op_deleteColumns: '删除列', op_setColumnWidth: '列宽', op_freezeRows: '冻结行', op_addChart: '新增图表' },
};

/** Body and header action share one controller per tab; AI commands find it by address. */
const controllers = new Map();
// A header that renders before its tab's controller exists (or after a new one replaces it) waits for this.
const controllerWatchers = new Set();
const setController = (tabId, controller) => { controller ? controllers.set(tabId, controller) : controllers.delete(tabId); for (const watcher of controllerWatchers) watcher(); };
const registry = new ControllerRegistry();
// Every format asks for an editor; the Host answers with a read-only view where it has none (Core DOCX/XLSX, no write access).
const EDIT_FORMATS = new Set(['docx', 'xlsx', 'pptx', ...Object.keys(AGENT_FORMATS)]);
/** The format a tab's address names, before any lease says it: the frame's sandbox is fixed from its first load. */
export const sandboxFormat = address => { try { return /\.([A-Za-z0-9]+)$/.exec(parseSource(address).path)?.[1]?.toLowerCase(); } catch { return undefined; } };
const initialMode = address => { try { return EDIT_FORMATS.has(/\.([A-Za-z0-9]+)$/.exec(parseSource(address).path)?.[1]?.toLowerCase()) ? 'edit' : 'view'; } catch { return 'view'; } };

/** Download from the parent document: the sandboxed frame cannot. */
function anchorDownload(document) {
  return (url, name) => { const a = document.createElement('a'); a.href = url; a.download = name ?? ''; a.rel = 'noopener'; document.body.appendChild(a); a.click(); a.remove(); };
}

/** Host React stays external to avoid a second renderer/hooks instance. */
export function install(ctx, React, environment = {}) {
  const { createElement: h, useEffect, useLayoutEffect, useRef, useState, useSyncExternalStore } = React;
  const doc = environment.document ?? globalThis.document;
  const sidebarOf = () => environment.sidebarRight ?? ctx.get?.('sidebarRight');
  // What the AI bridge records while it runs presentation tools (pages, outlines), for their chat rows.
  const toolViews = environment.toolViews ?? createToolViewStore();
  // The account all tabs share (preview, mock data); its hold reaches every tab's controller, shown or not.
  const account = environment.account ?? createAccountSource(undefined, text => console.warn(`officesdk: ${text}`));
  // One background release check for all open tabs, through any tab's session (owner check included).
  const updates = environment.updates ?? createUpdateWatch({ fetch: ({ refresh }) => {
    const controller = [...controllers.values()].find(candidate => candidate.lease) ?? [...controllers.values()][0];
    return controller ? controller.call('updateStatus', { refresh }) : Promise.resolve(undefined);
  } });
  // 「本次更新」 after the plugin changed version: asked once per load through any tab's session; the card appearing marks it seen.
  const anyController = () => [...controllers.values()].find(candidate => candidate.lease) ?? [...controllers.values()][0];
  const whatsNew = environment.whatsNew ?? createWhatsNewWatch({
    fetch: () => anyController()?.call('whatsNew', {}),
    mark: version => anyController()?.call('whatsNewSeen', { version }),
  });
  const holdAll = () => { const held = accountHeld(account.snapshot()); for (const controller of controllers.values()) controller.setAccountHold?.(held); };
  // Shown/hidden toolbar modules, shared by every tab and suite.
  const prefs = environment.toolbarPrefs ?? createToolbarPrefs();
  // The input boxes by Session, for the selection rewrite's message (selection-bridge.mjs).
  const composers = environment.composers ?? createComposerRegistry();
  const { ChangeRail, FocusChat, FocusDecision } = createEditSurfaces(React);
  const documentUis = new Map();
  const tabEditors = new Map();
  ctx.effect(() => () => { for (const editor of tabEditors.values()) editor.dispose(); });
  const useController = tabId => useSyncExternalStore(
    listener => {
      controllerWatchers.add(listener);
      const off = controllers.get(tabId)?.subscribe(listener);
      return () => { controllerWatchers.delete(listener); off?.(); };
    },
    () => controllers.get(tabId)?.state,
  );

  function Document({ resourceAddress, content, useTabInfo, useResource, useSessionStatus, renderFactorySlot, t }) {
    const { tab, sidebar } = useTabInfo();
    const frameContainer = useRef(null);
    const [view, setView] = useState({ loading: true });
    const [frameFullscreen,setFrameFullscreen]=useState(false);
    useEffect(()=>{
      const changed=()=>setFrameFullscreen(Boolean(doc.fullscreenElement && doc.fullscreenElement===controllers.get(tab.id)?.frame));
      changed();
      doc.addEventListener('fullscreenchange',changed);
      return ()=>doc.removeEventListener('fullscreenchange',changed);
    },[resourceAddress,tab.id]);
    const focused=Boolean(sidebar?.fullscreen || frameFullscreen);
    const request = content.kind === 'renderer' ? content : undefined;
    const meta = useResource?.(resourceAddress);
    useLayoutEffect(() => {
      if (!request || !frameContainer.current) return;
      let editor = tabEditors.get(tab.id);
      if (editor && (editor.controller.address !== resourceAddress || editor.controller.signal !== tab.signal)) {
        editor.dispose();
        editor = undefined;
      }
      if (!editor) editor = startEditor(frameContainer.current);
      else editor.surface.attach(frameContainer.current);
      const { controller } = editor;
      setView(controller.state);
      const unsubscribe = controller.subscribe(setView);
      return () => {
        unsubscribe();
        if (tabEditors.get(tab.id) !== editor) return;
        if (!tab.signal || tab.signal.aborted || !editor.surface.park()) editor.dispose();
      };
    }, [resourceAddress, tab.id, tab.signal, Boolean(request)]);

    function startEditor(container) {
      const surface = createEditorFrame(container, { printable: sandboxFormat(resourceAddress) === 'docx', title: t?.('title') || 'OfficeSDK' });
      const controller = new DocumentController({ rpc: ctx.connection.rpc, events: window, base: environment.base || doc.baseURI, address: resourceAddress,
        signal: tab.signal, update: state => surface.update(state), download: anchorDownload(doc), locale: t?.('lang') || undefined });
      const permissionKey = JSON.stringify([controller.source.sessionId, controller.source.path]);
      controller.setAiPermission('always');
      setController(tab.id, controller);
      controller.setAccountHold(accountHeld(account.snapshot()));
      controller.messages = { keptEditing: t?.('keptEditing'), readOnlyHint: t?.('readOnlyHint') };
      controller.frame = surface.frame;
      controller.onOpenWebsite = () => openOfficeWebsite(doc);
      // Selection rewrite: the page learns whether the conversation can take a message, and asks to send one.
      controller.onReady = () => {
        pushComposerState([controller], composers, controller.source.sessionId);
        controller.syncDocumentUi();
      };
      const documentUi = documentUis.get(permissionKey) ?? { draft: '', references: [], pending: [], candidates: [], changes: [], next: 1 };
      documentUis.set(permissionKey, documentUi);
      controller.documentUi = documentUi;
      controller.onSelectionRequest = message => handleSelectionRequest({ controller, message, registry: composers, lang: t?.('lang') === 'en' ? 'en' : 'zh' });
      const unregister = registry.add(resourceAddress, controller);
      const dispose = () => {
        tab.signal?.removeEventListener('abort', dispose);
        unregister();
        if (controllers.get(tab.id) === controller) setController(tab.id);
        tabEditors.delete(tab.id);
        controller.stop();
        surface.dispose();
      };
      const editor = { controller, surface, dispose };
      tabEditors.set(tab.id, editor);
      tab.signal?.addEventListener('abort', dispose, { once: true });
      void controller.open(request.revision, { mode: initialMode(resourceAddress), loaded: request.loaded, failed: request.failed });
      return editor;
    }
    useEffect(() => {
      const controller = controllers.get(tab.id);
      if (!controller || !request || controller.revision === undefined || request.revision === controller.revision) return;
      void controller.adopt(request.revision, meta?.value?.version, { loaded: request.loaded, failed: request.failed });
    }, [request?.revision]);
    useEffect(()=>{
      const controller=controllers.get(tab.id);
      if (!controller) return;
      const chat=composers.get(controller.source.sessionId)?.chat;
      const lastSeq=chat?.nodes?.reduce((last,node)=>Math.max(last,node.seq ?? -1),-1) ?? -1;
      controller.set({focusMode:focused,focusPeriod:focused?{since:Date.now(),lastSeq}:undefined});
      if (controller.state.ready) void controller.agent('documentUi',{focusMode:focused}).catch(()=>{});
    },[tab.id,resourceAddress,focused]);
    const pendingInteraction=useSessionStatus?.(statuses=>statuses?.get?.(controllers.get(tab.id)?.source.sessionId)?.pendingInteraction);
    const running = useSessionStatus?.(statuses => statuses?.get?.(controllers.get(tab.id)?.source.sessionId)?.running);
    const previousTurn = useRef();
    useEffect(() => {
      const controller = controllers.get(tab.id);
      if (previousTurn.current === true && running === false && !pendingInteraction && controller?.state.ready && /\.pptx$/i.test(controller.source?.path ?? '')) {
        void controller.agent('documentUi', {endGeneration:true}).catch(() => {});
      }
      previousTurn.current = running;
    }, [tab.id, running]);
    if (!request) return null;
    const controller = controllers.get(tab.id);
    // A session can still require an answer while its document is read-only or loading.
    const showFocusChat=controller && focused && view.focusPeriod &&
      ((view.ready && view.mode==='edit') || pendingInteraction || view.proposal || view.diverged);
    const overlay = view.diverged
      ? h('div', { role: 'alert', style: showFocusChat ? {padding:12} : banner }, t?.('diverged'),
          h('button', { style: button, onClick: () => controller?.open(request.revision, { mode: 'edit', loaded: request.loaded, failed: request.failed }) }, t?.('discard')),
          h('button', { style: button, onClick: () => controller?.exportDownload() }, t?.('keepCopy')))
      : view.error ? h('div', { role: 'alert', style: cover }, view.error,
          failureKind({ code: view.errorCode }).kind === 'retry' ? h('button', { style: button, onClick: () => request.reload?.() }, t?.('retry')) : null)
      : view.loading ? h('div', { role: 'status', 'aria-live': 'polite', style: { ...cover, display: 'flex', alignItems: 'center', justifyContent: 'center', pointerEvents: 'none', color: 'var(--dsw-alias-label-secondary,#666)' } }, view.engine?.total ? (t?.('engineDownloading') || 'Downloading the document engine {received}/{total} MB').replace('{received}', String(Math.round(view.engine.received / 1048576))).replace('{total}', String(Math.round(view.engine.total / 1048576))) + (view.engine.testing ? t?.('engineTesting') || ' (this platform is in test)' : '') + (view.engine.fallback ? t?.('engineFallback') || ' The npm registry is unreachable; downloading from another source.' : '') : t?.('loading') || 'Loading document…')
      : null;
    const proposal = view.proposal ? h(Proposal, { proposal: view.proposal, inline:!!showFocusChat, t, decide: approved => controller?.decide(approved), alwaysApply: () => controller?.setAiPermission('always') }) : null;
    return h('div', { style: { position: 'relative', width: '100%', height: '100%', minHeight: 0, overflow: 'hidden', display: 'flex' }, 'data-shimo-document': true },
      h('div', { className:'office-editor-pane' }, h('div', {ref:frameContainer,style:{width:'100%',height:'100%'}}),
      // The last save before switching files: the editor takes no input until it is done.
      view.locked ? h('div', { 'data-shimo-lock': true, style: { position: 'absolute', inset: 0, cursor: 'progress' } }) : null,
      showFocusChat && view.diverged ? h(FocusDecision,{controller},overlay) : overlay, !showFocusChat && proposal,
      showFocusChat && h(FocusChat,{key:resourceAddress,controller,period:view.focusPeriod,frameFullscreen,registry:composers,openDeck:path=>deckOpener(sidebarOf,hostPlatform).openDeck(controller.source.sessionId,path),renderFactorySlot,pendingInteraction,proposal,bottom:view.dockBottom,en:t?.('lang')==='en'})),
      view.ready && view.mode==='edit' && h(ChangeRail,{key:resourceAddress,controller,revealed:view.revealedChange,anchors:view.changeAnchors,records:view.changes ?? controller?.documentUi?.changes,en:t?.('lang')==='en'}));
  }

  /** The AI's change, shown in the tab it applies to: where, what was there, what will be. */
  function Proposal({ proposal, inline, t, decide, alwaysApply }) {
    const apply = useRef(null);
    useEffect(() => { if (!proposal.danger) apply.current?.focus(); }, [proposal]);
    const title = t?.(proposal.kind === 'document' ? 'proposalTitleDocument' : proposal.kind === 'spreadsheet' ? 'proposalTitleSpreadsheet' : 'proposalTitle');
    return h('div', { role: 'dialog', 'aria-modal': false, 'aria-label': title, 'data-shimo-proposal': true, style: inline ? {...card,position:'static',width:'auto',boxShadow:'none'} : card },
      h('div', { style: { fontWeight: 600 } }, title),
      proposal.summary ? h('div', { style: { color: 'var(--dsw-alias-label-secondary,#555)' } }, proposal.summary) : null,
      proposal.danger ? h('div', { role: 'alert', style: { color: '#c02d2e' } }, t?.(proposal.kind === 'document' ? 'dangerDocument' : proposal.kind === 'spreadsheet' ? 'dangerSpreadsheet' : 'danger')) : null,
      h('div', { style: { maxHeight: 220, overflow: 'auto', display: 'grid', gap: 4 } }, proposal.rows.map((row, index) => h('div', { key: index, style: rowStyle },
        h('span', { style: { color: 'var(--dsw-alias-label-secondary,#666)' } }, `${t?.(`op_${row.op}`) ?? row.op}${row.where ? ` · ${row.where}` : ''}`),
        row.before !== undefined ? h('del', { style: { color: '#b42318', whiteSpace: 'pre-wrap' } }, row.before) : null,
        row.after !== undefined ? h('ins', { style: { color: '#067647', textDecoration: 'none', whiteSpace: 'pre-wrap' } }, row.after) : null))),
      h('div', { style: { display: 'flex', gap: 8, justifyContent: 'flex-end' } },
        h('button', { style: button, onClick: () => decide(false) }, t?.('decline')),
        h('button', { ref: apply, style: proposal.danger ? { ...button, borderColor: '#c02d2e', color: '#c02d2e' } : { ...button, background: '#17181a', borderColor: '#17181a', color: '#fff' }, onClick: () => decide(true) }, t?.('apply')),
        h('button', { style: button, onClick: alwaysApply }, t?.('lang') === 'en' ? 'Always apply' : '始终应用')));
  }

  /**
   * The header row: a status area (authorization, save notices, downloads) and the P toolbar — modules in the design
   * package's fixed order, then the three-dot menu that shows or hides them. Colors come from the stylesheet and follow
   * DSH's theme without re-rendering.
   */
  function Action({ useTabInfo, t }) {
    const { tab, sidebar: rightPane, panel } = useTabInfo();
    const state = useController(tab.id);
    const [busy, setBusy] = useState(false);
    const [drafts, setDrafts] = useState();
    const [hintClosed, setHintClosed] = useState(false);
    // The shared check's answer, unless this header is installing (or installed) an update itself.
    const [installing, setUpdate] = useState();
    const checked = useSyncExternalStore(updates.subscribe, updates.snapshot);
    const update = installing ? { ...checked, ...installing } : checked;
    // One user-opened popover per header: 'more' or 'update'. Dialogs the tab raises itself may sit beside it.
    const [open, setOpen] = useState(null);
    const [stop, setStop] = useState('more');
    const hidden = useSyncExternalStore(prefs.subscribe, prefs.hidden);
    const accountNow = useSyncExternalStore(account.subscribe, account.snapshot);
    const root = useRef(null);
    const bar = useRef(null);
    const triggers = { more: useRef(null), update: useRef(null), saveAs: useRef(null), folder: useRef(null), browse: useRef(null), account: useRef(null) };
    // A toolbar button that had focus and disappears (its module hidden, the update installed) hands focus to the
    // three-dot button. Read from the DOM at render time, not focus events, which an unfocused window does not fire.
    const focusedBefore = useRef(null);
    focusedBefore.current = bar.current?.contains(doc?.activeElement) ? doc.activeElement : null;
    const controller = controllers.get(tab.id);
    useLayoutEffect(() => {
      const was = focusedBefore.current;
      const active = doc?.activeElement;
      if (was && !was.isConnected && (!active || active === doc.body)) triggers.more.current?.focus();
    });
    // A newer release on the install branch; DSH's HMR loads it right after the install (else on restart). Checked when a
    // tab's controller appears and every ten minutes while tabs are open; it never opens the popover or takes focus.
    useEffect(() => { if (controller) void updates.check(); }, [controller]);
    // A press in the editor page closes the update bubble and the toolbar's menus, as an outside click does (V3.0 P区).
    useEffect(() => controller?.onPress?.(() => setOpen(current => (['update', 'more', 'account'].includes(current) ? null : current))), [controller]);
    const notes = useSyncExternalStore(whatsNew.subscribe, whatsNew.snapshot);
    useEffect(() => { if (controller) void whatsNew.check(); }, [controller]);
    if (!controller || !state?.capabilities) return null;
    const close = restore => { const trigger = triggers[open]?.current; setOpen(null); if (restore) trigger?.focus(); };
    // An update reloads every open tab: it waits until none has edits still to write (checked again on the click).
    // The bubble stays open through the install and shows its progress; closing it (outside click, Esc) does not cancel it.
    const installUpdate = async () => {
      if (anyUnsaved(controllers)) return setOpen('update');
      setUpdate({ updating: true });
      // Afterwards every tab asks the branch again: the installed copy (or a newer release) shows in all of them; this
      // header keeps only what it saw itself (the result, or why the install failed).
      // Ready only when a newer copy was really installed (it reloads now or after a restart); otherwise nothing is pending.
      try {
        const result = await controller.call('updatePlugin', {});
        const took = result?.restartRequired === true || Boolean(result?.reloading);
        // Nothing newer installed: drop this header's own state, so the shared check decides again.
        setUpdate(took ? { available: false, installed: result.installed, restartRequired: result?.restartRequired === true } : undefined);
      }
      // The Host also counts tabs in other windows: then this one's own check passed, so say why here.
      catch (error) { setUpdate({ failed: error?.code === 'officesdk/unsaved-changes' ? t?.('updateUnsaved') : error.message }); }
      void updates.check(true);
    };
    const entry = updateEntry(update);
    const pBusy = toolbarBusy(busy, state);
    const held = accountHeld(accountNow);
    // While the account holds edits back, the write buttons are off with that reason (Open and Print stay).
    const { edits, modes } = heldButtons(editButtons(state, busy), modeButtons(state, pBusy), state, held);
    const offered = offeredModules({ undo: Boolean(edits.undo), redo: Boolean(edits.redo), save: Boolean(edits.save), saveAs: Boolean(edits.saveAs), folder: Boolean(edits.folder),
      print: Boolean(edits.print), mode: Boolean(modes), update: true, account: true });
    const shows = id => offered.includes(id) && !hidden.includes(id);
    const showMode = shows('mode');
    const showUpdate = shows('update') && entry !== 'hidden';
    const fileKeys = ['undo', 'redo', 'save', 'saveAs', 'folder', 'print'].filter(shows);
    // 个人中心 is a preview with demo data: its own checkbox (the preview switch) shows it, not the hidden list.
    const showAccount = accountNow.preview;
    const showFocus = showMode && state.backend === 'kit' && !panel?.isFloating && typeof sidebarOf()?.toggleFullscreen === 'function';
    const keys = [...fileKeys, ...(showFocus ? ['focus'] : []), ...showMode ? ['edit', 'view'] : [], ...showUpdate ? ['update'] : [], 'more', ...showAccount ? ['account'] : []];
    if (open === 'update' && !showUpdate) setOpen(null);
    if (open === 'saveAs' && (!edits.saveAs || held)) setOpen(null);
    // A save starting closes the account menu (busy), as it does the rest of the toolbar.
    if (open === 'account' && (!showAccount || pBusy)) setOpen(null);
    const focusStop = keys.includes(stop) ? stop : 'more';
    const stopProps = key => ({ 'data-key': key, tabIndex: key === focusStop ? 0 : -1, onFocus: () => setStop(key) });
    const onBarKey = event => {
      const index = keys.indexOf(event.target?.getAttribute?.('data-key'));
      const next = index < 0 ? undefined : rovingKey(event.key, index, keys.length);
      if (next === undefined) return;
      event.preventDefault();
      bar.current?.querySelector(`[data-key="${keys[next]}"]`)?.focus();
    };
    // A save that fails on the way out of edit keeps the tab editing and says why, as the save button does. Pressing a mode
    // by keyboard keeps focus on the toolbar: the reopened editor may grab it, so it comes back unless the user moved it.
    const switchMode = async (mode, button) => {
      const hadFocus = doc.activeElement === button;
      const switching = controller.switchMode(mode);
      const ready = controller.switched;
      try { await switching; await ready; }
      catch (error) { controller.set({ notice: `${t?.('saveFailed') ?? ''}${error.message}` }); }
      const active = doc.activeElement;
      if (hadFocus && (!active || active === doc.body || active.tagName === 'IFRAME' || root.current?.contains(active)))
        root.current?.querySelector('[data-key][aria-pressed="true"]')?.focus();
    };
    const toggleFocus = event => {
      const sidebar=sidebarOf();
      const target=sidebar?.commandTarget?.(event.currentTarget);
      if (target && target.host!=='float') sidebar.toggleFullscreen(target);
    };
    const modeButton = (mode, path, label) => {
      const { pressed, disabled, reason } = modes[mode];
      const blocked = disabled && !pressed;
      const title = reason === 'account' ? t?.('accountBlocked') : reason ? t?.('viewBlockedTitle') : blocked ? t?.('busyTitle') : t?.(pressed ? 'modeSelected' : 'modeSwitch')?.replace('{mode}', t?.(label));
      return h('button', { type: 'button', className: 'sbp-icon', ...stopProps(mode), 'aria-pressed': pressed, 'aria-label': t?.(label), title, 'aria-disabled': blocked || undefined,
        onClick: event => { if (!pressed && !blocked) void switchMode(mode, event.currentTarget); } }, glyph(path));
    };
    const run = async operation => { setBusy(true); try { await operation(); } catch (error) { if (failureKind(error).kind === 'conflict') setDrafts(await controller.listDrafts().catch(() => [])); } finally { setBusy(false); } };
    const caps = state.capabilities;
    // Outside the workspace the Host refuses until the user agrees here; the same click then writes.
    const save = async (approved, overwrite = false) => {
      setBusy(true);
      try { await controller.save({ approved, overwrite }); controller.set({ notice: undefined, saveOutside: undefined, overwritePrompt: undefined }); }
      catch (error) {
        if (error?.code === 'officesdk/sandbox-denied' && error.details?.outside && !approved) controller.set({ saveOutside: true });
        // Kit refuses to replace a file that changed on disk until the user agrees here.
        else if (error?.code === 'officesdk/source-changed' && error.details?.overwrite && !overwrite) controller.set({ overwritePrompt: { approved } });
        else controller.set({ saveOutside: undefined, overwritePrompt: undefined, notice: `${t?.('saveFailed') ?? ''}${error.message}` });
      } finally { setBusy(false); }
    };
    // Kit editing autosaves: the tab asks only whether the original file may be edited.
    const authorize = async () => {
      setBusy(true);
      try { await controller.authorize(); controller.set({ notice: undefined }); }
      catch (error) {
        if (error?.code === 'officesdk/source-changed' && error.details?.overwrite) controller.set({ overwritePrompt: { approved: true } });
        else controller.set({ notice: `${t?.('saveFailed') ?? ''}${error.message}` });
      } finally { setBusy(false); }
    };
    // Dialogs the tab raises by itself: beside the toolbar, without taking focus from the editor.
    const dockDecisions=Boolean(state.focusMode && state.ready && state.mode==='edit');
    const prompt = (text,onCancel,actions,alert=false) => {
      const content=[h('span',{role:alert?'alert':undefined,'aria-live':'polite'},text),h('div',{className:'sbp-actions'},...actions)];
      if (dockDecisions) return h(FocusDecision,{controller},h('div',{className:'office-chat-confirm',role:'alertdialog','aria-label':text,
        onKeyDown:e=>{if(cancelKey(e)){e.stopPropagation();onCancel();}}},...content));
      return h(Popover,{kind:'self',anchor:root,root,className:'sbp-dialog',role:'alertdialog',label:text,onCancel},...content);
    };
    const choice = (label, onClick, primary = false, disabled = false) => h('button', { type: 'button', className: primary ? 'sbp-btn sbp-primary' : 'sbp-btn', disabled, onClick }, label);
    const confirm = (text, onCancel, onConfirm, label, cancelLabel = t?.('cancel')) => prompt(text, onCancel, [choice(cancelLabel, onCancel), choice(label, onConfirm, true, busy)]);
    const editing = state.mode === 'edit';
    const kit = state.backend === 'kit';
    // The toolbar's Save and Cmd/Ctrl+S. A Kit tab not yet allowed to write asks first: allowing also starts autosave.
    // Resolves true once the edits are on disk; a prompt it raises (authorize, outside, overwrite) or a failure is false.
    const saveFromToolbar = async (consent = false, forSwitch = false) => {
      const now = controller.state;
      if (controller.accountHold) { controller.set({ notice: t?.('accountBlocked') }); return false; }
      if (now.backend === 'kit' && !now.authorized) {
        if (now.access === 'denied') controller.set({ saveState: 'failed', saveFailed: t?.('readOnlyHint'), saveFailedSeen: undefined });
        else controller.set({ authorizePrompt: true });
        return false;
      }
      const approved = consent || (now.backend === 'kit' && controller.autosaveApproved === true);
      try {
        const saved = forSwitch ? await controller.saveForSwitch({ approved }) : (await controller.saveNow({ approved }), true);
        if (saved) controller.set({ notice: undefined, saveOutside: undefined, overwritePrompt: undefined });
        return saved;
      } catch (error) {
        if (error?.code === 'officesdk/sandbox-denied' && error.details?.outside) controller.set({ saveOutside: true, saveState: undefined });
        // Kit refuses to replace a file that changed on disk until the user agrees here.
        else if (error?.code === 'officesdk/source-changed' && error.details?.overwrite) controller.set({ overwritePrompt: { approved }, saveState: undefined });
        return false;
      }
    };
    // Another file in this tab (OPEN-01): the current one is saved first, for real, or the tab stays with its edits.
    // The target waits on the controller, bound to this tab and its file; only Retry, Save As or a confirmed prompt
    // carries it on, and Cancel drops it.
    const sessionId = controller.source.sessionId;
    const replaceWith = address => {
      const sidebar = sidebarOf();
      if (!sidebar?.openResource) throw new Error(t?.('openUnsupported'));
      return sidebar.openResource(address, { replaceTab: tab.id });
    };
    const dropTarget = () => controller.set({ pendingOpen: undefined });
    const proceedOpen = async () => {
      const pending = controller.state.pendingOpen;
      if (!pending || pending.tabId !== tab.id || pending.source !== controller.address) return;
      if (controller.lease && controller.state.mode === 'edit') {
        // Focus leaves the editor for the save (the tab covers it meanwhile).
        triggers.more.current?.focus();
        if (controller.state.backend === 'kit' && !controller.state.authorized) await controller.flushEditor().catch(() => {});
      }
      const plan = switchPlan(controller.state, Boolean(controller.lease), controller.accountHold === true);
      if (plan === 'account') { dropTarget(); return void controller.set({ notice: t?.('accountBlocked') }); }
      // A prompt this raises (authorize, outside, overwrite, failure) carries the switch on when confirmed.
      if (plan !== 'switch') controller.set({ pendingOpen: { ...pending, awaiting: true } });
      if (plan === 'readOnly') return void controller.set({ saveState: 'failed', saveFailed: t?.('readOnlyHint'), saveFailedSeen: undefined });
      if (plan === 'authorize') return void controller.set({ authorizePrompt: true });
      if (plan === 'save' && !await saveFromToolbar(false, true)) return;
      dropTarget();
      try { await replaceWith(pending.address); }
      catch (error) { controller.set({ notice: `${t?.('openFailed') ?? ''}${error.message}` }); }
    };
    const openPicked = async path => {
      try {
        const target = await controller.probe(path);
        // The same file under another spelling (absolute vs workspace path, a link): nothing to do.
        if (target.realPath && target.realPath === await controller.ownRealPath()) return;
        // Open in another tab (any spelling of its path): that tab comes forward; this one stays as it is.
        for (const other of controllers.values()) {
          if (other === controller || !target.realPath || await other.ownRealPath?.() !== target.realPath) continue;
          return void sidebarOf()?.openResource?.(other.address, { revealIfOpened: true });
        }
        controller.set({ notice: undefined, pendingOpen: { address: fileAddress(sessionId, target.path, controller.platform), source: controller.address, tabId: tab.id } });
        await proceedOpen();
      } catch (error) { controller.set({ notice: `${t?.('openFailed') ?? ''}${error.message}` }); }
    };
    // The desktop's own file chooser when DSH exposes the chosen file's path; else the workspace browser.
    const pathFor = (() => { try { const bridge = globalThis.__DSH_HOST_PATHS__; return typeof bridge?.pathFor === 'function' ? file => bridge.pathFor(file) : undefined; } catch { return undefined; } })();
    const openFile = () => {
      if (!pathFor || !doc?.body) return setOpen(open === 'browse' ? null : 'browse');
      const input = doc.createElement('input');
      Object.assign(input, { type: 'file', accept: '.docx,.xlsx,.pptx' });
      input.style.display = 'none';
      input.addEventListener('change', () => {
        const file = input.files?.[0];
        input.remove();
        if (!file) return;
        let path;
        try { path = pickedPath(pathFor(file), controller.platform); } catch {}
        if (path) void openPicked(path);
        else controller.set({ notice: t?.('openUnsupported') });
      }, { once: true });
      input.addEventListener('cancel', () => input.remove(), { once: true });
      doc.body.appendChild(input);
      input.click();
    };
    // Save As done: the copy holds the edits, so the original keeps none of them; the tab moves to the copy, or to the
    // file OPEN-01 was going to.
    const afterSaveAs = async path => {
      const pending = controller.state.pendingOpen;
      const copy = await controller.probe(path).catch(() => ({ path }));
      // From here the original is left as it is on disk: autosave stays held (the panel does not resume it) and the
      // tab's release does not autosave. Kit's record of the edits is reset from disk; if Kit refuses, the edits stay
      // only in that record (the original file is still not written) and the tab says so.
      controller.pauseAutosave(true);
      controller.abandoning = true;
      if (controller.lease && controller.state.backend === 'kit' && (controller.state.dirty || controller.state.unstaged)) {
        try { await controller.call('resolveDiskChange', { id: controller.lease.id, useDisk: true }); }
        catch (error) { controller.set({ notice: `${t?.('saveAsKept') ?? ''}${error.message}` }); }
      }
      setOpen(null);
      controller.set({ pendingOpen: undefined, saveState: undefined, saveFailed: undefined, saveFailedSeen: undefined });
      try { await replaceWith(pending?.address ?? fileAddress(sessionId, copy.path, controller.platform)); }
      catch (error) {
        // The tab stays on the original after all: its edits are the user's again.
        controller.abandoning = false;
        controller.pauseAutosave(false);
        controller.set({ notice: `${t?.('openFailed') ?? ''}${error.message}` });
      }
    };
    // Cmd/Ctrl+S in the editor: not while a switch, reload or other save is under way.
    controller.onSaveRequest = () => {
      const now = controller.state;
      if (controller.switching || controller.accountHold || toolbarBusy(false, now)) return;
      if ((now.dirty || now.unstaged) && (now.backend !== 'kit' || now.authorized)) void saveFromToolbar();
    };
    const step = action => controller.editHistory(action).catch(error => controller.set({ notice: error.message }));
    const print = () => controller.print().catch(error => controller.set({ notice: error.message }));
    const status = (className, text, title, id) => h('span', { role: 'status', id, className: `sbp-text ${className}`, title: title ?? text }, text);
    const failedId = `sbp-save-failed-${tab.id}`;
    const reasonTitle = reason => reason === 'account' ? t?.('accountBlocked') : reason === 'busy' ? t?.('busyTitle') : reason === 'view' ? t?.('saveViewTitle') : reason === 'readOnly' ? t?.('readOnlyTitle') : undefined;
    const tool = (key, label, { disabled, reason, onClick, look, describedBy } = {}, ...content) => h('button', { type: 'button', className: 'sbp-icon', ...stopProps(key),
      'aria-label': label, title: reasonTitle(reason) ?? label, 'aria-disabled': disabled || undefined, 'aria-busy': look === 'saving' || undefined, 'data-look': look, 'aria-describedby': describedBy,
      onClick: () => { if (!disabled) onClick(); } }, ...content);
    const saveLook = edits.save?.look;
    const saveButton = () => tool('save', saveLook === 'saved' ? t?.('savedLabel') : saveLook === 'saving' ? t?.('saving') : saveLook === 'dirty' ? `${t?.('module_save')} · ${t?.('unsavedBadge')}` : t?.('module_save'),
      { ...edits.save, onClick: () => saveFromToolbar(), describedBy: saveLook === 'failed' ? failedId : undefined },
      saveLook === 'saving' ? h('span', { className: 'sbp-spin', 'aria-hidden': true }) : glyph(saveLook === 'saved' ? CHECK : SAVE),
      saveLook === 'dirty' || saveLook === 'failed' ? h('span', { className: 'sbp-dot', 'data-error': saveLook === 'failed' || undefined, 'aria-hidden': true }) : null);
    const updateLabel = entry === 'updating' ? t?.('updatingLabel') : entry === 'ready' ? t?.('updateReadyLabel') : t?.('updateAvailable')?.replace('{version}', update?.latest);
    // A busy tab holds back a new install; an install in progress or done can always be looked at.
    const updateBlocked = pBusy && entry === 'available';
    const raised = Boolean((!held && kit && editing && !state.authorized && state.dirty && state.handEdited && !hintClosed) || state.authorizePrompt
      || (state.saveState === 'failed' && !state.saveFailedSeen) || state.saveOutside || state.overwritePrompt || state.diskChanged);
    // One row: the status area shrinks and ellipsizes; the toolbar never wraps or squeezes.
    return h('div', { ref: root, className: 'sbp', 'data-shimo-actions': true },
      h('div', { className: 'sbp-status' },
        held ? status('sbp-warn', t?.('accountBlocked')) : null,
        held ? h('button', { type: 'button', className: 'sbp-btn sbp-primary', onClick: () => account.connect() }, t?.(accountNow.account.status === 'expired' ? 'accountReconnect' : 'accountConnect')) : null,
        !held && kit && editing ? (state.authorized
          ? status('sbp-ok', t?.('authorized'), t?.('authorizedTitle'))
          : h('button', { type: 'button', className: state.access === 'denied' ? 'sbp-btn' : 'sbp-btn sbp-primary', disabled: busy || state.access === 'denied',
            title: t?.(state.access === 'denied' ? 'readOnlyTitle' : 'authorizeTitle'), onClick: authorize }, t?.('authorize'))) : null,
        kit && editing && state.autosaveFailed ? status('sbp-error', t?.('autosaveFailed'), state.autosaveFailed) : null,
        kit && editing && state.dirty && state.unstaged ? status('sbp-error', t?.('unstagedBadge'), state.unstaged) : null,
        !kit && editing && state.dirty ? status(state.unstaged ? 'sbp-error' : 'sbp-warn', t?.(state.unstaged ? 'unstagedBadge' : state.saveNeedsApproval ? 'unsavedApprovalBadge' : 'unsavedBadge'), state.unstaged) : null,
        !held && !kit && editing && state.dirty && !shows('save') ? h('button', { type: 'button', className: 'sbp-btn', disabled: busy, onClick: () => save(false) }, busy ? t?.('saving') : t?.('save')) : null,
        state.saveState === 'failed' ? status('sbp-error', `${t?.('saveFailed') ?? ''}${state.saveFailed ?? ''}`, undefined, failedId) : null,
        state.notice ? status('sbp-error', state.notice) : null,
        state.fallback ? status('sbp-warn', t?.(state.fallback === 'unsaved' ? 'fallbackUnsaved' : 'fallbackPreview')) : null,
        modes?.edit && state.modeBlocked && state.dirty ? status('sbp-warn', t?.('modeBlocked')) : null,
        update?.installed ? status('sbp-ok', t?.('updateInstalled')?.replace('{version}', update.installed), t?.('updateInstalledTitle')) : null,
        update?.failed ? status('sbp-error', t?.('updateFailed'), update.failed) : null,
        // Save As replaces downloading a copy; Core tabs (no Save As) keep the download.
        caps.export && !edits.saveAs ? h('button', { type: 'button', className: 'sbp-btn', disabled: busy, onClick: () => run(() => controller.exportDownload()) }, t?.('export')) : null,
        // Kit keeps unsaved edits in its own record; drafts belong to the Core documents.
        kit ? null : h('button', { type: 'button', className: 'sbp-btn', onClick: async () => setDrafts(drafts ? undefined : await controller.listDrafts().catch(() => [])) }, t?.('drafts'))),
      h('div', { className: 'sbp-sep', 'aria-hidden': true }),
      h('div', { ref: bar, role: 'toolbar', 'aria-label': t?.('toolbar'), className: 'sbp-bar', onKeyDown: onBarKey },
        shows('undo') ? tool('undo', t?.('module_undo'), { ...edits.undo, onClick: () => step('undo') }, glyph(UNDO)) : null,
        shows('redo') ? tool('redo', t?.('module_redo'), { ...edits.redo, onClick: () => step('redo') }, glyph(REDO)) : null,
        shows('save') ? saveButton() : null,
        shows('saveAs') ? h('button', { type: 'button', ref: triggers.saveAs, className: 'sbp-icon', ...stopProps('saveAs'), 'aria-haspopup': 'dialog', 'aria-expanded': open === 'saveAs',
          'aria-label': t?.('module_saveAs'), title: reasonTitle(edits.saveAs.reason) ?? t?.('module_saveAs'), 'aria-disabled': edits.saveAs.disabled || undefined,
          onClick: () => { if (!edits.saveAs.disabled) setOpen(open === 'saveAs' ? null : 'saveAs'); } }, glyph(SAVE_AS)) : null,
        // Withdrawn for now (WITHDRAWN_MODULES in toolbar.mjs): a file opened here loses the chat's context, and opening
        // local files belongs in the 工作台 or a desktop client. shows('folder') is false; the open flow stays.
        shows('folder') ? h('button', { type: 'button', ref: triggers.folder, className: 'sbp-icon', ...stopProps('folder'), 'aria-haspopup': 'dialog', 'aria-expanded': open === 'browse',
          'aria-label': t?.('module_folder'), title: reasonTitle(edits.folder.reason) ?? t?.('module_folder'), 'aria-disabled': edits.folder.disabled || undefined,
          onClick: () => { if (!edits.folder.disabled) openFile(); } }, glyph(FOLDER)) : null,
        shows('print') ? tool('print', t?.('module_print'), { ...edits.print, onClick: print }, glyph(PRINT)) : null,
        fileKeys.length && (showMode || showUpdate) ? h('div', { className: 'sbp-sep', 'aria-hidden': true }) : null,
        showMode ? h('div', { role: 'group', 'aria-label': t?.('module_mode'), className: 'sbp-group' },
          showFocus ? h('button', { type: 'button', className: 'sbp-icon', ...stopProps('focus'), 'aria-pressed': Boolean(rightPane?.fullscreen),
            'aria-label': t?.(rightPane?.fullscreen ? 'exitFocus' : 'focusMode'), title: t?.(rightPane?.fullscreen ? 'exitFocus' : 'focusMode'), onClick: toggleFocus }, glyph(FOCUS)) : null,
          modeButton('edit', PEN, 'modeEdit'), modeButton('view', EYE, 'modeView')) : null,
        showUpdate ? h('button', { type: 'button', ref: triggers.update, className: 'sbp-icon', ...stopProps('update'), 'aria-haspopup': 'dialog', 'aria-expanded': open === 'update',
          'aria-label': updateLabel, title: updateBlocked ? t?.('busyTitle') : updateLabel, 'aria-disabled': updateBlocked || undefined, 'data-update': entry,
          // Opening the bubble on a new release asks again, so the version it shows is current. A second click closes it.
          onClick: () => { if (updateBlocked) return; if (open !== 'update' && entry === 'available') void updates.check(true); setOpen(open === 'update' ? null : 'update'); } },
          entry === 'updating' ? h('span', { className: 'sbp-spin', 'aria-hidden': true }) : glyph(DOWNLOAD),
          entry === 'available' || entry === 'ready' ? h('span', { className: 'sbp-badge', 'aria-hidden': true }, '!') : null) : null,
        fileKeys.length || showMode || showUpdate ? h('div', { className: 'sbp-sep', 'aria-hidden': true }) : null,
        h('button', { type: 'button', ref: triggers.more, className: 'sbp-icon', 'data-more': true, ...stopProps('more'), 'aria-haspopup': 'menu', 'aria-expanded': open === 'more',
          'aria-label': t?.('customize'), title: t?.('customize'),
          onClick: () => setOpen(open === 'more' ? null : 'more'),
          onKeyDown: event => { if (event.key === 'ArrowDown') { event.preventDefault(); setOpen('more'); } } }, glyph(MORE)),
        showAccount ? h('button', { type: 'button', ref: triggers.account, className: 'sbp-icon', ...stopProps('account'), 'aria-haspopup': 'menu', 'aria-expanded': open === 'account',
          'aria-label': t?.('module_account'), title: pBusy ? t?.('busyTitle') : t?.('module_account'), 'aria-disabled': pBusy || undefined,
          onClick: () => { if (!pBusy) setOpen(open === 'account' ? null : 'account'); },
          onKeyDown: event => { if (event.key === 'ArrowDown' && !pBusy) { event.preventDefault(); setOpen('account'); } } }, glyph(USER)) : null),
      open === 'more' ? h(ToolbarMenu, { t, anchor: triggers.more, root, modules: offered, hidden, close, preview: accountNow.preview, openSkills: controller ? () => setOpen('skills') : undefined }) : null,
      open === 'skills' && controller ? h(SkillsPanel, { t, controller, anchor: triggers.more, root, close }) : null,
      open === 'account' ? h(AccountMenu, { t, anchor: triggers.account, root, close, snapshot: accountNow }) : null,
      open === 'saveAs' ? h(SaveAsPanel, { t, controller, anchor: shows('saveAs') ? triggers.saveAs : root, root, close, done: afterSaveAs }) : null,
      open === 'browse' ? h(BrowsePanel, { t, controller, anchor: shows('folder') ? triggers.folder : root, root, close, choose: path => { setOpen(null); void openPicked(path); } }) : null,
      open === 'update' && showUpdate ? h(UpdatePanel, { update, entry, t, anchor: triggers.update, root, unsaved: anyUnsaved(controllers), close, install: installUpdate }) : null,
      // After an update: what changed, once. It waits while the user has a popover open or the tab asks something itself.
      notes && !open && !raised ? h(WhatsNewCard, { notes, t, root, shown: () => whatsNew.shown(), close: () => whatsNew.dismiss() }) : null,
      // Hand editing before authorization (not an editor's own load-time change): say once that nothing is saved yet.
      !held && kit && editing && !state.authorized && state.dirty && state.handEdited && !hintClosed
        ? prompt(t?.(state.access === 'denied' ? 'readOnlyHint' : 'authorizeHint'), () => setHintClosed(true),
          [choice(t?.('later'), () => setHintClosed(true)), state.access === 'denied' ? null : choice(t?.('authorize'), authorize, true, busy)]) : null,
      state.authorizePrompt ? confirm(t?.('saveOutside'), () => controller.set({ authorizePrompt: undefined, pendingOpen: undefined }),
        async () => { controller.set({ authorizePrompt: undefined }); await authorize(); if (controller.state.authorized && controller.state.pendingOpen?.awaiting) void proceedOpen(); }, t?.('saveConfirm')) : null,
      state.saveState === 'failed' && !state.saveFailedSeen
        ? prompt(`${t?.('saveFailed') ?? ''}${state.saveFailed ?? ''}`, () => controller.set({ saveFailedSeen: true, pendingOpen: undefined }),
          [choice(t?.('cancel'), () => controller.set({ saveFailedSeen: true, pendingOpen: undefined })),
            edits.saveAs && !edits.saveAs.disabled ? choice(t?.('saveAsAction'), () => { controller.set({ saveFailedSeen: true }); setOpen('saveAs'); }) : null,
            state.access === 'denied' ? null : choice(t?.('retry'), () => (state.pendingOpen ? proceedOpen() : saveFromToolbar()), true, pBusy)], true) : null,
      state.saveOutside ? confirm(t?.('saveOutside'), () => controller.set({ saveOutside: undefined, pendingOpen: undefined }),
        async () => { controller.set({ saveOutside: undefined }); if (await saveFromToolbar(true) && controller.state.pendingOpen?.awaiting) void proceedOpen(); }, t?.('saveConfirm')) : null,
      state.overwritePrompt ? confirm(t?.('overwritePrompt'), () => controller.set({ overwritePrompt: undefined, pendingOpen: undefined }),
        async () => { await save(state.overwritePrompt.approved, true); if (!controller.state.dirty && controller.state.pendingOpen?.awaiting) void proceedOpen(); }, t?.('overwriteConfirm')) : null,
      state.diskChanged ? confirm(t?.('diskChanged'), () => run(() => controller.resolveDiskChange(false)), () => run(() => controller.resolveDiskChange(true)), t?.('useDisk'), t?.('keepEdits')) : null,
      drafts ? h('div', { role: 'menu', className: 'sbp-light', style: menu }, drafts.length ? drafts.map(draft => h('div', { key: draft.id, style: { display: 'flex', gap: 6, alignItems: 'center' } },
        h('span', { style: { flex: 1, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' } }, `${draft.name} · ${new Date(draft.createdAt).toLocaleString()}`),
        h('button', { type: 'button', className: 'sbp-btn', onClick: () => controller.downloadDraft(draft.id) }, t?.('export')),
        h('button', { type: 'button', className: 'sbp-btn', onClick: async () => { await controller.discardDraft(draft.id); setDrafts(await controller.listDrafts()); } }, '×'))) : t?.('noDrafts')) : null);
  }

  /**
   * A white popover under its anchor (MENU-01), placed again whenever the window or the header changes size.
   * A user-opened one takes focus, closes on a click outside, a click into the editor frame (window blur) or focus
   * leaving the P area, and gives focus back to its trigger on Esc/⌘.. One the tab raises itself leaves focus where it
   * is; Tab reaches the dialog itself first, where Return runs its main action. Once focus is inside, Tab cycles and
   * Esc/⌘. cancels; however it closes, focus that was inside goes back to where it came from in the header.
   */
  function Popover({ kind, anchor, root, className, role, label, onClose, onCancel, onKey, initial, trap = true, holding, children }) {
    const box = useRef(null);
    const previous = useRef(null);
    const inside = useRef(false);
    const win = doc.defaultView ?? globalThis;
    useLayoutEffect(() => {
      const element = box.current;
      const trigger = anchor.current;
      if (!element || !trigger) return undefined;
      const apply = place => Object.assign(element.style, { left: `${place.left}px`, top: `${place.top}px`, maxWidth: `${place.maxWidth}px`, maxHeight: `${place.maxHeight}px` });
      const measure = () => {
        if (!anchor.current) return;
        Object.assign(element.style, { maxWidth: '', maxHeight: '' });
        const viewport = { width: doc.documentElement.clientWidth, height: doc.documentElement.clientHeight };
        const size = element.getBoundingClientRect();
        const place = placeMenu({ trigger: anchor.current.getBoundingClientRect(), width: size.width, height: size.height, viewport });
        apply(place);
        const landed = element.getBoundingClientRect();
        if (Math.abs(landed.left - place.left) > 0.5 || Math.abs(landed.top - place.top) > 0.5)
          apply(placeMenu({ trigger: anchor.current.getBoundingClientRect(), width: size.width, height: size.height, viewport, target: place, measured: landed }));
      };
      measure();
      if (kind === 'user') (initial?.(element) ?? element.querySelector('button:not(:disabled),[tabindex]'))?.focus();
      // Status changes and sidebar drags move the header without resizing the window.
      const observer = typeof ResizeObserver === 'function' && root.current ? new ResizeObserver(measure) : undefined;
      observer?.observe(root.current);
      win.addEventListener('resize', measure);
      return () => { observer?.disconnect(); win.removeEventListener('resize', measure); };
    }, []);
    useEffect(() => {
      const returnFocus = () => {
        const back = previous.current;
        if (inside.current && (!doc.activeElement || doc.activeElement === doc.body) && back && root.current?.contains(back)) back.focus();
      };
      if (kind !== 'user') return returnFocus;
      // A native picker the popover opened blurs the window: the popover stays for its answer.
      const outside = event => { if (!holding?.current && !box.current?.contains(event.target) && !anchor.current?.contains(event.target)) onClose(false); };
      const away = () => { if (!holding?.current) onClose(false); };
      doc.addEventListener('pointerdown', outside, true);
      win.addEventListener('blur', away);
      return () => { doc.removeEventListener('pointerdown', outside, true); win.removeEventListener('blur', away); };
    }, []);
    const onKeyDown = event => {
      if (onKey?.(event)) return;
      if (cancelKey(event)) {
        event.preventDefault();
        event.stopPropagation();
        if (kind === 'user') return onClose(true);
        return onCancel?.();
      }
      if (event.key === 'Enter' && event.target === box.current) {
        event.preventDefault();
        box.current.querySelector('.sbp-primary:not(:disabled)')?.click();
        return;
      }
      if (event.key !== 'Tab' || !trap || !inside.current) return;
      const items = [...box.current.querySelectorAll('button:not(:disabled),a[href],input,select,textarea')];
      if (!items.length) return;
      const at = items.indexOf(doc.activeElement);
      const next = event.shiftKey ? (at <= 0 ? items.length - 1 : at - 1) : (at === items.length - 1 ? 0 : at + 1);
      event.preventDefault();
      items[next].focus();
    };
    const onFocus = event => {
      if (!box.current?.contains(event.relatedTarget)) previous.current = event.relatedTarget;
      inside.current = true;
    };
    const onBlur = event => {
      if (box.current?.contains(event.relatedTarget)) return;
      inside.current = false;
      // Focus moved somewhere else on the page, outside the P area: a user-opened popover closes without taking it back.
      if (kind === 'user' && !holding?.current && event.relatedTarget && !root.current?.contains(event.relatedTarget)) onClose(false);
    };
    return h('div', { ref: box, role, 'aria-label': label, 'aria-modal': false, tabIndex: kind === 'self' ? 0 : -1,
      className: `sbp-pop ${className}`, onKeyDown, onFocus, onBlur }, children);
  }

  /** The three-dot menu: one checkbox item per module this tab offers; choosing one closes the menu. */
  function ToolbarMenu({ t, anchor, root, modules, hidden, close, preview, openSkills }) {
    const items = useRef([]);
    const checked = id => (id === 'account' ? preview === true : !hidden.includes(id));
    const choose = id => { if (id === 'account') account.setPreview(!preview); else prefs.toggle(id); close(true); };
    const onKey = event => {
      const index = items.current.indexOf(doc.activeElement);
      const step = menuKey(event.key, Math.max(index, 0), modules.length);
      if (!step) return false;
      if (step.close === 'leave') { close(false); return true; }
      event.preventDefault();
      event.stopPropagation();
      if (step.toggle) { if (index >= 0) choose(modules[index]); return true; }
      if (step.close) { close(true); return true; }
      items.current[step.index]?.focus();
      return true;
    };
    return h(Popover, { kind: 'user', anchor, root, className: 'sbp-menu', role: 'menu', label: t?.('customize'), onClose: close, onKey, trap: false, initial: element => element.querySelector('[role=menuitemcheckbox]') },
      h('div', { className: 'sbp-menu-title', 'aria-hidden': true }, t?.('customize')),
      modules.map((id, index) => h('button', { key: id, type: 'button', role: 'menuitemcheckbox', 'aria-checked': checked(id), tabIndex: -1, className: 'sbp-item',
        ref: element => { items.current[index] = element; }, onClick: () => choose(id) },
        h('span', { className: 'sbp-check', 'aria-hidden': true }, '✓'), t?.(id === 'account' ? 'module_accountPreview' : `module_${id}`))),
      openSkills ? h('button', { type: 'button', role: 'menuitem', tabIndex: 0, className: 'sbp-item', 'data-skills': true, onClick: openSkills },
        h('span', { className: 'sbp-check', 'aria-hidden': true }), t?.('skillsMenu')) : null);
  }

  /**
   * Skills (spec 261110-universal-package §4.5): each skill's version, whether its resources are here, a switch that
   * hides it from the model, and Download / Clear cache; Check for updates reads the skill mirror's catalog (when the package names one). The Host does
   * the work; this polls its status while something is downloading.
   */
  function SkillsPanel({ t, controller, anchor, root, close }) {
    const [state, setState] = useState();
    const [busy, setBusy] = useState(false);
    const [problem, setProblem] = useState();
    const load = () => controller.call('skillsStatus', {}).then(setState, error => setProblem(error?.message));
    useEffect(() => { void load(); }, []);
    useEffect(() => {
      if (!state?.skills?.some(row => row.state === 'downloading')) return undefined;
      const timer = setTimeout(() => { void load(); }, 1000);
      return () => clearTimeout(timer);
    }, [state]);
    const act = (method, args) => { setBusy(true); setProblem(undefined); return controller.call(method, args).then(result => { setState(previous => ({ ...previous, ...result })); }, error => setProblem(error?.message)).finally(() => { setBusy(false); void load(); }); };
    const rows = state?.skills ?? [];
    const can = state?.canChange && state?.managed;
    return h(Popover, { kind: 'user', anchor, root, className: 'sbp-dialog sbp-skills', role: 'dialog', label: t?.('skillsTitle'), onClose: close, initial: element => element.querySelector('button') },
      h('div', { className: 'sbp-dialog-title' }, t?.('skillsTitle')),
      !state ? h('span', { className: 'sbp-quiet' }, t?.('loading') || 'Loading…') : null,
      ...rows.map(row => h('div', { key: row.name, className: 'sbp-row', 'data-skill': row.name },
        h('label', null, h('input', { type: 'checkbox', checked: row.enabled, disabled: busy || !state?.canChange, onChange: event => act('skillSetEnabled', { name: row.name, enabled: event.target.checked }) }), ` ${row.name}`),
        h('span', { className: 'sbp-quiet', title: row.error }, `v${row.version ?? '-'}${row.latest && row.latest !== row.version ? ` → v${row.latest}` : ''} · ${t?.(`skillState_${row.state}`) ?? row.state}`),
        can ? h('span', null,
          h('button', { type: 'button', className: 'sbp-btn', disabled: busy || row.state === 'ready' || row.state === 'downloading', onClick: () => act('skillPrefetch', { name: row.name }) }, t?.('skillsDownload')),
          h('button', { type: 'button', className: 'sbp-btn', disabled: busy, onClick: () => act('skillClearCache', { name: row.name }) }, t?.('skillsClear'))) : null)),
      problem ? h('span', { role: 'alert', className: 'sbp-strong' }, problem) : null,
      h('div', { className: 'sbp-actions' },
        can ? h('button', { type: 'button', className: 'sbp-btn', disabled: busy, onClick: () => act('skillsCheckUpdates', {}) }, t?.('skillsCheck')) : null,
        can ? h('button', { type: 'button', className: 'sbp-btn', disabled: busy, onClick: () => act('skillPrefetch', { name: null }) }, t?.('skillsDownloadAll')) : null,
        h('button', { type: 'button', className: 'sbp-btn sbp-primary', onClick: () => close(true) }, t?.('skillsDone'))));
  }

  /**
   * Save As (docs/01 §5): name with its extension apart (only the name is selected), the current format only, and the
   * place — the file's folder, or one chosen with the desktop's folder picker. Autosave waits while it is open.
   * An existing name asks Cancel / Replace in place; a place outside the workspace says that saving allows writing.
   */
  function SaveAsPanel({ t, controller, anchor, root, close, done }) {
    const finished = useRef(false);
    const source = controller.source.path;
    // Paths in the Host's spelling: the picker and the tab's own path may be Windows paths there.
    const { platform } = controller;
    const fileName = baseName(source, platform);
    const ext = fileName.slice(fileName.lastIndexOf('.'));
    const suggested = copyName(fileName, t?.('copySuffix'));
    const [stem, setStem] = useState(suggested.slice(0, suggested.length - ext.length));
    const [folder, setFolder] = useState(parentOf(source, platform));
    const [where, setWhere] = useState();
    const [problem, setProblem] = useState();
    const [replace, setReplace] = useState(false);
    const [working, setWorking] = useState(false);
    const input = useRef(null);
    const chooseButton = useRef(null);
    const holding = useRef(false);
    const picker = (() => { try { const bridge = globalThis.__DSH_DIRECTORY_PICKER__; return typeof bridge?.pick === 'function' ? bridge : undefined; } catch { return undefined; } })();
    useEffect(() => {
      controller.pauseAutosave(true);
      return () => {
        if (finished.current) return;
        controller.pauseAutosave(false);
        // Leaving Save As unfinished also leaves the switch it was part of (OPEN-01 Cancel).
        if (controller.state.pendingOpen) controller.set({ pendingOpen: undefined });
      };
    }, []);
    useEffect(() => {
      let live = true;
      controller.probe(folder || '.', 'folder').then(result => { if (live) setWhere(result); }, () => { if (live) setWhere(undefined); });
      return () => { live = false; };
    }, [folder]);
    const choose = async () => {
      holding.current = true;
      try { const chosen = await picker.pick(); if (typeof chosen === 'string' && isAbsolute(chosen, platform)) setFolder(chosen); }
      catch {} finally { holding.current = false; chooseButton.current?.focus(); }
    };
    const submit = async (overwrite = false) => {
      const issue = nameProblem(stem);
      if (issue) { setProblem(issue); input.current?.focus(); return; }
      setWorking(true);
      setProblem(undefined);
      try {
        // Choosing the place here is the approval to write it (owner only, checked by the Host).
        const result = await controller.exportCopy(joinPath(folder, `${stem.trim()}${ext}`, platform), { overwrite, approved: true });
        finished.current = true;
        await done(result.path);
      } catch (error) {
        if (error?.code === 'officesdk/exists') setReplace(true);
        else { setReplace(false); setProblem(error?.code === 'officesdk/same-file' ? 'same' : error?.code === 'officesdk/target-busy' ? 'busy' : { message: error?.message }); }
      } finally { setWorking(false); }
    };
    const problemText = problem === 'empty' ? t?.('nameEmpty') : problem === 'characters' ? t?.('nameCharacters') : problem === 'dot' ? t?.('nameDot')
      : problem === 'same' ? t?.('nameSame') : problem === 'busy' ? t?.('targetBusy') : problem?.message ? `${t?.('saveAsFailed') ?? ''}${problem.message}` : undefined;
    const folderLabel = folder ? segments(folder, platform).at(-1) ?? '/' : t?.('workspace');
    return h(Popover, { kind: 'user', anchor, root, className: 'sbp-dialog sbp-save-as', role: 'dialog', label: t?.('saveAsTitle'), onClose: close, holding,
      initial: element => { const field = element.querySelector('input'); field?.select(); return field; } },
      h('label', { className: 'sbp-field' }, h('span', null, t?.('saveAsName')),
        h('input', { ref: input, value: stem, 'aria-invalid': problem && problem !== 'same' ? true : undefined, 'aria-describedby': problemText ? 'sbp-save-as-problem' : undefined,
          onChange: event => { setStem(event.target.value); setProblem(undefined); setReplace(false); },
          onKeyDown: event => { if (event.key === 'Enter') { event.preventDefault(); void submit(); } } })),
      h('div', { className: 'sbp-field' }, h('span', null, t?.('saveAsFormat')), h('span', { className: 'sbp-quiet' }, ext)),
      h('div', { className: 'sbp-field' }, h('span', null, t?.('saveAsWhere')),
        h('span', { className: 'sbp-where', title: folder || t?.('workspace') }, folderLabel),
        picker ? h('button', { type: 'button', ref: chooseButton, className: 'sbp-btn', onClick: choose }, t?.('saveAsChoose')) : null),
      where?.access === 'outside' ? h('span', { className: 'sbp-quiet' }, t?.('saveAsOutside')) : null,
      problemText ? h('span', { id: 'sbp-save-as-problem', role: 'alert', className: 'sbp-problem' }, problemText) : null,
      replace ? h('div', { role: 'alert', className: 'sbp-strong' }, t?.('replacePrompt')?.replace('{name}', `${stem.trim()}${ext}`)) : null,
      h('div', { className: 'sbp-actions' },
        h('button', { type: 'button', className: 'sbp-btn', onClick: () => (replace ? setReplace(false) : close(true)) }, t?.('cancel')),
        h('button', { type: 'button', className: 'sbp-btn sbp-primary', disabled: working, onClick: () => submit(replace) }, t?.(replace ? 'replace' : 'saveAsSave'))));
  }

  /** Open without the desktop's file chooser: the workspace's folders and Office files; click to select, double-click to open. */
  function BrowsePanel({ t, controller, anchor, root, close, choose }) {
    const source = controller.source.path;
    const [folder, setFolder] = useState(isAbsolute(source, controller.platform) ? '' : parentOf(source, controller.platform));
    const [listing, setListing] = useState();
    const [selected, setSelected] = useState();
    const [error, setError] = useState();
    useEffect(() => {
      let live = true;
      setSelected(undefined);
      controller.browse(folder).then(result => { if (live) { setListing(result); setError(undefined); } }, failure => { if (live) setError(failure?.message); });
      return () => { live = false; };
    }, [folder]);
    const join = name => (listing?.path ? `${listing.path}/${name}` : name);
    // The Host lists folders with '/' on every platform.
    const up = listing?.path ? parentOf(listing.path) : undefined;
    return h(Popover, { kind: 'user', anchor, root, className: 'sbp-dialog sbp-browse', role: 'dialog', label: t?.('openTitle'), onClose: close },
      h('div', { className: 'sbp-dialog-title' }, listing?.path || t?.('workspace')),
      error ? h('span', { role: 'alert', className: 'sbp-problem' }, error) : null,
      h('div', { role: 'list', 'aria-label': t?.('openTitle'), className: 'sbp-list' },
        listing && !listing.entries.length ? h('span', { className: 'sbp-quiet' }, t?.('emptyFolder')) : null,
        (listing?.entries ?? []).map(entry => h('button', { key: `${entry.kind}:${entry.name}`, type: 'button', role: 'listitem', className: 'sbp-item',
          'aria-pressed': entry.kind === 'file' ? selected === entry.name : undefined,
          onClick: () => (entry.kind === 'folder' ? setFolder(join(entry.name)) : setSelected(entry.name)),
          onDoubleClick: () => { if (entry.kind === 'file') choose(join(entry.name)); } },
          h('span', { className: 'sbp-check', 'aria-hidden': true }, entry.kind === 'folder' ? '▸' : selected === entry.name ? '✓' : ''), entry.name))),
      h('div', { className: 'sbp-actions' },
        h('button', { type: 'button', className: 'sbp-btn', disabled: up === undefined, onClick: () => setFolder(up) }, t?.('folderUp')),
        h('button', { type: 'button', className: 'sbp-btn', onClick: () => close(true) }, t?.('cancel')),
        h('button', { type: 'button', className: 'sbp-btn sbp-primary', disabled: !selected, onClick: () => choose(join(selected)) }, t?.('openAction'))));
  }

  /**
   * 个人中心 (ACCOUNT-01): the account and its email, usage in credits, and the website. Every row is a menu item so
   * the email and credits can be reached with the arrow keys; rows without an action are disabled. Not connected:
   * the first row connects; usage shows no value. Usage unknown: that row retries.
   */
  function AccountMenu({ t, anchor, root, close, snapshot }) {
    const items = useRef([]);
    const { status, email, credits } = snapshot.account;
    const connected = status === 'connected';
    const website = () => {
      // An anchor in the main document, like the downloads: opens a new browser tab and leaves this document as it is.
      openOfficeWebsite(doc);
      close(true);
    };
    const rows = [
      { key: 'account', icon: USER, label: connected ? t?.('accountConnected') : t?.(status === 'expired' ? 'accountReconnect' : 'accountConnect'), detail: email, tall: true,
        act: connected ? undefined : () => { account.connect(); close(true); } },
      { key: 'usage', icon: USAGE, label: t?.('accountUsage'),
        value: !connected ? undefined : credits === null ? t?.('accountUsageUnavailable') : t?.('accountCredits')?.replace('{n}', String(credits)),
        act: connected && credits === null ? () => account.refresh() : undefined },
      { key: 'website', icon: GLOBE, label: t?.('accountWebsite'), link: true, act: website },
      // Demo states: the preview's data is made up, so the menu says so and lets the user try each state.
      ...DEMO_STATES.map(state => ({ key: `demo-${state.id}`, label: t?.(`demo_${state.id}`), radio: true, checked: demoState(snapshot.account) === state.id,
        act: () => { account.setMock(state.account); close(true); } })),
    ];
    const renderRow = row => { const index = rows.indexOf(row); return h('button', { key: row.key, type: 'button', role: row.radio ? 'menuitemradio' : 'menuitem', 'aria-checked': row.radio ? row.checked : undefined, tabIndex: -1,
        className: `sbp-item sbp-account-row${row.tall ? ' sbp-account-tall' : ''}`,
        'aria-disabled': row.act ? undefined : true, ref: element => { items.current[index] = element; }, title: row.detail,
        onClick: () => row.act?.() },
        row.radio ? h('span', { className: 'sbp-check', 'aria-hidden': true }, row.checked ? '✓' : '') : glyph(row.icon),
        h('span', { className: 'sbp-account-text' }, h('span', null, row.label), row.detail ? h('span', { className: 'sbp-account-detail' }, row.detail) : null),
        row.value ? h('span', { className: 'sbp-account-value' }, row.value) : null,
        row.link ? h('span', { className: 'sbp-account-link', 'aria-hidden': true }, glyph(LINK_OUT)) : null); };
    const onKey = event => {
      const index = items.current.indexOf(doc.activeElement);
      const step = accountMenuKey(event.key, Math.max(index, 0), rows.length);
      if (!step) return false;
      if (step.close === 'leave') { close(false); return true; }
      event.preventDefault();
      event.stopPropagation();
      if (step.activate) { rows[index]?.act?.(); return true; }
      if (step.close) { close(true); return true; }
      items.current[step.index]?.focus();
      return true;
    };
    return h(Popover, { kind: 'user', anchor, root, className: 'sbp-menu sbp-account', role: 'menu', label: `${t?.('module_account')} · ${t?.('accountDemo')}`, onClose: close, onKey, trap: false,
      initial: element => element.querySelector('[role=menuitem]') },
      h('div', { className: 'sbp-menu-title sbp-account-demo', 'aria-hidden': true }, t?.('accountDemo')),
      ...rows.filter(row => !row.radio).map(row => renderRow(row)),
      h('div', { role: 'group', 'aria-label': t?.('accountDemoStates'), key: 'demo' },
        h('div', { className: 'sbp-menu-title sbp-account-section', 'aria-hidden': true }, t?.('accountDemoStates')),
        rows.filter(row => row.radio).map(row => renderRow(row))));
  }

  /**
   * The update bubble (design package V3.0 P区), on the user's click only: the release's version and date on the left,
   * the black Update button top right, a thin rule, then what changed. Update is off while any tab has edits to write;
   * while installing it shows the progress, then that the update is ready. Values all come from the Host's check.
   */
  function UpdatePanel({ update, entry, t, anchor, root, unsaved, close, install }) {
    const view = updatePanelView(update, entry, unsaved);
    const body = useRef(null);
    // A disabled button cannot hold focus: the bubble itself takes it, so Esc still closes it (never the entry's focus).
    useEffect(() => {
      const box = body.current?.closest('.sbp-pop');
      const active = doc.activeElement;
      if (box && (!active || active === doc.body || (box.contains(active) && active.disabled))) box.focus({ preventScroll: true });
    }, [view.action.disabled, entry]);
    const lang = t?.('lang') === 'en' ? 'en-US' : 'zh-CN';
    const date = view.date ? new Intl.DateTimeFormat(lang, { year: 'numeric', month: 'long', day: 'numeric' }).format(new Date(view.date)) : undefined;
    const act = view.action.install ? install : view.action.close ? () => close(true) : undefined;
    return h(Popover, { kind: 'user', anchor, root, className: 'sbp-update sbp-update-opening', role: 'dialog', label: t?.('updateTitle'), onClose: close,
      initial: element => element.querySelector('.sbp-update-primary:not(:disabled)') ?? element },
      h('div', { ref: body, className: 'sbp-update-head' },
        h('div', { className: 'sbp-update-meta' },
          h('div', { className: 'sbp-update-version' }, fill(t?.('updateProduct'), { version: view.version ?? '' })),
          date ? h('time', { className: 'sbp-update-date', dateTime: view.date }, date) : null),
        h('button', { type: 'button', className: 'sbp-update-primary sbp-primary', disabled: view.action.disabled, 'aria-busy': view.action.busy || undefined, onClick: act },
          view.action.busy ? h('span', { className: 'sbp-spin', 'aria-hidden': true }) : null, t?.(view.action.label))),
      h('section', { className: 'sbp-update-notes' },
        h('h3', null, t?.('updateNotes')),
        view.notes.length ? h('ul', null, view.notes.map((text, index) => h('li', { key: index }, text))) : h('p', { className: 'sbp-update-fallback' }, t?.('updateNotesFallback'))),
      view.message ? h('p', { className: `sbp-update-message sbp-update-${view.message.tone}`, role: view.message.tone === 'warning' ? 'alert' : 'status' }, view.message.key ? t?.(view.message.key) : view.message.text) : null);
  }
  /** 「本次更新」: the points of each version since the last one shown, raised by the tab without taking focus. */
  function WhatsNewCard({ notes, t, root, shown, close }) {
    const [expanded, setExpanded] = useState(false);
    // The card can mount in the same render as the header it hangs from, before that header's ref is attached: the
    // Popover places itself from that ref, so it waits one effect for it.
    const [placed, setPlaced] = useState(false);
    useEffect(() => { setPlaced(true); }, []);
    useEffect(() => (placed ? shown() : undefined), [placed]);
    if (!placed) return null;
    const view = whatsNewView(notes, expanded, undefined, t?.('lang'));
    const title = fill(t?.(view.title), view.values);
    return h(Popover, { kind: 'self', anchor: root, root, className: 'sbp-dialog sbp-whats-new', role: 'dialog', label: title, onCancel: close },
      h('div', { className: 'sbp-dialog-title' }, h('span', null, title), h('span', { className: 'sbp-tag' }, t?.('whatsNewTag'))),
      h('div', { className: 'sbp-notes', 'data-shimo-whats-new': true }, view.groups.map(group => h('div', { key: group.version, className: 'sbp-notes-group' },
        view.grouped ? h('div', { className: 'sbp-quiet' }, `v${group.version}`) : null,
        h('ul', null, group.highlights.map((text, index) => h('li', { key: index }, text)))))),
      h('div', { className: 'sbp-actions' },
        view.hidden > 0 || expanded ? h('button', { type: 'button', className: 'sbp-btn', onClick: () => setExpanded(!expanded) },
          expanded ? t?.('whatsNewLess') : fill(t?.('whatsNewMore'), { n: view.total })) : null,
        h('button', { type: 'button', className: 'sbp-btn sbp-primary', onClick: close }, t?.('whatsNewOk'))));
  }
  const glyph = d => h('svg', { width: 16, height: 16, viewBox: '0 0 24 24', fill: 'none', stroke: 'currentColor', strokeWidth: 1.8, strokeLinecap: 'round', strokeLinejoin: 'round', 'aria-hidden': true }, h('path', { d }));

  ctx.effect(() => ctx.locale.register(NS, MESSAGES));
  ctx.effect(() => mountStyles(doc));
  ctx.effect(() => account.subscribe(holdAll));
  ctx.effect(() => {
    const win = doc?.defaultView;
    const changed = event => { if (account.keys.includes(event.key)) account.reload(); };
    win?.addEventListener('storage', changed);
    return () => win?.removeEventListener('storage', changed);
  });
  ctx.effect(() => {
    const win = doc?.defaultView;
    const changed = event => { if (event.key === prefs.key) prefs.reload(); };
    win?.addEventListener('storage', changed);
    return () => win?.removeEventListener('storage', changed);
  });
  // AI commands need the on-screen Session and the right Sidebar; without them documents only preview.
  // The Host's platform as any open tab's lease said it, else as the last AI command said it; before either, POSIX (see fileAddress).
  let bridge;
  const hostPlatform = () => [...controllers.values()].find(controller => controller.platform)?.platform ?? bridge?.lastPlatform;
  ctx.effect(() => {
    const sidebar = environment.sidebarRight ?? ctx.get?.('sidebarRight');
    if (!sidebar?.mounted || !sidebar.openResource) return () => {};
    const previous = bridge?.lastPlatform;
    bridge = new AgentBridge({ rpc: ctx.connection.rpc, openResource: (address,options) => sidebar.openResource(address,options), controllers: registry, mounted: sidebar.mounted, views: toolViews, platform: hostPlatform });
    bridge.lastPlatform = previous;
    return bridge.start();
  });
  // Chat rows for the presentation tools. Opening a file uses the AI tab's own address, and only for the Session the
  // right Sidebar (and the bridge) follows.
  registerDeckChat(ctx, React);
  registerToolViews(ctx, React, { store: toolViews, deck: deckOpener(sidebarOf, hostPlatform) });
  const stylePicker = registerStylePicker(ctx, React, { rpc: ctx.connection?.rpc });
  registerDocumentContext(ctx, React, { sidebarOf, controllers, composers, document: doc });
  // The model's style recommendation, above the input box.
  registerStyleFlows(ctx, React, { stores: stylePicker.stores, document: doc });
  // Selection rewrite: each input box registers itself, and every open tab of its Session hears when it can send.
  registerSelectionComposer(ctx, React, { registry: composers });
  ctx.effect(() => composers.subscribe(sessionId => pushComposerState(controllers.values(), composers, sessionId)));
  for (const extension of EXTENSIONS) {
    const id = `officesdk/${extension}`;
    ctx.effect(() => ctx.documentPreviews.register({ id, extensions: [extension], binaryExtensions: [extension], priority: 'extension', title: () => ctx.locale.bind(NS)('title'), loading: 'renderer', wrap: false }));
    ctx.effect(() => ctx.slots.inject('sidebar.right.tab.document', () => ctx.slots.register({ name: 'sidebar.right.tab.document', key: id, locale: NS }, Document)));
    ctx.effect(() => ctx.slots.inject('sidebar.right.tab.document.action', () => ctx.slots.register({ name: 'sidebar.right.tab.document.action', key: id, locale: NS }, Action)));
  }
}

const cover = { position: 'absolute', inset: 0, padding: 16, fontSize: 13, background: 'var(--dsw-alias-bg-layer-1,#fff)', color: 'var(--dsw-alias-label-primary,#333)' };
const banner = { position: 'absolute', left: 8, right: 8, top: 8, display: 'flex', gap: 8, alignItems: 'center', padding: '6px 10px', fontSize: 12, background: 'var(--dsw-alias-bg-layer-1,#fff)', border: '1px solid var(--dsw-alias-border,#e5e7eb)', borderRadius: 4 };
// The preview panel around the tab sets a monospace font; the AI's notes in the tab use the interface font.
const uiFont = '-apple-system, system-ui, "Segoe UI", "PingFang SC", "Microsoft YaHei", sans-serif';
const button = { height: 28, padding: '0 10px', fontSize: 12, border: '1px solid var(--dsw-alias-border,#d1d5db)', borderRadius: 4, background: 'transparent', cursor: 'pointer' };
// Line icons (24×24 paths) from the v2.9 design package.
const DOWNLOAD = 'M4 15v6h16v-6M12 3v12M7 10l5 5 5-5';
const FOCUS = 'M13 4h7v7M4 13v7h7';
const PEN = 'm4 16 12-12 4 4L8 20H4zM14 6l4 4';
const EYE = 'M2 12s4-7 10-7 10 7 10 7-4 7-10 7-10-7-10-7zM12 9a3 3 0 1 0 0 6 3 3 0 0 0 0-6';
const MORE = 'M5 12h.1M12 12h.1M19 12h.1';
const UNDO = 'M8 5 3 10l5 5M3 10h10a7 7 0 0 1 7 7';
const REDO = 'm16 5 5 5-5 5M21 10H11a7 7 0 0 0-7 7';
const SAVE = 'M4 3h13l4 4v14H3V3zM7 3v6h9V3M7 21v-8h10v8M13 5v2';
const PRINT = 'M6 8V3h12v5M6 18H3v-8h18v8h-3M6 14h12v8H6zM17 11h1';
const SAVE_AS = 'M11 21H3V3h14l4 4v4M7 3v6h9V3M7 21v-8h5m1 6 7-7 3 3-7 7h-3z';
const FOLDER = 'M2 6h8l2 3h9v3M2 6v15h17l3-9H6l-4 9';
const USER = 'M12 3a4 4 0 1 0 0 8 4 4 0 0 0 0-8M4 21v-3a8 6 0 0 1 16 0v3';
const USAGE = 'M4 18a9 9 0 1 1 16 0M12 12l6-6M10 15a2 2 0 1 0 4 0 2 2 0 0 0-4 0M5 10l1 1M12 5v2M19 10l-1 1';
const GLOBE = 'M12 3a9 9 0 1 0 0 18 9 9 0 0 0 0-18M3 12h18M12 3c-5 5-5 13 0 18M12 3c5 5 5 13 0 18';
const LINK_OUT = 'M7 17 17 7M7 7h10v10';
const CHECK = 'm5 12 5 5 9-10';
const card = { position: 'absolute', right: 12, bottom: 12, zIndex: 20, width: 'min(420px, calc(100% - 24px))', display: 'grid', gap: 8, padding: 12, fontSize: 12, fontFamily: uiFont,
  background: 'var(--dsw-alias-bg-layer-1,#fff)', color: 'var(--dsw-alias-label-primary,#222)', border: '1px solid var(--dsw-alias-border,#e5e7eb)', borderRadius: 8, boxShadow: '0 8px 24px rgba(0,0,0,.14)' };
const rowStyle = { display: 'grid', gap: 2, padding: '4px 6px', borderRadius: 4, background: 'var(--dsw-alias-bg-layer-2,#f6f7f9)' };
// Core drafts list (not in the Kit package): white like every popover.
const menu = { position: 'absolute', right: 8, top: 32, zIndex: 10, minWidth: 260, maxWidth: 360, padding: 8, whiteSpace: 'normal', display: 'grid', gap: 6, fontSize: 12, background: '#fff', color: '#242424', border: '1px solid #c9c9cc', borderRadius: 10 };
