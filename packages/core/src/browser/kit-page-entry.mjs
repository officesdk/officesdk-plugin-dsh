import { createKitDocumentAgent } from './document-agent.mjs';
import { createKitPresentationAgent } from './kit-agent.mjs';
import { createKitSheetAgent } from './sheet-agent.mjs';
import { startKitPage } from './kit-page.mjs';

// The editor bundle is a separate, pinned component: it is loaded by URL, never bundled into this page.
const loadEditor = url => import(/* @vite-ignore */ url);
const animate = () => window.matchMedia?.('(prefers-reduced-motion: reduce)').matches !== true;
// AI editing needs the suite's JS API, all public in office-kit's editor page since 0.3.0-rc.1: Presentation
// (getJsApi / editorInternals, K1), Writer (the Word API) or Sheet (writable api.sheet, K16).
// Other suites have no agent.
// `locale` is the page's interface language (kit-page.mjs pageLocale: the tab's `lang`, else <html lang>).
const createAgent = ({ editor, document: page, locale }) => {
  if (editor.suite === 'presentation' && typeof editor.api?.presentation?.getJsApi === 'function') return createKitPresentationAgent({ editor, document: page, animate, locale });
  if (editor.suite === 'document' && editor.api?.writer?.Word) return createKitDocumentAgent({ editor, document: page, animate, locale });
  // Sheet writes need the members office-kit's editor page makes writable in edit mode (office-kit K16).
  const sheet = editor.suite === 'sheet' ? editor.api?.sheet : undefined;
  if (sheet && typeof sheet.workbook?.addWorksheet === 'function' && typeof sheet.activeSheet?.getRange === 'function') return createKitSheetAgent({ editor, document: page, animate, locale });
  return undefined;
};
startKitPage({ window, document, loadEditor, createAgent }).catch(error => console.error('officesdk editor page failed', error));
