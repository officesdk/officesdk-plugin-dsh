import { extname } from 'node:path';
import { fail } from './errors.mjs';

/** Format differences live here and in the browser adapters only. */
export const FORMATS = Object.freeze({
  docx: Object.freeze({ snapshot: 'mow', mainPart: 'word/document.xml', parts: ['package'] }),
  xlsx: Object.freeze({ snapshot: 'modoc', mainPart: 'xl/workbook.xml', parts: ['content'] }),
  pptx: Object.freeze({ snapshot: 'mop', mainPart: 'ppt/presentation.xml', parts: ['snapshot', 'asset'] }),
});
export const OPERATIONS = Object.freeze(['import', 'preview', 'edit', 'save', 'export']);

export function formatOf(path) {
  const format = typeof path === 'string' ? extname(path).slice(1).toLowerCase() : '';
  if (!Object.hasOwn(FORMATS, format)) fail('unsupported-format', 'Supported formats are DOCX, XLSX and PPTX.');
  return format;
}

/** The converter declares what it can do; the host may switch off writes (for example a remote filesystem). */
export function capabilities(format, converterCapabilities = {}, hostCapabilities = {}) {
  if (!Object.hasOwn(FORMATS, format)) fail('unsupported-format', 'Supported formats are DOCX, XLSX and PPTX.');
  const supported = converterCapabilities[format] ?? {};
  const result = {};
  for (const operation of OPERATIONS) result[operation] = supported[operation] === true && hostCapabilities[operation] !== false;
  result.preview &&= result.import;
  result.edit &&= result.preview;
  result.save &&= result.edit;
  return result;
}

/** Content types follow the snapshot encoding, not the backend name. */
export function contentType(format, content) {
  if (format === 'docx') return 'application/json; charset=utf-8';
  if (format === 'xlsx') return 'text/plain; charset=utf-8';
  return content.endsWith('.json') ? 'application/json; charset=utf-8' : 'application/octet-stream';
}
