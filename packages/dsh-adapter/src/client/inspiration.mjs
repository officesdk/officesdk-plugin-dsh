/**
 * The inspiration library's fixed facts and pure rules (design handoff OfficeSDK-DSH-Inspiration-v2.0,
 * docs/03-最终交互规格.md): the three file kinds, the built-in Word and Excel templates (PPT ones are the shimo-pptx
 * style packs the Host lists), the mixed catalog with each template's pages, the featured / mine / scene filters, the
 * fixed prompt a use writes at the start of the draft, the masonry layout, and the checks a dropped or chosen file gets
 * before "my templates" says the upload service is not connected. No file is ever read here: only its name, size and count.
 */
import { TEMPLATE_COVERS } from './template-covers.mjs';

export const KINDS = Object.freeze(['word', 'excel', 'ppt']);
export const KIND_LABEL = Object.freeze({ word: 'Word', excel: 'Excel', ppt: 'PPT' });
/** A page's frame per kind: Word A4 portrait (210:297), Excel 4:3, PPT 16:9; height over width for the masonry. */
export const FRAME_RATIO = Object.freeze({ word: '210 / 297', excel: '4 / 3', ppt: '16 / 9' });
export const FRAME_HEIGHT = Object.freeze({ word: 297 / 210, excel: 3 / 4, ppt: 9 / 16 });
/**
 * Extensions the "my templates" entry accepts per kind: the handoff's proposal, not a statement of what any service
 * supports — there is no upload service yet.
 */
export const UPLOAD_EXTENSIONS = Object.freeze({
  word: Object.freeze(['.docx', '.doc', '.png', '.jpg', '.jpeg', '.webp']),
  excel: Object.freeze(['.xlsx', '.xls', '.png', '.jpg', '.jpeg', '.webp']),
  ppt: Object.freeze(['.pptx', '.ppt', '.png', '.jpg', '.jpeg', '.webp']),
});
/**
 * Built-in Word and Excel templates: the live templates the bundled skills ship (the Host's template-pick.mjs has the
 * same ids): Word gongwen, no Excel one. Their covers are first pages the real editor rendered (template-covers.mjs, made by
 * tooling/covers/render-template-covers.mjs); a template without one shows the missing-cover face.
 */
export const BUILTIN = Object.freeze({
  word: Object.freeze([
    { id: 'meeting-minutes', name: { zh: '会议纪要', en: 'Meeting minutes' }, category: { zh: '会议协作', en: 'Meetings' } },
    { id: 'gongwen', name: { zh: '公文版式', en: 'Official document' }, category: { zh: '公文通知', en: 'Official notices' } },
  ]),
  excel: Object.freeze([]),
});
/** The monochrome bulb with the thin bolt, path for path from icons/inspiration-bulb.svg (24×24, stroke currentColor 1.5). */
export const BULB_PATHS = Object.freeze(['M9 18c0-1.7-.7-2.5-2-3.8a7 7 0 1 1 10 0c-1.3 1.3-2 2.1-2 3.8H9ZM9 20h6M10.5 22h3', 'M12.7 6.3 10 10.3h3.4L11.3 14.5']);

/** The fixed prompt a use writes at the start of the draft, per kind and language (§7). */
export const PROMPTS = Object.freeze({
  zh: Object.freeze({ word: '使用此文档技能帮助我创建文档：', excel: '使用此表格技能帮助我创建文档：', ppt: '使用此幻灯片技能帮助我创建文档：' }),
  en: Object.freeze({ word: 'Use this document skill to help me create a document: ', excel: 'Use this spreadsheet skill to help me create a document: ', ppt: 'Use this slides skill to help me create a document: ' }),
});
/** The skill each kind's templates belong to: the data-layer association a use records (never shown as a box). */
export const SKILL_OF = Object.freeze({ word: 'shimo-docx', excel: 'shimo-xlsx', ppt: 'shimo-pptx' });

export const isKind = kind => KINDS.includes(kind);
export const promptOf = (kind, lang = 'zh') => (PROMPTS[lang] ?? PROMPTS.zh)[kind] ?? '';
const localized = (value, lang) => value?.[lang] || value?.zh || value?.en || '';

/** The draft without the fixed prompt a use wrote, only while the draft still starts with it word for word. */
export function stripPrompt(draft, prompt) {
  const text = typeof draft === 'string' ? draft : '';
  return prompt && text.startsWith(prompt) ? text.slice(prompt.length) : text;
}

/**
 * The draft after a use: the new prompt first, then what the user wrote (the previous use's prompt goes only while
 * the draft still starts with it word for word; text the user rewrote stays). No line break is inserted.
 */
export function composeDraft(draft, previous, next) {
  const rest = stripPrompt(draft, previous);
  return `${next}${rest.replace(/^\s+/, '')}`;
}

const scenariosOf = (card, lang) => {
  const list = card?.scenarios?.[lang]?.length ? card.scenarios[lang] : card?.scenarios?.zh?.length ? card.scenarios.zh : card?.scenarios?.en;
  return Array.isArray(list) ? [...new Set(list.filter(item => typeof item === 'string' && item))] : [];
};

/**
 * The cards of a kind, all shaped alike: `{ id, kind, title, summary, cover, pages, categories, builtin }`. PPT comes
 * from the Host's style packs (thumbnail then previews, their summary and scenarios), Word and Excel from BUILTIN
 * (their rendered first page, or none). `available` (listStyles `templates`: `{ word: [ids], excel: [ids] }`, from a Host
 * whose package lists its skills) keeps only the built-in templates this copy can create from: the public build has
 * Word gongwen and no Excel template (spec 261113-bundle-public-skills). Without it (older Hosts) every built-in shows.
 */
export function templatesOf(kind, styles, lang = 'zh', available) {
  if (kind === 'ppt') {
    return (Array.isArray(styles) ? styles : []).map(card => {
      const pages = [...new Set([card.thumb, ...(Array.isArray(card.previews) ? card.previews : [])].filter(page => typeof page === 'string' && page))];
      return { id: card.id, kind, title: localized(card.name, lang) || card.id, summary: localized(card.summary, lang), cover: pages[0] ?? null, pages, builtin: true, categories: scenariosOf(card, lang) };
    });
  }
  const allowed = Array.isArray(available?.[kind]) ? available[kind] : undefined;
  return (BUILTIN[kind] ?? []).filter(item => !allowed || allowed.includes(item.id)).map(item => {
    const cover = TEMPLATE_COVERS[kind]?.[item.id] ?? null;
    return { id: item.id, kind, title: localized(item.name, lang), summary: '', cover, pages: cover ? [cover] : [], builtin: true, categories: [localized(item.category, lang)] };
  });
}

/** Every kind mixed (§3): Word, Excel, then PPT, each in its catalog order. */
export const libraryTemplates = (styles, lang = 'zh', available) => KINDS.flatMap(kind => templatesOf(kind, styles, lang, available));

/** Scene filters: the most used categories first; ties keep first-use order. At most `max` (ten in the library). */
export function categoriesOf(templates, max = 10) {
  const counts = new Map();
  for (const item of templates ?? []) for (const name of item.categories ?? []) counts.set(name, (counts.get(name) ?? 0) + 1);
  return [...counts.entries()].sort((a, b) => b[1] - a[1]).slice(0, max).map(([name]) => name);
}

export const SCOPE_MINE = 'mine';
export const SCOPE_FEATURED = 'featured';
/** The cards a scope shows: featured (every built-in one), mine (personal ones; none until an upload service exists), or a scene. */
export function scopeTemplates(templates, scope) {
  const list = templates ?? [];
  if (scope === SCOPE_MINE) return list.filter(item => !item.builtin);
  if (!scope || scope === SCOPE_FEATURED) return list.filter(item => item.builtin);
  return list.filter(item => item.categories?.includes(scope));
}

export const CARD_WIDTH = 280;
/** Bottom row plus the card's own padding, in px, added to the cover height when the masonry balances columns. */
const CARD_CHROME = 48;
const CARD_PADDING = 16;
/** How many fixed-width columns fit: as many as the width holds, at least one, never more than the cards. */
export function columnCount(width, gap, count, cardWidth = CARD_WIDTH) {
  if (!count) return 0;
  return Math.min(count, Math.max(1, Math.floor((width + gap) / (cardWidth + gap))));
}
/** The masonry: each card in turn goes into the currently shortest column (ties: the leftmost). */
export function masonry(items, columns, gap = 20, cardWidth = CARD_WIDTH) {
  const lists = Array.from({ length: columns }, () => []);
  const heights = Array(columns).fill(0);
  for (const item of items ?? []) {
    if (!columns) break;
    const column = heights.indexOf(Math.min(...heights));
    lists[column].push(item);
    heights[column] += (cardWidth - CARD_PADDING) * (FRAME_HEIGHT[item.kind] ?? 1) + CARD_CHROME + gap;
  }
  return lists;
}

/** Hover carousel timing (§4): first turn after 1 s, then every 3 s over the first three pages, 650 ms fades. */
export const HOVER_DELAY_MS = 1000;
export const HOVER_EVERY_MS = 3000;
export const HOVER_PAGES = 3;

/**
 * The collapsed summary's line count (§5): the lines that fit between the summary's top and the preview's bottom,
 * keeping 24px for 「更多」 when the text overflows; at least one.
 */
export function summaryLines(available, lineHeight, overflow) {
  if (!(lineHeight > 0)) return 1;
  return Math.max(1, Math.floor((Math.max(0, available) - (overflow ? 24 : 0)) / lineHeight));
}

/** The detail's example image: the second page, or the first for a one-page template (§5). */
export const featurePage = pages => (pages?.length ? pages[pages.length > 1 ? 1 : 0] : undefined);

/** A move within `count` pages (lightbox and carousel): ←/→/Home/End, clamped; undefined for other keys. */
export function pageKey(key, index, count) {
  if (!count) return undefined;
  if (key === 'ArrowLeft') return Math.max(0, index - 1);
  if (key === 'ArrowRight') return Math.min(count - 1, index + 1);
  if (key === 'Home') return 0;
  if (key === 'End') return count - 1;
  return undefined;
}

/**
 * The kinds this copy offers (`available`: listStyles `templates`): a Word or Excel kind whose list the Host sent empty
 * has no skill behind it (the public build has no shimo-xlsx, spec 261113-bundle-public-skills), so its switch tab and
 * "my templates" entry are left out. PPT always stays; without a list (older Hosts) every kind does.
 */
export function kindsOf(available) {
  return KINDS.filter(kind => kind === 'ppt' || !Array.isArray(available?.[kind]) || available[kind].length > 0);
}

/** The mode switch's keys (§2): →/↓ next, ←/↑ previous (both wrap), Home first, End last, over `count` kinds. */
export function modeKey(key, index, count = KINDS.length) {
  if (key === 'ArrowRight' || key === 'ArrowDown') return (index + 1) % count;
  if (key === 'ArrowLeft' || key === 'ArrowUp') return (index + count - 1) % count;
  if (key === 'Home') return 0;
  if (key === 'End') return count - 1;
  return undefined;
}

/**
 * Check what was chosen or dropped for "my templates": one non-empty file with an extension of `kind`. Returns
 * `{ error: 'multiple' | 'empty' | 'type' | 'none' }` or `{ file: { name, size } }`. Never reads the contents.
 */
export function checkUpload(files, kind) {
  const list = files ? Array.from(files) : [];
  if (!list.length) return { error: 'none' };
  if (list.length > 1) return { error: 'multiple' };
  const [file] = list;
  const name = typeof file?.name === 'string' ? file.name : '';
  const dot = name.lastIndexOf('.');
  const extension = dot >= 0 ? name.slice(dot).toLowerCase() : '';
  if (!UPLOAD_EXTENSIONS[kind]?.includes(extension)) return { error: 'type' };
  if (!(file.size > 0)) return { error: 'empty' };
  return { file: { name, size: file.size } };
}
