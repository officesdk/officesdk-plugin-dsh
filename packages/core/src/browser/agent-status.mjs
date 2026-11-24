/**
 * The agent cursor's status text, made by the plugin from what a step does (spec 261105 §7.3, §7.6, decision 2a): the
 * model's tool parameters stay as they are. Texts follow the v1.1 handoff's jobs (写入标题, 撰写正文, 生成内容, …).
 */
export const STATUS_TEXTS = Object.freeze({
  zh: { selecting: '框选中', title: '写入标题', body: '撰写正文', content: '生成内容', points: '梳理要点', conclusion: '完善结论', shape: '绘制图形', image: '插入图片',
    page: '创建页面', format: '调整格式', delete: '删除中', fill: '填充数据', formula: '编写公式', sheet: '命名工作表', table: '创建表格', chart: '生成图表', edit: '修改中' },
  en: { selecting: 'Selecting', title: 'Writing title', body: 'Writing text', content: 'Writing content', points: 'Listing points', conclusion: 'Wrapping up', shape: 'Drawing shape', image: 'Placing image',
    page: 'Creating slide', format: 'Formatting', delete: 'Deleting', fill: 'Filling data', formula: 'Writing formulas', sheet: 'Naming sheet', table: 'Creating table', chart: 'Making chart', edit: 'Editing' },
});

/** "Agent · <status>" in the locale's words. */
export const agentLabel = (locale, key) => `Agent · ${(STATUS_TEXTS[locale] ?? STATUS_TEXTS.zh)[key] ?? key}`;

const TITLE_PLACEHOLDERS = new Set(['Title', 'CenterTitle']);

/**
 * PPT: `step` from validatePptxOps; `shape` the target's record from the deck (placeholder, size) when it exists;
 * `conclusion` true for the lowest non-title text this batch writes on its slide.
 */
export function pptStatus(step, { shape, conclusion = false } = {}) {
  switch (step.op) {
    case 'addSlide': return 'page';
    case 'deleteSlide': case 'deleteShape': return 'delete';
    case 'addShape': case 'setFill': case 'setGeometry': case 'setBackground': return step.op === 'addShape' && step.text ? (conclusion ? 'conclusion' : 'content') : 'shape';
    case 'addImage': return 'image';
    case 'addTable': return 'table';
    case 'addChart': return 'chart';
    case 'setFont': case 'setAlignment': return 'format';
    case 'setText': case 'addTextBox': {
      const size = step.font?.size ?? shape?.fontSize;
      // A title: a title placeholder (new slides know the kind), a shape named like one (inspectDeck reports names, not
      // placeholder kinds), or large text.
      if (TITLE_PLACEHOLDERS.has(shape?.placeholder) || (/title|标题/i.test(shape?.name ?? '') && !/sub\s*title|副标题/i.test(shape?.name ?? '')) || (typeof size === 'number' && size >= 28)) return 'title';
      return conclusion ? 'conclusion' : 'content';
    }
    default: return 'edit';
  }
}

/** Word: `step` from validateDocxOps. */
export function docxStatus(step) {
  switch (step.op) {
    case 'insertParagraph': return ['title', 'subtitle', 'heading1', 'heading2', 'heading3'].includes(step.style) ? 'title' : step.list ? 'points' : 'body';
    case 'setText': return 'body';
    case 'formatParagraph': return 'format';
    case 'deleteParagraph': return 'delete';
    case 'insertTable': case 'setCell': return 'table';
    default: return 'edit';
  }
}

/** Excel: `step` from validateXlsxOps. */
export function xlsxStatus(step) {
  switch (step.op) {
    case 'setValues': return step.values.some(row => row.some(value => typeof value === 'string' && value.startsWith('='))) && step.values.every(row => row.every(value => value === null || (typeof value === 'string' && value.startsWith('=')))) ? 'formula' : 'fill';
    case 'setFormula': return 'formula';
    case 'formatRange': case 'setColumnWidth': case 'freezeRows': return 'format';
    case 'clearRange': case 'deleteRows': case 'deleteColumns': return 'delete';
    case 'addSheet': case 'renameSheet': return 'sheet';
    case 'addChart': return 'chart';
    default: return 'edit';
  }
}

/** Of the text steps a batch runs on one slide, the index of the lowest non-title one (its 完善结论), or -1. */
export function conclusionIndex(steps, slideId, geometryOf = () => undefined) {
  let best = -1, lowest = -Infinity;
  steps.forEach((step, index) => {
    if (!['setText', 'addTextBox', 'addShape'].includes(step.op) || (step.op === 'addShape' && !step.text)) return;
    const where = step.slideId ?? step.slideRef ?? step.createdSlide;
    if (where !== slideId) return;
    const geometry = step.geometry ?? geometryOf(step);
    if (pptStatus(step, { shape: geometryOf(step) }) === 'title' || typeof geometry?.top !== 'number') return;
    if (geometry.top > lowest) { lowest = geometry.top; best = index; }
  });
  return best;
}
