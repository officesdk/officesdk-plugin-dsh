import { baseName } from './paths.mjs';

const NODE = 'officesdk-edit-reference';
const QUOTE = /^ {0,3}>[^\r\n]*(?:\r?\n {0,3}>[^\r\n]*)*/gm;
const unquote = text => text.replace(/^ {0,3}>[ \t]?/gm, '');
const locationLine = /^(?:第 \d+(?:–\d+)? 段|paragraphs? \d+(?:–\d+)?|\S+ \/ \S+|.+!\$?[A-Z]+\$?\d+(?::\$?[A-Z]+\$?\d+)?)$/i;

function selectionRequest(text) {
  const lines = text.trim().split('\n');
  const header = /^(?:修改「([^」]+)」|(?:Rewrite the text|Change the cells) I selected in "([^"]+)")/.exec(lines[0]);
  if (!header || !/(?:选区 |selection )sel-[0-9a-f]{8}\b/i.test(lines[0])) return;
  if (!/^(?:（用 selection_rewrite 修改）|\(Answer with selection_rewrite\.\))$/.test(lines.at(-1))) return;
  const end = lines.findLastIndex(line => /^(?:要求：|Request: )/.test(line));
  if (end < 1) return;
  const name = baseName(header[1] ?? header[2], 'win32');
  if (!/\.(?:docx|xlsx|pptx)$/i.test(name)) return;
  const selected = lines.slice(1, end);
  const table = selected[0]?.startsWith('| ') && /^\|(?: --- \|)+$/.test(selected[1] ?? '');
  const original = table
    ? selected.slice(2).map(row => row.slice(2, -2).replace(/ \| /g, '\t').replace(/\\\|/g, '|')).join('\n')
    : selected.map(line => line.replace(/^> ?/, '')).join('\n');
  const instruction = lines.slice(end, -1).join('\n').replace(/^(?:要求：|Request: )/, '');
  return { name, original, instruction, en: Boolean(header[2]) };
}

function selectedText(body, name) {
  const request = selectionRequest(body);
  if (request) return request.original;
  const lines = body.replace(/^\n+/, '').split('\n');
  if (locationLine.test(lines[0])) {
    lines.shift();
    const last = lines.at(-1);
    if (last && /[\\/]/.test(last) && baseName(last, 'win32') === name) lines.pop();
  }
  return lines.join('\n');
}

// Own only Office edit references. Match the original sent text in either order without changing model input.
export const editQuoteRule = {
  id: NODE, order: -20,
  match(text, from) {
    if (from) return;
    if (!/(?:编辑目标：|Editing target: )/.test(text)) {
      if (!text.startsWith('修改「') && !text.startsWith('Rewrite the text') && !text.startsWith('Change the cells')) return;
      const request = selectionRequest(text.replace(/\r\n/g, '\n'));
      if (request) return { id: NODE, start: 0, end: text.length, data: request };
      return;
    }
    for (const match of text.matchAll(QUOTE)) {
      const decoded = unquote(match[0].replace(/\r\n/g, '\n'));
      const header = /^(编辑目标：|Editing target: )([^\n]+?\.(?:docx|xlsx|pptx))[ \t]*(?:\n|$)/i.exec(decoded);
      if (!header) continue;
      const name = baseName(header[2], 'win32');
      const instruction = [text.slice(0, match.index).trim(), text.slice(match.index + match[0].length).trim()].filter(Boolean).join('\n\n');
      return {
        id: NODE, start: 0, end: text.length,
        data: { name, original: selectedText(decoded.slice(header[0].length), name), instruction, en: header[1] === 'Editing target: ' }
      };
    }

    // Older composers without insertQuote submitted a plain target header and a copy of the tool request.
    const plain = /^(编辑目标：|Editing target: )([^\r\n]+?\.(?:docx|xlsx|pptx))[ \t]*\r?\n([\s\S]+)$/i.exec(text);
    if (!plain) return;
    const body = plain[3].trim().replace(/\r\n/g, '\n');
    const end = /^(?:（用 selection_rewrite 修改）|\(Answer with selection_rewrite\.\))$/m.exec(body);
    const request = end && selectionRequest(body.slice(0, end.index + end[0].length));
    if (!request && body.includes('\n')) return;
    return {
      id: NODE, start: 0, end: text.length,
      data: { name: baseName(plain[2], 'win32'), original: request?.original ?? '',
        instruction: request?.instruction ?? body, en: plain[1] === 'Editing target: ' }
    };
  },
};

const graphemes = new Intl.Segmenter(undefined, { granularity: 'grapheme' });

export function fitFilename(name, width, measure) {
  if (width <= 0 || measure(name) <= width) return name;
  const dot = name.lastIndexOf('.');
  const extension = dot > 0 ? name.slice(dot) : '';
  const stem = [...graphemes.segment(dot > 0 ? name.slice(0, dot) : name)].map(part => part.segment);
  function shorten(count) {
    const tail = Math.floor(count / 2);
    return stem.slice(0, count - tail).join('') + '...' + (tail ? stem.slice(-tail).join('') : '') + extension;
  }
  let low = 0;
  let high = stem.length - 1;
  while (low < high) {
    const mid = Math.ceil((low + high) / 2);
    if (measure(shorten(mid)) <= width) low = mid;
    else high = mid - 1;
  }
  return shorten(low);
}

export function registerEditQuote(scope, React) {
  if (!scope.uiConversation?.userText) return;
  const { createElement: h, useEffect, useLayoutEffect, useRef, useId, useState } = React;

  function EditQuote({ node, renderFactorySlot, interactive }) {
    const { name, original, instruction, en } = node.data;
    const label = en ? 'Editing' : '正在编辑';
    const anchor = useRef(null);
    const field = useRef(null);
    const tip = useRef(null);
    const timer = useRef(null);
    const [shown, setShown] = useState(name);
    const [open, setOpen] = useState(false);
    const id = useId();

    useLayoutEffect(() => {
      const el = field.current;
      const doc = el.ownerDocument;
      const context = doc.createElement('canvas').getContext('2d');
      function fit() {
        const style = doc.defaultView.getComputedStyle(el);
        context.font = [style.fontWeight, style.fontSize, style.fontFamily].join(' ');
        setShown(fitFilename(name, el.getBoundingClientRect().width, text => context.measureText(text).width));
      }
      fit();
      const observer = new ResizeObserver(fit);
      observer.observe(el);
      doc.fonts?.addEventListener('loadingdone', fit);
      return () => {
        observer.disconnect();
        doc.fonts?.removeEventListener('loadingdone', fit);
      };
    }, [name]);
    useEffect(() => () => clearTimeout(timer.current), []);

    useLayoutEffect(() => {
      if (!open) return;
      const el = tip.current;
      const win = el.ownerDocument.defaultView;
      function position() {
        const box = anchor.current.getBoundingClientRect();
        const rightLimit = win.innerWidth - el.offsetWidth - 8;
        const bottomLimit = win.innerHeight - el.offsetHeight - 8;
        const top = box.top >= el.offsetHeight + 14 ? box.top - el.offsetHeight - 6 : box.bottom + 6;
        el.style.left = Math.max(8, Math.min(box.left, rightLimit)) + 'px';
        el.style.top = Math.max(8, Math.min(top, bottomLimit)) + 'px';
      }
      const escape = event => { if (event.key === 'Escape') setOpen(false); };
      el.showPopover?.();
      position();
      win.addEventListener('scroll', position, true);
      win.addEventListener('resize', position);
      win.addEventListener('keydown', escape);
      return () => {
        el.hidePopover?.();
        win.removeEventListener('scroll', position, true);
        win.removeEventListener('resize', position);
        win.removeEventListener('keydown', escape);
      };
    }, [open]);

    function show() {
      clearTimeout(timer.current);
      setOpen(true);
    }
    function hide() {
      clearTimeout(timer.current);
      timer.current = setTimeout(() => setOpen(false), 120);
    }
    return h('span', { className: 'office-edit-reference' },
      h('button', { type: 'button', className: 'office-edit-reference-target', ref: anchor,
        'aria-label': label + ' ' + name, 'aria-describedby': open ? id : undefined,
        onPointerEnter: show, onPointerLeave: hide, onFocus: show, onBlur: hide, onClick: show },
        h('span', { className: 'office-edit-reference-prefix' }, label),
        h('span', { className: 'office-edit-reference-name', ref: field },
          h('span', { 'aria-hidden': true }, name),
          h('span', { className: 'office-edit-reference-visible', 'aria-hidden': true }, shown))),
      open && h('span', { id, ref: tip, role: 'tooltip', popover: 'manual', className: 'office-edit-reference-tip',
        onPointerEnter: show, onPointerLeave: hide },
        original && h('small', null, en ? 'Quoted text' : '引用原文'), original || name),
      instruction && h('span', { className: 'office-edit-instruction' }, renderFactorySlot
        ? renderFactorySlot('conversation.user.text', { text: instruction, facts: {}, interactive }) : instruction));
  }
  scope.uiConversation.userText.register(editQuoteRule);
  const slot = 'conversation.user.text.node';
  scope.slots.inject(slot, () => scope.slots.register({ name: slot, key: NODE }, EditQuote));
}
