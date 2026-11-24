/**
 * The expression grammar of a style pack's layouts.json (styles/<id>/layouts.json): a whitelisted subset of Python
 * expression syntax that skills/shimo-pptx/scripts/brand_deck.py interprets. This side only parses and checks it
 * (and folds expressions without names), so the Host and CI can reject a bad layout description without Python.
 * Keep both sides in step: tooling/style-distill/fixtures/expr-cases.json holds cases both must judge the same way.
 */
export const EXPR_FUNCTIONS = Object.freeze(['len', 'str', 'int', 'float', 'min', 'max', 'abs', 'round', 'lines', 'line', 'tw', 'fit', 'pad2', 'fmt', 'year', 'pct']);
const KEYWORDS = new Set(['and', 'or', 'not', 'if', 'else', 'in', 'True', 'False', 'None', 'lambda', 'for', 'is', 'await', 'yield']);
// Python keywords that are never names: an expression using one as a name is rejected (as Python's parser does).
const PY_RESERVED = new Set(['as', 'assert', 'async', 'break', 'class', 'continue', 'def', 'del', 'elif', 'except', 'finally', 'from', 'global', 'import', 'nonlocal', 'pass', 'raise', 'return', 'try', 'while', 'with']);
const OPERATORS = ['//', '==', '!=', '<=', '>=', '**', '+', '-', '*', '/', '%', '<', '>', '(', ')', '[', ']', '{', '}', ',', ':', '.', '='];

export class ExprError extends Error {
  constructor(source, message) { super(`表达式 ${JSON.stringify(source)}：${message}`); this.name = 'ExprError'; }
}

function tokenize(source) {
  const tokens = [];
  let i = 0;
  while (i < source.length) {
    const c = source[i];
    if (c === ' ' || c === '\t') { i += 1; continue; }
    if (/[0-9]/.test(c) || (c === '.' && /[0-9]/.test(source[i + 1] ?? ''))) {
      const m = /^(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?/.exec(source.slice(i));
      if (/^0\d/.test(m[0])) throw new ExprError(source, `整数不能以 0 开头：${m[0]}`);
      if (/^[A-Za-z_]/.test(source.slice(i + m[0].length))) throw new ExprError(source, `数字写法不支持：${source.slice(i, i + m[0].length + 1)}`);
      tokens.push({ kind: 'num', value: Number(m[0]) }); i += m[0].length; continue;
    }
    if (c === '"' || c === "'") {
      let j = i + 1, text = '';
      for (; j < source.length && source[j] !== c; j += 1) {
        if (source[j] === '\n' || source[j] === '\r') throw new ExprError(source, '字符串里不能直接换行（写 \\n）');
        if (source[j] === '\\') {
          const next = source[j + 1];
          const map = { n: '\n', t: '\t', '\\': '\\', "'": "'", '"': '"' };
          if (!(next in map)) throw new ExprError(source, `字符串里不支持转义 \\${next ?? ''}`);
          text += map[next]; j += 1;
        } else text += source[j];
      }
      if (j >= source.length) throw new ExprError(source, '字符串没有结束引号');
      tokens.push({ kind: 'str', value: text }); i = j + 1; continue;
    }
    if (/[A-Za-z_]/.test(c)) {
      // ASCII names only (Python also allows other letters; layouts.json does not need them, so this side stays stricter).
      const m = /^[A-Za-z_][A-Za-z0-9_]*/.exec(source.slice(i));
      if (PY_RESERVED.has(m[0])) throw new ExprError(source, `不允许 ${m[0]}`);
      if (source[i + m[0].length] === '"' || source[i + m[0].length] === "'") throw new ExprError(source, `不支持带前缀的字符串（${m[0]}'…'）`);
      tokens.push({ kind: KEYWORDS.has(m[0]) ? 'kw' : 'name', value: m[0] }); i += m[0].length; continue;
    }
    const op = OPERATORS.find(candidate => source.startsWith(candidate, i));
    if (!op) throw new ExprError(source, `不认识的字符 ${JSON.stringify(c)}`);
    tokens.push({ kind: 'op', value: op }); i += op.length;
  }
  tokens.push({ kind: 'end' });
  return tokens;
}

/** Parse and whitelist-check; returns the tree. Throws ExprError naming what is not allowed. */
export function parseExpr(source) {
  if (typeof source !== 'string') throw new ExprError(String(source), '不是字符串');
  const tokens = tokenize(source);
  let at = 0;
  const peek = () => tokens[at];
  const is = (kind, value) => peek().kind === kind && (value === undefined || peek().value === value);
  const eat = (kind, value) => {
    if (!is(kind, value)) throw new ExprError(source, `这里应是 ${value ?? kind}，却是 ${peek().value ?? '结尾'}`);
    return tokens[at++];
  };
  const forbid = () => {
    const t = peek();
    if (t.kind === 'kw' && ['lambda', 'for', 'is', 'await', 'yield'].includes(t.value)) throw new ExprError(source, `不允许 ${t.value}`);
    if (t.kind === 'op' && (t.value === '**' || t.value === '=')) throw new ExprError(source, `不允许运算符 ${t.value}`);
  };

  function expression() {
    const body = orExpr();
    if (is('kw', 'if')) {
      at += 1;
      const test = orExpr();
      eat('kw', 'else');
      return { type: 'if', test, body, orelse: expression() };
    }
    forbid();
    return body;
  }
  function orExpr() {
    let node = andExpr();
    while (is('kw', 'or')) { at += 1; node = { type: 'bool', op: 'or', values: [node, andExpr()] }; }
    return node;
  }
  function andExpr() {
    let node = notExpr();
    while (is('kw', 'and')) { at += 1; node = { type: 'bool', op: 'and', values: [node, notExpr()] }; }
    return node;
  }
  function notExpr() {
    if (is('kw', 'not')) { at += 1; return { type: 'unary', op: 'not', operand: notExpr() }; }
    return comparison();
  }
  function comparison() {
    const left = arith();
    const ops = [], rights = [];
    for (;;) {
      if (is('op') && ['==', '!=', '<', '<=', '>', '>='].includes(peek().value)) { ops.push(tokens[at++].value); rights.push(arith()); continue; }
      if (is('kw', 'in')) { at += 1; ops.push('in'); rights.push(arith()); continue; }
      if (is('kw', 'not') && tokens[at + 1].kind === 'kw' && tokens[at + 1].value === 'in') { at += 2; ops.push('not in'); rights.push(arith()); continue; }
      forbid();
      break;
    }
    return ops.length ? { type: 'compare', left, ops, rights } : left;
  }
  function arith() {
    let node = term();
    while (is('op', '+') || is('op', '-')) { const op = tokens[at++].value; node = { type: 'bin', op, left: node, right: term() }; }
    return node;
  }
  function term() {
    let node = unary();
    while (is('op') && ['*', '/', '//', '%'].includes(peek().value)) { const op = tokens[at++].value; node = { type: 'bin', op, left: node, right: unary() }; }
    forbid();
    return node;
  }
  function unary() {
    if (is('op', '-') || is('op', '+')) { const op = tokens[at++].value; return { type: 'unary', op, operand: unary() }; }
    return postfix();
  }
  function args() {
    const list = [];
    eat('op', '(');
    while (!is('op', ')')) {
      if (is('op', '*') || is('op', '**')) throw new ExprError(source, '调用不能用星号展开');
      if (is('name') && tokens[at + 1].kind === 'op' && tokens[at + 1].value === '=') throw new ExprError(source, '调用不能用关键字参数');
      list.push(expression());
      if (!is('op', ')')) eat('op', ',');
    }
    eat('op', ')');
    return list;
  }
  function postfix() {
    let node = atom();
    for (;;) {
      if (is('op', '(')) {
        if (node.type !== 'name') throw new ExprError(source, '只能调用内置函数或 .get()');
        if (!EXPR_FUNCTIONS.includes(node.id)) throw new ExprError(source, `调用了不允许的函数 ${node.id}`);
        node = { type: 'call', fn: node.id, args: args() };
      } else if (is('op', '[')) {
        at += 1;
        const bound = () => (is('op', ':') || is('op', ']') ? null : expression());
        const lower = bound();
        if (is('op', ':')) {
          at += 1;
          const upper = bound();
          let step = null;
          if (is('op', ':')) { at += 1; step = bound(); }
          eat('op', ']');
          node = { type: 'slice', value: node, lower, upper, step };
        } else {
          if (lower === null) throw new ExprError(source, '下标不能为空');
          let index = lower;
          if (is('op', ',')) {
            const items = [lower];
            while (is('op', ',')) { at += 1; if (is('op', ']')) break; items.push(expression()); }
            index = { type: 'tuple', items };
          }
          eat('op', ']');
          node = { type: 'index', value: node, index };
        }
      } else if (is('op', '.')) {
        at += 1;
        const attr = eat('name').value;
        if (attr !== 'get') throw new ExprError(source, `只允许 .get()，不允许 .${attr}`);
        if (!is('op', '(')) throw new ExprError(source, '.get 必须直接调用');
        const list = args();
        if (list.length < 1 || list.length > 2) throw new ExprError(source, '.get() 带 1–2 个参数');
        node = { type: 'get', object: node, args: list };
      } else return node;
    }
  }
  function sequence(close) {
    const items = [];
    let comma = false;
    while (!is('op', close)) {
      items.push(expression());
      if (is('kw', 'for')) throw new ExprError(source, '不允许推导式');
      if (is('op', ',')) { at += 1; comma = true; } else break;
    }
    eat('op', close);
    return { items, comma };
  }
  function atom() {
    const t = peek();
    if (t.kind === 'num' || t.kind === 'str') {
      at += 1;
      if (t.kind === 'str' && is('str')) throw new ExprError(source, '不支持相邻字符串拼接');
      return { type: 'const', value: t.value };
    }
    if (t.kind === 'kw' && ['True', 'False', 'None'].includes(t.value)) { at += 1; return { type: 'const', value: t.value === 'True' ? true : t.value === 'False' ? false : null }; }
    if (t.kind === 'name') { at += 1; return { type: 'name', id: t.value }; }
    if (t.kind === 'op' && t.value === '(') {
      at += 1;
      const { items, comma } = sequence(')');
      if (!comma && items.length === 1) return items[0];
      return { type: 'tuple', items };
    }
    if (t.kind === 'op' && t.value === '[') { at += 1; return { type: 'list', items: sequence(']').items }; }
    if (t.kind === 'op' && t.value === '{') {
      at += 1;
      const keys = [], values = [];
      while (!is('op', '}')) {
        if (is('op', '*') || is('op', '**')) throw new ExprError(source, '字典不能用星号展开');
        keys.push(expression());
        if (!is('op', ':')) throw new ExprError(source, '只支持字典，不支持集合');
        at += 1;
        values.push(expression());
        if (is('kw', 'for')) throw new ExprError(source, '不允许推导式');
        if (!is('op', '}')) eat('op', ',');
      }
      eat('op', '}');
      return { type: 'dict', keys, values };
    }
    forbid();
    throw new ExprError(source, `这里不能出现 ${t.value ?? '结尾'}`);
  }

  const tree = expression();
  if (!is('end')) { forbid(); throw new ExprError(source, `多余的 ${peek().value}`); }
  return tree;
}

/** Every variable name an expression reads (function names and `.get` excluded). */
export function exprNames(tree, out = new Set()) {
  if (!tree || typeof tree !== 'object') return out;
  if (tree.type === 'name') { out.add(tree.id); return out; }
  for (const [key, value] of Object.entries(tree)) {
    if (key === 'type' || key === 'fn') continue;
    if (Array.isArray(value)) value.forEach(item => exprNames(item, out));
    else if (value && typeof value === 'object') exprNames(value, out);
  }
  return out;
}

const NUMERIC = { '+': (a, b) => a + b, '-': (a, b) => a - b, '*': (a, b) => a * b, '/': (a, b) => a / b, '//': (a, b) => Math.floor(a / b), '%': (a, b) => a - b * Math.floor(a / b) };
/** The value of an expression made only of numbers and arithmetic, else undefined (geometry that depends on content). */
export function foldNumber(tree) {
  if (tree.type === 'const') return typeof tree.value === 'number' ? tree.value : undefined;
  if (tree.type === 'unary' && tree.op !== 'not') { const v = foldNumber(tree.operand); return v === undefined ? undefined : tree.op === '-' ? -v : v; }
  if (tree.type === 'bin') {
    const a = foldNumber(tree.left), b = foldNumber(tree.right);
    return a === undefined || b === undefined ? undefined : NUMERIC[tree.op](a, b);
  }
  return undefined;
}
