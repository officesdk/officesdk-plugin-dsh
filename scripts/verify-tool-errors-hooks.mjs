/**
 * Module hooks for verify-tool-errors.mjs: every @deepseek-ai package the bundles import becomes a stub, because
 * a test install omits DSH's own peers. `defineTool` returns its definition so the check can call `execute`; every
 * other export is one Proxy that can be called, chained (schemastery) and extended (TypertRemoteService).
 */
let exports = {};
export function initialize(data) { exports = data.exports; }

export async function resolve(specifier, context, next) {
  return Object.hasOwn(exports, specifier) ? { url: `stub:${specifier}`, shortCircuit: true } : next(specifier, context);
}

export async function load(url, context, next) {
  if (!url.startsWith('stub:')) return next(url, context);
  const lines = ['const any = new Proxy(function stub() {}, { get: (target, key) => key === "prototype" ? target.prototype : key === Symbol.toPrimitive ? () => "" : key === "then" ? undefined : any, apply: () => any, construct: () => ({}) });'];
  for (const name of exports[url.slice('stub:'.length)]) {
    if (name === 'default') lines.push('export default any;');
    else if (name === 'defineTool') lines.push('export const defineTool = definition => definition;');
    else lines.push(`export const ${name} = any;`);
  }
  return { format: 'module', source: lines.join('\n'), shortCircuit: true };
}
