// A copy/redesign can spend its first model step reading the old file. Seat the
// new task on submission; a real file address replaces only that task's tab.
export function newPresentationRequest(text, activePresentation = false) {
  const target = /pptx?\b|幻灯片|演示文稿|演示稿|slides?\b|presentation|deck\b/i.test(text);
  const separate = /另存|另做|另[做出]一份|重新生成|(?:生成|做|导出|出|保存).{0,8}(?:一份新|另一份|新版本|新版|新\s*ppt)|new (?:version|copy|deck|presentation)|save.{0,8}(?:copy|new)/i.test(text);
  const create = /(?:做|制作|起草|写|设计|准备|出|生成|创建|新建|make|build|draft|create|generate).{0,90}(?:pptx?\b|演示文稿|幻灯片|演示稿|deck\b|slides?\b|presentation)/i.test(text);
  return !/(?:如何|怎么才能|怎么做|怎样|how (?:do|can|to)).{0,35}(?:生成|制作|做|create|make)|(?:不用|不要|不需要|无需).{0,4}(?:生成|另存|新建|做).{0,10}(?:ppt|演示|文件)|(?:原地|直接在原稿|覆盖原)/i.test(text)
    && (target && (separate || create) || activePresentation && separate);
}

export function watchPresentationTurns({composers,sidebarOf,activePresentation,begin,abandon,startedAt=Date.now()}) {
  const turns = new Map();
  function update(sessionId) {
    const entry = composers.get(sessionId);
    if (!entry || sidebarOf()?.mounted?.getSnapshot?.() !== sessionId) return;
    const node = entry.chat?.nodes?.findLast(item => item.kind === 'user' || item.kind === 'steering');
    if (!node || node.time < startedAt) return;
    const turn = node.turn ?? entry.snapshot?.timeline?.turnOrder?.at(-1);
    if (!Number.isInteger(turn)) return;
    const previous = turns.get(sessionId);
    if (previous?.turn === turn || !entry.running) return;
    const text = (node.content ?? []).filter(block => (block.kind ?? block.type) === 'text').map(block=>block.text).join(' ');
    if (!newPresentationRequest(text,activePresentation(sessionId))) return;
    if (previous && !previous.done) abandon(sessionId,previous);
    const task = {turn,done:false};
    turns.set(sessionId,task);
    begin(sessionId,task);
  }
  const unsubscribe = composers.subscribeEntries(update);
  return {turns,dispose:unsubscribe};
}
