// Rank the requested deliverable, rather than incidental words in its source material.
const wordIntents = {
  'meeting-minutes': ['会议纪要','会议记录','例会纪要','例会记录','meeting minutes','meeting notes','会后纪要','会后整理','整理会议决定'],
  'weekly-report': ['工作周报','个人周报','工作日报','周报','日报','weekly report','weekly update','daily report'],
  'product-requirements': ['产品需求文档','产品需求','需求说明','功能需求','需求文档','product requirements','prd'],
  'project-plan': ['项目计划','项目排期','里程碑计划','实施计划','project plan','project schedule','milestone plan'],
  'project-proposal': ['项目建议书','立项建议','立项申请','项目提案','立项建议书','project proposal','business case'],
  'formal-letter': ['正式信函','商务信函','邀请函','致函','formal letter','business letter'],
  gongwen: ['公文','红头','通报','请示','official notice','notice','通知','official document'],
  blank: ['空白文档','空白页','blank document'],
};
const pptIntents = {
  'slide-002':['季度复盘','季度总结','工作总结','经营分析','工作汇报','团队汇报','经营汇报','业绩汇报','进度汇报','季度汇报','team performance','quarterly results','季度业绩','年终总结','quarterly review','business review','team update','work report','retrospective'],
  'slide-008':['员工培训','入职培训','培训','经验分享','知识分享','新人上手','新员工入职','training','onboarding','knowledge sharing','workshop'],
  'slide-000':['空白幻灯片','空白演示','空白ppt','blank deck','blank slide'],
};
function normalize(text) {
  return String(text ?? '').normalize('NFKC').toLowerCase().replace(/[-_]+/g,' ').trim();
}
function bestMatch(text, entries) {
  const query = normalize(text);
  let best;
  let high = 0;
  for (const [id,terms] of entries) {
    const hits = [...new Set(terms.map(normalize).filter(Boolean))].filter(term => query.includes(term));
    const score = hits.length ? Math.max(...hits.map(term=>term.length)) + Math.min(hits.length-1,3) : 0;
    if (score > high) { best = id; high = score; }
  }
  return best;
}
function byPurpose(intent, path, entries) {
  return bestMatch(intent, entries) ?? bestMatch(String(path).split(/[\\/]/).pop()?.replace(/\.[^.]+$/,''), entries);
}
export function wordTemplate(intent, path, ids) {
  return byPurpose(intent, path, Object.entries(wordIntents).filter(([id]) => ids.includes(id))) ?? (ids.includes('blank') ? 'blank' : 'gongwen');
}
export function pptTemplate(intent, path, packs, count) {
  const entries = packs.filter(pack => !count || !pack.count || count === pack.count).map(({id, terms})=>[id,[...(pptIntents[id] ?? []),...terms]]);
  return byPurpose(intent, path, entries);
}
