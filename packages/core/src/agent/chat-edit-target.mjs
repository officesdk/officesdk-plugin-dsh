// A chat message can identify a preview range. Only the subsequent editor command can write it.
const editVerb=/改|润色|扩写|缩写|缩短|精简|重写|替换|换成|换为|删|补充|调整|优化|翻译|纠正|校对|扩充|[增添]加|\b(?:rewrite|replace|revise|edit|change|shorten|expand|polish|translate|delete|correct|reword)\b/i;
export const isChatEdit=text=>typeof text==='string' && editVerb.test(text) &&
  !/(?:不要|不许|无需|别)(?:直接|实际|自动)?(?:修改|改写|改动|编辑)|do not (?:edit|modify|change)/i.test(text);

// Local rewrites replace values or text. Structural or formatting commands belong
// to the regular suite tools even when a cell or a text box happens to be selected.
export const needsSuiteTools = instruction => /(?:字号|字体|加粗|斜体|字体颜色|背景色|底色|字色|设为.{0,2}色|变[红蓝绿黄黑白]色?|行高|列宽|合并单元格|拆分单元格|[删增插添][除加入添]?.{0,3}(?:行|列|工作表|幻灯片)|图表|动画|布局|\b(?:font|bold|italic|color|colour|layout|chart|animation|merge cells|insert (?:rows?|columns?|slides?))\b)/i.test(instruction);

export const usesChatSelection=text=>!/(?:标题|第\s*[\d一二三四五六七八九十]+\s*(?:自然)?段|首段|末段|最后一段|全文|整篇|整份文档|整个文档|\btitle\b|\bparagraph\s+\d|(?:whole|entire) document|(?:first|last) paragraph)/i.test(text)
  && !/[「“"‘]/.test(text.split(/改[成为]|替换[成为]|改写[成为]|\bto\b|\bwith\b/i)[0]);

// A document ordinal is an address, not an instruction to search again inside the excerpt.
// Replace only the primary subject, leaving the user's style requirements and quoted text intact.
const paragraphNumber = /第\s*[一二三四五六七八九十百零〇\d]+\s*(?:个\s*)?(?:自然)?段|首段|末段|最后一段|\bparagraph\s+\d+\b|\b(?:first|second|last) paragraph\b/gi;
const sourceClause = instruction => instruction.split(/(?:改[成为]|替换[成为]|改写[成为]|换成|改为|[「“"‘])/u)[0];
const directVerb = /(?:改写|改|重写|修改|扩写|润色|精简|缩写|缩短|调整|优化|翻译|校对|改编|rewrite|revise|edit|expand|shorten|polish|reword)/i;

export function resolvedSelectionInstruction(request) {
  const {instruction,anchor,kind} = request;
  if (kind !== 'docx' || anchor.paragraphs[0] !== anchor.paragraphs[1]) return instruction;
  const clause = sourceClause(instruction);
  const addresses = [...clause.matchAll(paragraphNumber)];
  if (addresses.length !== 1) return instruction;
  const {index,0:address} = addresses[0];
  const prefix = clause.slice(0,index), suffix = clause.slice(index+address.length);
  // A reference used for comparison or wording is not the primary target.
  if (/(?:参考|参照|按照|模仿|对照|保留|保留原文|提到|像|例如|like|based on|refer to)\s*$/i.test(prefix)) return instruction;
  if (!directVerb.test(prefix) && !directVerb.test(suffix) && !/^(?:.*[把将对]\s*)$/.test(prefix)) return instruction;
  const reference = /[a-z]/i.test(address) ? 'the supplied passage' : '当前提供的选区';
  return instruction.slice(0,index)+reference+instruction.slice(index+address.length);
}

function ordinal(value) {
  if (/^\d+$/.test(value)) return Number(value);
  const digit=c=>'零一二三四五六七八九'.indexOf(c);
  if(value.includes('十')){const [t,u]=value.split('十');return (t?digit(t):1)*10+(u?digit(u):0);}
  return digit(value);
}
export function chatTarget(instruction,paragraphs,selection) {
  if (!isChatEdit(instruction)) return;
  const texts=paragraphs.map(p=>p.text);
  const usable=texts.map((_,i)=>i).filter(i=>!paragraphs[i].table && texts[i].trim());
  if (!usable.length) return;
  const first=usable[0];
  const make=(a,b=a,start=0,end=Array.from(texts[b]).length)=>{
    const lines=texts.slice(a,b+1);
    const original=a===b?Array.from(lines[0]).slice(start,end).join(''):
      [Array.from(lines[0]).slice(start).join(''),...lines.slice(1,-1),Array.from(lines.at(-1)).slice(0,end).join('')].join('\n');
    return {kind:'docx',original,anchor:{paragraphs:[a,b],start,end},view:{first:a,texts:lines}};
  };
  // Quotes after “改成” name the replacement. Only quote text occurring uniquely in the document is a source.
  const source=instruction.split(/(?:改[成为]|替换[成为]|改写[成为]|\bwith\b)/i)[0];
  for (const match of source.matchAll(/[「“"‘]([^」”"’\n]+)[」”"’]/g)) {
    const needle=match[1], where=usable.filter(i=>texts[i].includes(needle));
    if(!where.length)continue;
    if(where.length!==1 || texts[where[0]].indexOf(needle)!==texts[where[0]].lastIndexOf(needle))return;
    const i=where[0], start=Array.from(texts[i].slice(0,texts[i].indexOf(needle))).length;
    return make(i,i,start,start+Array.from(needle).length);
  }
  const body=usable.filter(i=>!/(?:heading|title|标题)/i.test(paragraphs[i].style??''));
  const index=instruction.match(/第\s*([一二三四五六七八九十\d]+)\s*(?:自然)?段|\bparagraph\s+(\d+)/i);
  if(index || /(?:首段|第一自然段|最后一段|末段|\b(?:first|last) paragraph\b)/i.test(instruction)){
    const n=index?ordinal(index[1]??index[2]):/(?:最后一段|末段|last paragraph)/i.test(instruction)?body.length:1;
    if(body[n-1]!==undefined)return make(body[n-1]);
    return;
  }
  if(/(?:文档标题|主标题|^.{0,8}标题|\b(?:the |document )?title\b)/i.test(source))return make(first);
  if(/(?:全文|整篇|整份文档|整个文档|whole document|entire document)/i.test(instruction))return make(first,usable.at(-1));
  if(selection?.kind==='docx' && selection.original && selection.anchor){
    const {paragraphs:[a,b],start,end}=selection.anchor;
    if(a>=0 && b<texts.length && a<=b){const ref=make(a,b,start,end);if(ref.original===selection.original)return ref;}
  }
}
