/**
 * Live-deck copy constraints (shimo-pptx SKILL.md §4「文案」) as a test oracle: large titles (cover, photo-statement)
 * are two lines of 2–6 characters (content-page titles are one line and not checked here), and no generation-process
 * wording on a finished slide.
 */
const PROCESS = /限流|裁切失败|找不到图|未写入|封面已出现|八页都已写入|裁切没问题/;

export function checkTitle(title) {
  const lines = String(title ?? '').split('\n').map(line => line.trim()).filter(Boolean);
  const issues = [];
  if (lines.length !== 2) issues.push('标题须两行');
  for (const line of lines) {
    const n = [...line].length;
    if (n < 2 || n > 6) issues.push(`「${line}」${n} 字，每行须 2–6 字`);
  }
  return issues;
}

export function checkBody(text) {
  const issues = [];
  if (PROCESS.test(String(text ?? ''))) issues.push('成品页不要写过程说明');
  return issues;
}
