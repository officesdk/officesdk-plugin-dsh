// The editor's current anchor names the change. Keep the first undo payload and the last location.
export function sameChangeAnchor(a, b) {
  if (!a || !b) return false;
  if (a.paragraphs && b.paragraphs) return a.paragraphs[0] === b.paragraphs[0] && a.paragraphs[1] === b.paragraphs[1] &&
    a.start === b.start && a.end === b.end;
  if (a.slideId && b.slideId) return a.slideId === b.slideId && a.shapeId === b.shapeId && a.start === b.start && a.end === b.end;
  return typeof a.sheet === 'string' && a.sheet === b.sheet && a.row === b.row && a.column === b.column &&
    a.rowCount === b.rowCount && a.columnCount === b.columnCount;
}

// Within a changed paragraph, keep a unique occurrence with a matching neighbour
// or an unchanged prefix/suffix (including the start or end of the paragraph).
export function rebaseChangeText(text, original, captured, anchor) {
  if (!original || typeof text!=='string' || typeof captured!=='string') return;
  const pos=text.indexOf(original);
  if (pos<0 || text.indexOf(original,pos+1)>=0) return;
  const chars=Array.from(captured), prefix=chars.slice(0,anchor.start).join(''), suffix=chars.slice(anchor.end).join('');
  const before=text.slice(0,pos), after=text.slice(pos+original.length);
  if (before===prefix || after===suffix || (prefix && before.endsWith(prefix.slice(-12))) || (suffix && after.startsWith(suffix.slice(0,12)))) {
    const start=Array.from(before).length;
    return {...anchor,start,end:start+Array.from(original).length};
  }
}

// A later, wider edit can replace an earlier edit's glyphs entirely. Its current
// geometry is also the visual home of the older card. Undo still checks its own text.
export function overlapsPreviousChange(older, newer) {
  const a=older.revert?.anchor, b=newer.revert?.anchor;
  if (!a || !b || older.kind!==newer.kind) return false;
  if (older.kind==='xlsx') return a.sheet===b.sheet && a.row<b.row+b.rowCount && b.row<a.row+a.rowCount &&
    a.column<b.column+b.columnCount && b.column<a.column+a.columnCount;
  if (!a.paragraphs || !b.paragraphs || typeof newer.before!=='string') return false;
  const changed=newer.located?.view?.texts, previous=older.located?.view;
  if (!changed?.length || !previous?.texts) return false;
  const first=b.paragraphs[0], before=newer.before.split('\n');
  const end=Array.from(before.at(-1)).length+(before.length===1?b.start:0), last=first+before.length-1;
  if (a.paragraphs[1]<first || a.paragraphs[0]>last) return false;
  before[0]=Array.from(changed[0]).slice(0,b.start).join('')+before[0];
  before[before.length-1]+=Array.from(changed.at(-1)).slice(b.end).join('');
  let range=a;
  for (let i=Math.max(a.paragraphs[0],first);i<=Math.min(a.paragraphs[1],last);i++) {
    const oldLine=previous.texts[i-previous.first], beforeLine=before[i-first];
    if (oldLine===beforeLine) continue;
    if (a.paragraphs[0]!==a.paragraphs[1]) return false;
    range=rebaseChangeText(beforeLine,older.revert.original,oldLine,a);
    if (!range) return false;
  }
  return (a.paragraphs[1]!==first || range.end>b.start) && (a.paragraphs[0]!==last || range.start<end);
}

export function recordChange(records, change, previousId) {
  const index = previousId == null ? -1 : records.findIndex(r => r.id === previousId && r.kind === change.kind);
  if (index < 0) return [...records, change];
  const previous = records[index];
  const revert = change.kind !== 'xlsx'
    ? {...change.revert, text:previous.revert.text}
    : {...change.revert, cells:previous.revert.cells};
  const updated = {...change, id:previous.id, latestId:change.latestId ?? change.id, before:previous.before, revert};
  return records.map((record, i) => i === index ? updated : record);
}

// Older tabs may contain a separate card for each round. Replay only provable chains:
// the new location, with its replacement reversed, must reproduce the older paragraph snapshot.
export function restoreChangeHistory(records) {
  return records.reduce((result, change) => {
    let previous;
    if (change.kind === 'docx' && !change.latestId && change.revert?.anchor?.paragraphs && typeof change.before === 'string' && change.located?.view?.texts?.length) {
      const {anchor} = change.revert, lines = change.located.view.texts, old = change.before.split('\n');
      const first = anchor.paragraphs[0], start = anchor.start;
      const beforeAnchor = {paragraphs:[first,first+old.length-1], start,
        end:Array.from(old.at(-1)).length + (old.length === 1 ? start : 0)};
      old[0] = Array.from(lines[0]).slice(0,start).join('') + old[0];
      old[old.length-1] += Array.from(lines.at(-1)).slice(anchor.end).join('');
      previous = result.findLast(r => r.kind === change.kind && r.after === change.before &&
        sameChangeAnchor(r.revert?.anchor,beforeAnchor) && r.located?.view?.texts?.length === old.length &&
        old.every((line,i) => r.located.view.texts[i] === line));
    } else if (change.kind === 'xlsx' && !change.latestId && change.revert) {
      previous = result.findLast(r => r.kind === change.kind && r.after === change.before &&
        sameChangeAnchor(r.revert?.anchor,change.revert.anchor));
    }
    return recordChange(result,change,previous?.id);
  }, []);
}
