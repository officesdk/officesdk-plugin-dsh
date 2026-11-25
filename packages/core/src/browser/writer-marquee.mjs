import { createWriterInlinePreview } from './writer-inline-preview.mjs';
import { WORK_BORDER_STOPS } from './work-border.mjs';

// Writer draws separate SVG text runs even on a single line. Join runs on their visual baseline.
function rowsOf(rects) {
  const rows=[];
  for (const rect of rects) {
    const row=rows.at(-1), right=rect.left+rect.width, bottom=rect.top+rect.height;
    if (row && Math.min(row.bottom,bottom)-Math.max(row.top,rect.top) > Math.min(rect.height,row.bottom-row.top)*.6) {
      row.left=Math.min(row.left,rect.left); row.right=Math.max(row.right,right);
      row.top=Math.min(row.top,rect.top); row.bottom=Math.max(row.bottom,bottom);
    } else rows.push({left:rect.left,top:rect.top,right,bottom});
  }
  return rows;
}

function contour(rows) {
  const groups=[];
  for (const row of rows) {
    const last=groups.at(-1)?.at(-1);
    if (last && Math.min(last.right,row.right)>Math.max(last.left,row.left) && row.top-last.bottom<(row.bottom-row.top)*1.5) groups.at(-1).push(row);
    else groups.push([row]);
  }
  return groups.map(group=>{
    const first=group[0], last=group.at(-1);
    const left=[[first.left,first.top]], right=[[first.right,first.top]];
    for (let i=1;i<group.length;i++) {
      const prev=group[i-1], next=group[i], mid=(prev.bottom+next.top)/2;
      if(prev.left!==next.left) left.push([prev.left,mid],[next.left,mid]);
      if(prev.right!==next.right) right.push([prev.right,mid],[next.right,mid]);
    }
    right.push([last.right,last.bottom]);left.push([last.left,last.bottom]);
    return 'M'+right.concat(left.reverse()).map(point=>point.join(' ')).join('L')+'Z';
  }).join('');
}

let serial=0;
export function writerMarquee(box) {
  const document=box.ownerDocument, ns='http://www.w3.org/2000/svg';
  const svg=document.createElementNS(ns,'svg'), path=document.createElementNS(ns,'path');
  const defs=document.createElementNS(ns,'defs'), gradient=document.createElementNS(ns,'linearGradient');
  gradient.id='office-work-'+(++serial);
  for (const [offset,color] of WORK_BORDER_STOPS) {
    const stop=document.createElementNS(ns,'stop');stop.setAttribute('offset',offset);stop.setAttribute('stop-color',color);gradient.append(stop);
  }
  const bars=document.createElementNS(ns,'path');
  bars.setAttribute('fill','none');bars.setAttribute('stroke','#b7c1cc');bars.setAttribute('stroke-linecap','round');
  const overhang=document.createElementNS(ns,'path');
  overhang.setAttribute('fill','white');
  defs.append(gradient);svg.append(defs,overhang,path,bars);box.append(svg);
  Object.assign(svg.style,{width:'100%',height:'100%',overflow:'visible',display:'block'});
  path.setAttribute('stroke','url(#'+gradient.id+')');path.setAttribute('stroke-width','1.4');path.setAttribute('stroke-linejoin','round');
  let drawPreview;
  return (found,progress,phase,ms,previewText)=>{
    const {left,top}=found, width=Math.max(1,found.width*progress), height=Math.max(1,found.height*progress);
    const rows=rowsOf(found.rects.map(r=>({...r,left:r.left-left,top:r.top-top})));
    const dragging=progress<1;
    path.setAttribute('d',dragging? 'M0 0H'+width+'V'+height+'H0Z':contour(rows));
    const placeholder=phase==='placeholder' && !dragging;
    bars.style.display=placeholder?'':'none';
    if (placeholder) {
      const h=Math.max(2,Math.min(6,(rows[0].bottom-rows[0].top)*.38));
      bars.setAttribute('stroke-width',String(h));bars.setAttribute('opacity',String(.4+.5*(1+Math.sin(ms/175))/2));
      bars.setAttribute('d',rows.map(r=>{
        const inset=Math.min(h,(r.right-r.left)/3), y=(r.top+r.bottom)/2;
        return 'M'+(r.left+inset)+' '+y+'H'+(r.right-inset);
      }).join(''));
    }
    if (previewText !== undefined && !drawPreview) drawPreview=createWriterInlinePreview(svg);
    const previewFits=drawPreview?.(found,previewText) ?? false;
    // Hide the original overhanging glyphs with the placeholder, without widening its border.
    const cover = !dragging && (placeholder || phase==='clear' || previewFits);
    const bleed = [];
    if (cover) for (const r of found.rects) {
      const right = r.left+r.width, y = r.top-top;
      if (r.inkRight > right) bleed.push('M'+(right-left)+' '+y+'H'+(r.inkRight-left)+'v'+r.height+'H'+(right-left)+'Z');
      if (r.inkLeft < r.left) bleed.push('M'+(r.inkLeft-left)+' '+y+'H'+(r.left-left)+'v'+r.height+'H'+(r.inkLeft-left)+'Z');
    }
    overhang.setAttribute('d',bleed.join(''));
    path.setAttribute('fill',previewFits?'white':dragging?'#609cff24':phase==='clear'?'white':placeholder?'#f2f4f7':'none');
    Object.assign(box.style,{display:'block',left:left+'px',top:top+'px',width:found.width+'px',height:found.height+'px'});
    const last=rows.at(-1);
    return dragging?{left,top,width,height}:{left:left+last.left,top:top+last.top,width:last.right-last.left,height:last.bottom-last.top,previewFits};
  };
}
