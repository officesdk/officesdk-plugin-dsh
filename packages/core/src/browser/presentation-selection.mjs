import {needsSuiteTools} from '../agent/chat-edit-target.mjs';
import {SELECTION_LIMITS, cpLength} from '../agent/selection.mjs';
import {sameChangeAnchor} from './change-history.mjs';
import {createDocOverlay} from './doc-overlay.mjs';
import {inspectDeck, screenRect} from './presentation-agent.mjs';
import {textChunks} from './selection.mjs';
import {previewWriterPrefixes, startWriterWork, unionRects} from './selection-edit.mjs';

const stale = () => Object.assign(new Error('The selected slide text changed. Select the text again.'), {code:'source-changed'});

function textReference(slideId, shapeId, full, start = 0, end = full.length) {
  const original = full.slice(start,end);
  if (!original.trim()) return;
  if (cpLength(original) > SELECTION_LIMITS.maxText) return {unsupported:'too-long'};
  return {kind:'pptx',anchor:{slideId,shapeId,start,end},original,view:{text:full}};
}

// Key the decoration to the visible slide and the actual glyphs within this shape.
// The same shape also appears in the thumbnail. Its coordinates are never suitable here.
function slideTextBounds(document, {anchor,view}) {
  if (!anchor || !view) return;
  const {slideId,shapeId,start,end}=anchor, text=view.text;
  const surface=Array.from(document.querySelectorAll('svg[data-mop-svg-definition-namespace^="mop-main"]'))
    .find(node=>node.dataset.mopSvgSlide === slideId);
  const body=Array.from(surface?.querySelectorAll('[data-mop-svg-text-body]') ?? [])
    .find(node=>node.dataset.mopLogicalId === shapeId);
  if (!body) return;
  const rects=[];
  let from=0;
  for (const run of body.querySelectorAll('text')) {
    const value=run.textContent;
    if (!value) continue;
    const at=text.indexOf(value,from);
    if (at<0) continue;
    from=at+value.length;
    const lo=Math.max(at,start), hi=Math.min(from,end);
    if (hi<=lo) continue;
    const walker=document.createTreeWalker(run,4);
    let offset=at;
    while (walker.nextNode()) {
      const node=walker.currentNode, next=offset+node.length;
      if (next>lo && offset<hi) {
        const range=document.createRange();
        range.setStart(node,Math.max(0,lo-offset));
        range.setEnd(node,Math.min(node.length,hi-offset));
        const box=range.getBoundingClientRect();
        if(box.width && box.height) rects.push({left:box.left,top:box.top,width:box.width,height:box.height,run});
      }
      offset=next;
    }
  }
  if (rects.length) return {...unionRects(rects),rects,clip:surface.getBoundingClientRect()};
  if (start===0 && end===text.length) {
    const box=screenRect(document,{id:shapeId},slideId);
    if(box) return {...box,rects:[box]};
  }
}

export function presentationSelectionMethods({api, document, animate, locale, wait, locked}) {
  const overlay=createDocOverlay(document,{showStatus:false,animate,wait});
  let composing;
  const started=e=>{ if(e.target.matches?.('input[data-realm="ppt-editor"]')) composing=e.target; };
  const ended=e=>{ if(composing===e.target) composing=undefined; };
  document.addEventListener('compositionstart',started,true);
  document.addEventListener('compositionend',ended,true);
  const locate=ref=>slideTextBounds(document,ref);
  const beginSelectionWork=ref=>startWriterWork({overlay,locale},()=>locate(ref),true,true);
  const shapeRange=(ctx,anchor)=>ctx.presentation.slides.getItem(anchor.slideId).shapes.getItem(anchor.shapeId).textFrame.textRange;
  // Validate against the live shape, including its text around the selection. A
  // different edit can shift offsets. Only rebase a uniquely identifiable passage.
  function resolve(full, anchor, original, prior) {
    if (full.slice(anchor.start,anchor.end)===original && (prior===undefined || full===prior)) return anchor;
    if (!original || prior===undefined || prior.slice(anchor.start,anchor.end)!==original) throw stale();
    const i=full.indexOf(original);
    if (i<0 || i!==full.lastIndexOf(original)) throw stale();
    const prefix=prior.slice(0,anchor.start), suffix=prior.slice(anchor.end);
    if (!full.slice(0,i).endsWith(prefix.slice(-12)) && !full.slice(i+original.length).startsWith(suffix.slice(0,12))) throw stale();
    return {...anchor,start:i,end:i+original.length};
  }
  async function access(ref, operation) {
    return api().PowerPoint.run(async ctx=>{
      const whole=shapeRange(ctx,ref.anchor);
      whole.load('text');await ctx.sync();
      const full=whole.text;
      const anchor=resolve(full,ref.anchor,ref.original,ref.view?.text);
      const selected=whole.getSubstring(anchor.start,anchor.end-anchor.start);
      selected.load('text');await ctx.sync();
      if (selected.text!==ref.original) throw stale();
      return operation(ctx,selected,full,anchor);
    });
  }
  const write=(ref,text,signal)=>locked(()=>access(ref,async(ctx,range,full,anchor)=>{
    if(signal?.aborted) throw Object.assign(new Error('Stopped by the user.'),{code:'cancelled'});
    // The pinned PPT implementation can select a substring but its text setter
    // only writes an entire shape. Send the selection through the slide editor's
    // own text input so neighbouring runs retain their exact formatting.
    const input=document.querySelector('input[data-realm="ppt-editor"][data-debug-role="text-input"]');
    if (!input || composing) throw Object.assign(new Error('Finish entering the current slide text and retry.'),{code:'busy'});
    range.setSelected(); await ctx.sync();
    const key=key=>input.dispatchEvent(new document.defaultView.KeyboardEvent('keydown',{bubbles:true,cancelable:true,key,code:key}));
    if (!text && anchor.end!==anchor.start) key('Backspace');
    for (const [i,line] of text.split('\n').entries()) {
      if (i) key('Enter');
      if (!line) continue;
      input.value=line;
      input.dispatchEvent(new document.defaultView.InputEvent('input',{bubbles:true,inputType:'insertText',data:line}));
    }
    const whole=shapeRange(ctx,anchor);
    whole.load('text');await ctx.sync();
    if (whole.text!==full.slice(0,anchor.start)+text+full.slice(anchor.end)) throw new Error('The slide editor did not apply the selected text.');
    return {anchor:{...anchor,end:anchor.start+text.length},view:{text:full.slice(0,anchor.start)+text+full.slice(anchor.end)}};
  }));
  const recordRef=r=>({...r.revert,view:r.located?.view});
  return {
    locate, beginSelectionWork, dispose() {
      overlay.dispose();
      document.removeEventListener('compositionstart',started,true);
      document.removeEventListener('compositionend',ended,true);
    },
    async readSelection() {
      return api().PowerPoint.run(async ctx=>{
        const slides=ctx.presentation.getSelectedSlides(), shapes=ctx.presentation.getSelectedShapes();
        slides.load('items/id');shapes.load('items/id,items/name,items/left,items/top,items/width,items/height');
        await ctx.sync();
        if (slides.items.length!==1) return;
        const slide=slides.items[0].id;
        const selected=ctx.presentation.getSelectedTextRangeOrNullObject?.();
        if (selected) {
          selected.load('text,start,length');await ctx.sync();
          if (!selected.isNullObject && selected.length>0) {
            const shape=selected.getParentTextFrame().getParentShape();shape.load('id');await ctx.sync();
            const whole=shapeRange(ctx,{slideId:slide,shapeId:shape.id});whole.load('text');await ctx.sync();
            const ref=textReference(slide,shape.id,whole.text,selected.start,selected.start+selected.length);
            return ref?.original===selected.text || ref?.unsupported ? ref : undefined;
          }
          if (!selected.isNullObject) return;
        }
        if (!shapes.items.length) return;
        if (shapes.items.length===1) {
          const shape=shapes.items[0], whole=shapeRange(ctx,{slideId:slide,shapeId:shape.id});
          try {whole.load('text');await ctx.sync();if(whole.text?.trim()) return textReference(slide,shape.id,whole.text);} catch {}
        }
        // Non-text objects keep their existing contextual editing entry.
        const geometry=shapes.items.map(({id,left,top,width,height})=>({id,left,top,width,height}));
        const original=shapes.items.map(s=>`${s.name || s.id} (slide: ${slide}, shape: ${s.id})`).join('\n');
        return {kind:'document-context',original,slide,geometry};
      });
    },
    async chatSelection(instruction,selection) {
      if (needsSuiteTools(instruction)) return;
      const subject=instruction.split(/(?:改成|改为|替换为|换成|\bto\b|\bwith\b)/i)[0];
      const quote=subject.match(/[「“"]([^」”"\n]+)[」”"]/u)?.[1];
      const title=/(?:标题|\btitle\b)/i.test(subject);
      const page=subject.match(/第\s*([一二三四五六七八九十千百零〇两\d]+)\s*[页张]|\bslide\s+(\d+)/i);
      if (!quote && !title) return !page && selection?.kind==='pptx' ?
        access(selection,(_ctx,_range,full,anchor)=>({...selection,anchor,view:{text:full}})) : undefined;
      const deck=await inspectDeck(api());
      const ordinal=page?.[1] ?? page?.[2];
      const position=ordinal && (/^\d+$/.test(ordinal) ? Number(ordinal) : '一二三四五六七八九十'.split('').indexOf(ordinal)+1);
      const slide=page ? deck.slides[position-1] : deck.slides.find(s=>s.id===deck.selectedSlideIds?.[0]);
      if(!slide) return;
      const hits=slide.shapes.filter(s=>s.text && (quote ? s.text.includes(quote) :
        /(?:副标题|subtitle)/i.test(subject) ? /subtitle|副标题/i.test(s.name) : /title|标题/i.test(s.name) && !/subtitle|副标题/i.test(s.name)));
      if(hits.length!==1) return;
      const shape=hits[0], start=quote ? shape.text.indexOf(quote) : 0;
      if(quote && start!==shape.text.lastIndexOf(quote)) return;
      return textReference(slide.id,shape.id,shape.text,start,start+(quote ?? shape.text).length);
    },
    async rewriteSelection(args,signal,prepared,captured,records) {
      if(args.text===args.original) return {reply:{applied:false,before:args.original,after:args.text}};
      const ref={...args,view:captured?.view};
      if (!ref.view) ref.view=await access(ref,(_ctx,_range,full)=>({text:full}));
      const work=prepared ?? beginSelectionWork(ref);
      try {
        await work.ready;
        await overlay.phase(work.key,'clear',160);await overlay.phase(work.key,'write');
        const parts=animate() && args.text ? textChunks(args.text,{min:1,max:24}) : [args.text];
        const staged=animate() ? await previewWriterPrefixes(overlay,work.key,parts,signal) : {parts};
        let after, live=ref;
        try {
          for (const prefix of staged.parts) {
            after=await write(live,prefix,signal);
            live={...after,kind:'pptx',original:prefix};
            if (!staged.reserved) work.point(()=>locate(live));
            if (prefix!==args.text) await overlay.phase(work.key,'write',36);
          }
        } catch(error) {
          if(after) await write(live,args.original);
          throw error;
        }
        const previous=records?.findLast(r=>r.kind==='pptx' && sameChangeAnchor(r.revert?.anchor,ref.anchor) && r.after===args.original);
        return {previousId:previous?.id,reply:{applied:true,before:args.original,after:args.text},
          revert:{anchor:after.anchor,original:args.text,text:args.original},located:after};
      } finally { if(!prepared) work.dispose(); }
    },
    async selectChange(record) {
      const ref=recordRef(record);
      return access(ref,async(ctx,range,full,anchor)=>{
        ctx.presentation.setSelectedSlides([anchor.slideId]); await ctx.sync();
        range.setSelected(); await ctx.sync();
        return {kind:'pptx',anchor,original:ref.original,view:{text:full}};
      });
    },
    async changeLocations(records) {
      const locations=new Map();
      for(const record of records.filter(r=>r.kind==='pptx')) {
        try {
          const location=await access(recordRef(record),(_ctx,_r,full,anchor)=>({anchor,view:{text:full}}));
          locations.set(record.id,()=>locate(location));
        } catch { /* A manual edit may have replaced this passage. */ }
      }
      return locations;
    },
    revertSelection: (revert,signal,located)=>write({...revert,view:located?.view},revert.text,signal),
  };
}
