/** Overlay only: neither edits nor saves the document. Host must explicitly complete or cancel it. */
export function createTableStage({ document, overlay, locate, scheduler = document.defaultView }) {
  let mask, frame, active=false, framed=false, selecting;
  const key='project-table-stage';
  function draw() {
    frame=undefined;
    if (!active) return;
    const box=locate();
    if (box && box.width>0 && box.height>0) {
      Object.assign(mask.style,{display:'block',left:`${box.left-2}px`,top:`${box.top-2}px`,width:`${box.width+4}px`,height:`${box.height+4}px`});
      if(box.clip) mask.style.clipPath=`inset(${Math.max(0,box.clip.top-box.top)}px ${Math.max(0,box.left+box.width-box.clip.right)}px ${Math.max(0,box.top+box.height-box.clip.bottom)}px ${Math.max(0,box.clip.left-box.left)}px)`;
      else mask.style.clipPath='';
    } else mask.style.display='none';
    frame=scheduler.requestAnimationFrame(draw);
  }
  return {
    get active(){return active;},
    get framed(){return framed;},
    async begin() {
      if(active) throw new Error('Table stage already active');
      mask=document.createElement('div');mask.dataset.shimoAgent='table-mask';
      Object.assign(mask.style,{position:'fixed',display:'none',background:'#fff',zIndex:'2147482990',pointerEvents:'auto'});
      document.body.append(mask);active=true;draw();
    },
    async activate() {
      if (!active) throw new Error('Table stage is not active');
      if (framed) return;
      selecting ??= overlay.select(key,locate,{edge:true,writer:false,persistent:true,label:'Agent · 生成表格',selecting:'Agent · 框选中'}).then(()=>{framed=true;}).finally(()=>{selecting=undefined;});
      await selecting;
    },
    end(){
      if(!active)return;
      active=false;framed=false;if(frame!==undefined)scheduler.cancelAnimationFrame(frame);frame=undefined;
      overlay.drop(key);mask.remove();mask=undefined;
    },
  };
}
