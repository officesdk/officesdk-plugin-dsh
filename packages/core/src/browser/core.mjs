/** Convert only embedded Office resources to authenticated local lease routes. */
export function resourceUrl(input,meta) {
  if (typeof input!=='string') throw new Error('Invalid imported resource URL');
  const base=new URL(meta.resourcesBaseUrl);
  if (/^data:image\/(?:png|jpeg|gif|webp|svg\+xml);/i.test(input)) return input;
  if (input.startsWith(base.href)) return input;
  let relative=input;
  // Presentation resolves MOP URIs against the page origin before invoking its interceptor.
  if (/^[a-z][a-z0-9+.-]*:\/\//i.test(relative)) {
    const absolute=new URL(relative);
    if(absolute.protocol!==base.protocol || absolute.host!==base.host || absolute.username || absolute.password || absolute.search || absolute.hash) throw new Error('Unsupported imported resource URL: '+input);
    relative=absolute.pathname.replace(/^\/+/, '');
  }
  else if (relative.startsWith('mop-asset:/') || relative.startsWith('modoc-assets:/')) relative=relative.slice(relative.indexOf(':/')+2).replace(/^\/+/, '');
  else if (relative.startsWith('image/')) relative='images/'+relative.slice(6);
  else relative=relative.replace(/^\/+/, '');
  if (!/^(?:media|images|embedding|embeddings)\//.test(relative) || /[\\%?#:\u0000-\u001f\u007f]/.test(relative)) throw new Error('Unsupported imported resource URL: '+input);
  const parts=relative.split('/');
  if (parts.some(part=>!part || part==='.' || part==='..')) throw new Error('Invalid imported resource path');
  return new URL(parts.map(encodeURIComponent).join('/'),base).href;
}
/** SDK lifetime is owned by the iframe, including creation that finishes after abort. */
export async function mountSDK({create,root,signal,afterReady=async()=>{}}) {
  let sdk, destroyed=false;
  const destroy=()=>{ if (!sdk || destroyed) return; destroyed=true; return sdk.destroy(); };
  const abort=()=>{Promise.resolve(destroy()).catch(()=>{});};
  signal.addEventListener('abort',abort,{once:true});
  try {
    signal.throwIfAborted(); sdk=await create();
    if(signal.aborted){await destroy();signal.throwIfAborted();}
    await sdk.init(); signal.throwIfAborted(); await sdk.mount(root);signal.throwIfAborted();await sdk.ready();signal.throwIfAborted();
    await afterReady(sdk);signal.throwIfAborted();return sdk;
  } catch(error) {await destroy();signal.removeEventListener('abort',abort);throw error;}
}
export function loadScript(src) {
  return new Promise((resolve,reject)=>{const script=document.createElement('script');script.src=src;script.async=false;script.onload=()=>resolve();script.onerror=()=>reject(new Error('Runtime script failed: '+src));document.head.appendChild(script);});
}
export function loadStyle(href) {
  return new Promise((resolve,reject)=>{const link=document.createElement('link');link.rel='stylesheet';link.href=href;link.onload=()=>resolve();link.onerror=()=>reject(new Error('Runtime stylesheet failed: '+href));document.head.appendChild(link);});
}
export async function loadManifest(base,entry=false) {
  const response=await fetch(new URL('manifest.json',base));if(!response.ok)throw new Error('SDK manifest unavailable: '+response.status);
  const manifest=await response.json();const assets=entry?manifest.entry:manifest;
  if(!assets?.scripts?.length)throw new Error('Invalid SDK runtime manifest');
  await Promise.all((assets.styles||[]).map(path=>loadStyle(new URL(path,base))));
  for(const path of assets.scripts) {await loadScript(new URL(path,base));window._L?.();}
}
