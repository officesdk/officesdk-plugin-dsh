/** Meta URLs are host-relative; keep any application mount that precedes the route prefix. */
export function resolveMeta(meta, href) {
  const pathname = new URL(href).pathname;
  const marker = pathname.lastIndexOf(meta.routePrefix + '/doc/');
  if (marker < 0) throw new Error('Preview route prefix mismatch');
  const mount = pathname.slice(0, marker);
  const resolved = { ...meta };
  for (const key of ['contentUrl', 'resourcesBaseUrl', 'runtimeBaseUrl', 'sdkApiBaseUrl', 'stageUrl']) {
    if (meta[key] === undefined) continue;
    const url = new URL(mount + meta[key], href);
    if (url.origin !== new URL(href).origin) throw new Error('External preview route blocked');
    resolved[key] = url.href;
  }
  return resolved;
}
