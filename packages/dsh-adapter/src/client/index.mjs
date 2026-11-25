import { install, inject } from './plugin.mjs';
export { inject };

export function apply(ctx) { install(ctx, require('react'), { createPortal: require('react-dom').createPortal }); }
