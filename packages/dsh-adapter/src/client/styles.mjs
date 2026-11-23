/**
 * The P area's stylesheet. Colors follow DSH's effective theme: DSH marks `body[data-ds-dark-theme]` from its theme
 * snapshot (its own "system" preference already resolved), so a theme switch only swaps CSS variables — nothing
 * re-renders and dirty state, visibility, focus and open popovers stay. Only when no DSH theme presenter runs
 * (`html` lacks `data-ds-theme-source`) does the system scheme decide. Values: design package v2.9 docs/01 §2.
 * Popovers and dialogs stay white with dark text in both themes (MENU-01).
 */
const LIGHT = '--sbp-icon:#3c4046;--sbp-disabled:#a6aab0;--sbp-hover:#0000000a;--sbp-pressed:#00000014;--sbp-open:#0000000f;--sbp-divider:#00000026;'
  + '--sbp-focus:#0067c0;--sbp-dirty:#9a6400;--sbp-spin-track:#00000026;--sbp-spin-head:#3c4046;--sbp-badge-ring:#ffffff;'
  + '--sbp-ok:#286243;--sbp-warn:#9a6400;--sbp-error:#c02d2e;--sbp-muted:#6b7178;--sbp-saved-text:#286243;--sbp-saved-bg:#e8f2eb;';
const DARK = '--sbp-icon:#d5d8dc;--sbp-disabled:#ffffff42;--sbp-hover:#ffffff18;--sbp-pressed:#ffffff30;--sbp-open:#ffffff18;--sbp-divider:#ffffff35;'
  + '--sbp-focus:#78b6f6;--sbp-dirty:#ffdb81;--sbp-spin-track:#ffffff40;--sbp-spin-head:#e4effc;--sbp-badge-ring:#262b31;'
  + '--sbp-ok:#9fd8b4;--sbp-warn:#ffdb81;--sbp-error:#ff9b94;--sbp-muted:#a9afb8;--sbp-saved-text:#e1efff;--sbp-saved-bg:#262b31;';
const FONT = '-apple-system,BlinkMacSystemFont,"PingFang SC","Segoe UI",sans-serif';

export const CSS = `
.sbp{${LIGHT}position:relative;display:flex;align-items:center;gap:4px;white-space:nowrap;color:var(--sbp-icon);font-size:12px}
[data-ds-dark-theme] .sbp{${DARK}}
@media (prefers-color-scheme:dark){:root:not([data-ds-theme-source]) body:not([data-ds-dark-theme]) .sbp{${DARK}}}
/* White surfaces keep the light tokens whatever the theme. */
.sbp .sbp-pop,.sbp .sbp-light{${LIGHT}}
.sbp-status{display:flex;align-items:center;gap:4px;min-width:0}
.sbp-status:empty,.sbp-status:empty+.sbp-sep{display:none}
.sbp-text{min-width:0;max-width:240px;flex-shrink:1;overflow:hidden;text-overflow:ellipsis}
.sbp-ok{color:var(--sbp-ok)}.sbp-warn{color:var(--sbp-warn)}.sbp-error{color:var(--sbp-error)}.sbp-muted{color:var(--sbp-muted)}
.sbp-btn{height:24px;padding:0 8px;flex-shrink:0;font:12px/22px ${FONT};border:1px solid var(--sbp-divider);border-radius:5px;background:transparent;color:inherit;cursor:pointer;box-sizing:border-box}
.sbp-btn:hover:not(:disabled){background:var(--sbp-hover)}
.sbp-btn:disabled{color:var(--sbp-disabled);cursor:default}
.sbp-primary,.sbp-primary:hover:not(:disabled){background:#2f6bff;border-color:#2f6bff;color:#fff}
.sbp-sep{width:1px;height:16px;margin:0 3px;background:var(--sbp-divider);flex-shrink:0}
.sbp-bar{display:flex;align-items:center;gap:1px;flex-shrink:0}
.sbp-group{display:flex;align-items:center;gap:1px}
.sbp-icon{position:relative;display:inline-flex;align-items:center;justify-content:center;width:25px;height:24px;padding:4px;flex-shrink:0;border:0;border-radius:5px;background:transparent;color:var(--sbp-icon);cursor:pointer;box-sizing:border-box}
.sbp-icon svg{width:16px;height:16px;stroke-width:1.35}
.sbp-icon[data-more] svg{stroke-width:3.2}
.sbp-icon:hover:not([aria-disabled=true]){background:var(--sbp-hover)}
.sbp-icon:active:not([aria-disabled=true]),.sbp-icon[aria-pressed=true]{background:var(--sbp-pressed)}
.sbp-icon[aria-expanded=true]{background:var(--sbp-open)}
.sbp-icon[aria-disabled=true]{color:var(--sbp-disabled);cursor:default}
.sbp-icon[aria-pressed=true][aria-disabled=true]{color:var(--sbp-icon)}
.sbp :focus-visible{outline:2px solid var(--sbp-focus);outline-offset:-1px}
html[data-input-modality=pointer] .sbp :focus-visible:not(:read-write){outline-color:transparent}
.sbp-badge{position:absolute;top:0;right:0;width:11px;height:11px;display:grid;place-items:center;border-radius:50%;background:#e7a444;color:#3b2609;font:700 9px/11px ${FONT};box-shadow:0 0 0 1px var(--sbp-badge-ring);pointer-events:none}
.sbp-dot{position:absolute;top:3px;right:3px;width:6px;height:6px;border-radius:50%;background:var(--sbp-dirty);pointer-events:none}
.sbp-dot[data-error]{background:var(--sbp-error)}
.sbp-icon[data-look=saved],.sbp-icon[data-look=saved]:hover{color:var(--sbp-saved-text);background:var(--sbp-saved-bg)}
.sbp-spin{display:block;width:13px;height:13px;box-sizing:border-box;border:1.5px solid var(--sbp-spin-track);border-top-color:var(--sbp-spin-head);border-radius:50%;animation:sbp-spin .7s linear infinite}
@keyframes sbp-spin{to{transform:rotate(360deg)}}
@media (prefers-reduced-motion:reduce){.sbp-spin{animation:none}}
.sbp-pop{position:fixed;z-index:1000;box-sizing:border-box;overflow:auto;background:#fff;border:1px solid #c9c9cc;border-radius:10px;box-shadow:0 8px 22px #00000024,0 1px 4px #00000016;color:#242424;font:400 13px/20px ${FONT};white-space:normal;text-align:left}
.sbp .sbp-pop :focus-visible{outline:2px solid #0067c0;outline-offset:-1px}
.sbp-pop .sbp-btn{border-color:#c9c9cc;color:#242424;font-size:13px}
.sbp-pop .sbp-btn:hover:not(:disabled){background:#0000000a}
.sbp-pop .sbp-btn:disabled{color:#a6aab0}
.sbp-pop .sbp-primary,.sbp-pop .sbp-primary:hover:not(:disabled){background:#2f6bff;border-color:#2f6bff;color:#fff}
.sbp-pop .sbp-primary:disabled{background:#a9c0f5;border-color:#a9c0f5;color:#fff}
.sbp-menu{width:max-content;padding:6px 5px 8px}
.sbp-menu-title{padding:0 24px 3px;color:#a9afb8;font-size:12px;line-height:24px;white-space:nowrap}
.sbp-item{display:flex;align-items:center;gap:2px;width:100%;min-height:24px;padding:2px 6px;margin:0;border:0;border-radius:4px;background:transparent;color:inherit;font:inherit;text-align:left;white-space:nowrap;cursor:pointer;outline:none}
.sbp-item:hover,.sbp-item:focus{background:#007aff;color:#fff}
.sbp-item:active{background:#0069df}
.sbp .sbp-pop .sbp-item:focus-visible{outline:none}
.sbp-check{flex:0 0 16px;width:16px;text-align:center;font-weight:500;visibility:hidden}
.sbp-item[aria-checked=true] .sbp-check{visibility:visible}
.sbp-dialog{display:grid;gap:8px;width:320px;padding:12px}
/* Update bubble (design package V3.0 P区): white in both themes, 306px, version/date left, black Update top right. It
   opens downward by revealing from the top (clip-path, no transform: the popover is placed from its measured box). */
.sbp-pop.sbp-update{width:306px;padding:18px;border:1px solid #d9d9dc;border-radius:14px;background:#fff;box-shadow:0 12px 32px #00000024,0 2px 5px #0000000d;color:#1d1d1f;font:400 13px/20px ${FONT};outline:none}
.sbp-update-opening{animation:sbp-update-open .24s cubic-bezier(.2,.8,.2,1) both}
@keyframes sbp-update-open{from{opacity:0;clip-path:inset(0 0 75% 0 round 14px)}to{opacity:1;clip-path:inset(0 round 14px)}}
@media (prefers-reduced-motion:reduce){.sbp-update-opening{animation:none}}
.sbp-update-head{display:flex;align-items:flex-start;justify-content:space-between;gap:12px}
.sbp-update-meta{flex:1;min-width:0}
.sbp-update-version{overflow-wrap:anywhere;font-size:18px;line-height:24px;font-weight:600;letter-spacing:-.3px}
.sbp-update-date{display:block;margin-top:4px;color:#86868b;font-size:12px;line-height:18px}
.sbp .sbp-pop .sbp-update-primary.sbp-primary{flex:none;display:inline-flex;align-items:center;justify-content:center;gap:6px;box-sizing:border-box;min-width:72px;height:30px;padding:0 14px;border:0;border-radius:7px;background:#202124;color:#fff;font:500 13px/30px ${FONT};cursor:pointer}
.sbp .sbp-pop .sbp-update-primary.sbp-primary:hover:not(:disabled){background:#08090a}
.sbp .sbp-pop .sbp-update-primary.sbp-primary:disabled{background:#ededf0;color:#96969c;cursor:default}
.sbp-update-primary .sbp-spin{--sbp-spin-track:#00000026;--sbp-spin-head:#96969c}
.sbp-update-notes{margin-top:16px;padding-top:16px;border-top:1px solid #eeeef0}
.sbp-update-notes h3{margin:0 0 8px;font-size:13px;line-height:20px;font-weight:600}
.sbp-update-notes ul{margin:0;padding-left:17px;color:#424245}
.sbp-update-notes li{margin:0 0 7px;padding-left:1px}
.sbp-update-notes li:last-child{margin-bottom:0}
.sbp-update-fallback{margin:0;color:#424245}
.sbp-update-message{margin:16px 0 0;font-size:12px;line-height:18px;color:#6e6e73}
.sbp-update-warning{color:#956021}
.sbp-whats-new{width:300px}
.sbp-notes{display:grid;gap:6px;max-height:260px;overflow:auto}
.sbp-notes ul{margin:0;padding-left:16px;display:grid;gap:4px}
.sbp-notes li{line-height:18px}
.sbp-save-as{width:377px}
.sbp-account{width:196px}
.sbp-account-row svg{flex:0 0 16px;width:16px;height:16px;color:#62666b;stroke-width:1.6}
.sbp-account-row:hover svg,.sbp-account-row:focus svg{color:#fff}
.sbp-account-tall{align-items:flex-start;min-height:44px;border-bottom:1px solid #e1e5eb;border-radius:4px 4px 0 0;margin-bottom:3px}
.sbp-account-tall svg{margin-top:2px}
.sbp-account-text{display:flex;flex-direction:column;min-width:0}
.sbp-account-detail{color:#969ba2;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;max-width:150px}
.sbp-account-value{margin-left:auto;color:#969ba2}
.sbp-account-link{margin-left:auto;display:inline-flex}
.sbp-account-link svg{width:13px;height:13px;flex-basis:13px;color:#969ba2}
.sbp-account-row[aria-disabled=true]{cursor:default}
.sbp-account-demo{padding:0 6px 2px;text-align:right;color:#b07a1e;line-height:18px}
.sbp-account-section{padding:6px 6px 0;border-top:1px solid #e1e5eb;margin-top:3px}
.sbp-account-row[aria-checked=true] .sbp-check{visibility:visible}
.sbp-browse{width:320px}
.sbp-field{display:grid;grid-template-columns:72px 1fr auto;align-items:center;gap:6px}
.sbp-field input{min-width:0;height:24px;padding:0 6px;border:1px solid #c9c9cc;border-radius:5px;font:inherit;color:inherit;background:#fff}
.sbp-field input[aria-invalid=true]{border-color:#c02d2e}
.sbp-where{min-width:0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.sbp-problem{color:#c02d2e}
.sbp-list{display:grid;max-height:240px;overflow:auto}
.sbp-item[aria-pressed=true]{background:#dbe5f1}
.sbp-dialog-title{display:flex;justify-content:space-between;align-items:center;font-weight:600}
.sbp-tag{padding:1px 5px;border-radius:4px;background:#fff4df;color:#916022;font-size:10px;font-weight:400;line-height:16px}
.sbp-row{display:flex;justify-content:space-between;color:#969ba2}
.sbp-actions{display:flex;justify-content:flex-end;gap:6px}
.sbp-strong{font-weight:600}.sbp-quiet{color:#767d86}
/* Chat rows of the plugin's presentation tools (tool-views.mjs): DSH's row metrics and theme variables, own classes only. */
.sbt{display:flex;flex-direction:column;min-width:0;container-type:inline-size}
.sbt-line{display:flex;align-items:center;gap:6px;min-width:0}
.sbt-row{display:flex;align-items:center;flex:0 1 auto;min-width:0;height:calc(24px + var(--dsh-content-font-delta,0px));overflow:hidden;color:var(--dsw-alias-label-tertiary,#8a8f98);font-size:var(--dsh-content-font-size-secondary,13px);line-height:calc(24px + var(--dsh-content-font-delta,0px));white-space:nowrap;transition:color .1s}
.sbt-row:hover{color:var(--dsw-alias-label-secondary,#5c6068)}
.sbt-row[data-expandable]{cursor:pointer}
.sbt-row[data-busy] .sbt-title{animation:sbt-pulse 1.6s ease-in-out infinite}
@keyframes sbt-pulse{50%{opacity:.5}}
.sbt-lead{display:inline-flex;align-items:center;justify-content:center;flex:none;width:calc(16px + var(--dsh-content-font-delta,0px));margin-right:6px}
.sbt-lead svg{width:calc(14px + var(--dsh-content-font-delta,0px));height:calc(14px + var(--dsh-content-font-delta,0px))}
.sbt-title{flex:none}
.sbt-sep{flex:none;width:2px;height:2px;margin:0 8px;border-radius:1px;background:var(--dsw-alias-label-caption,#b0b4ba)}
.sbt-part{flex:0 1 auto;min-width:0;overflow:hidden;text-overflow:ellipsis}
.sbt-pages{flex-shrink:0}
.sbt-file{flex-shrink:1;max-width:40%}
.sbt-error{color:var(--dsw-alias-state-error-primary,#c02d2e)}
.sbt-warn{color:var(--dsw-alias-state-warn-label,#9a6400)}
.sbt-open{display:inline-flex;align-items:center;gap:4px;flex:none;height:22px;padding:0 8px;border:.5px solid var(--dsw-alias-border-l4,#d0d3d8);border-radius:999px;background:transparent;color:var(--dsw-alias-label-secondary,#5c6068);font:inherit;font-size:12px;line-height:20px;white-space:nowrap;cursor:pointer;opacity:.75;transition:opacity .1s,background .1s}
.sbt:hover .sbt-open,.sbt-open:focus-visible{opacity:1}
.sbt-open:hover{background:var(--dsw-alias-interactive-bg-hover-solid,#0000000a);color:var(--dsw-alias-label-primary,#222)}
.sbt-open svg{width:12px;height:12px}
/* A chat column beside the right Sidebar: the button keeps its icon (and its label for assistive tech), the row keeps its words. */
@container (max-width:640px){.sbt-open{padding:0 5px}.sbt-open span{display:none}}
.sbt-body{display:flex;flex-direction:column;gap:4px;max-height:260px;margin:4px 0 4px 22px;padding:8px 12px;overflow:auto;border:.5px solid var(--dsw-alias-border-l1,#e5e7eb);border-radius:var(--dsw-radius-lg,8px);background:var(--dsw-alias-markdown-code-block,#f6f7f9);color:var(--dsw-alias-label-secondary,#5c6068);font-size:12px;line-height:18px}
.sbt-list{display:flex;flex-direction:column;gap:2px;margin:0;padding:0 0 0 20px}
.sbt-list li{min-width:0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.sbt-list li>span+span{margin-left:8px}
.sbt-op{color:var(--dsw-alias-label-primary,#222)}
.sbt-where{color:var(--dsw-alias-label-tertiary,#8a8f98)}
.sbt-detail{margin:0;white-space:pre-wrap;overflow-wrap:anywhere;font:inherit}
.sbt-inspect{align-self:flex-start;padding:0;border:0;background:none;color:var(--dsw-alias-label-tertiary,#8a8f98);font:inherit;text-decoration:underline;cursor:pointer}
.sbt :focus-visible{outline:1.5px solid var(--dsw-focus-ring-color,var(--dsw-alias-state-business-primary,#0067c0));outline-offset:-1px}
@media (prefers-reduced-motion:reduce){.sbt-row,.sbt-open{transition:none}.sbt-row[data-busy] .sbt-title{animation:none}}
/* Input box PPT style picker (style-picker.mjs): the button follows DSH's composer colours; the popover is white like every popover here. */
.sbs{display:inline-flex;align-items:center;min-width:0}
.sbs-btn{display:inline-flex;align-items:center;gap:4px;height:28px;max-width:220px;padding:0 8px;border:0;border-radius:8px;background:transparent;color:var(--dsw-alias-label-secondary,#5c6068);font:400 13px/28px ${FONT};cursor:pointer;white-space:nowrap}
.sbs-icon{justify-content:center;width:28px;padding:0}
.sbs-btn:hover,.sbs-btn[aria-expanded=true]{background:var(--dsw-alias-fill-secondary,#0000000a);color:var(--dsw-alias-label-primary,#222)}
.sbs-btn:focus-visible,.sbs-card:focus-visible,.sbs-link:focus-visible,.sbs-x:focus-visible{outline:2px solid #0067c0;outline-offset:1px}
.sbs-head{display:flex;align-items:baseline;gap:8px}
.sbs-heading{font-weight:600;white-space:nowrap}
.sbs-hint{flex:1;min-width:0;color:#6b7178;font-size:12px}
.sbs-x{width:24px;height:24px;border:0;border-radius:6px;background:transparent;color:#6b7178;font-size:18px;line-height:22px;cursor:pointer}
.sbs-x:hover{background:#0000000a}
.sbs-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(200px,1fr));gap:10px}
.sbs-card{display:flex;flex-direction:column;gap:4px;min-width:0;padding:6px;border:1px solid #e3e4e8;border-radius:10px;background:#fff;color:inherit;font:inherit;text-align:left;cursor:pointer}
.sbs-card:hover:not(:disabled){border-color:#2f6bff;background:#f5f8ff}
.sbs-card[aria-selected=true]{border-color:#2f6bff}
.sbs-card:disabled{cursor:progress;opacity:.7}
.sbs-thumb{display:block;width:100%;aspect-ratio:16/9;object-fit:cover;border-radius:6px;background:#f0f1f3}
.sbs-swatch{display:flex;align-items:center;justify-content:center;color:#fff;font-weight:600}
.sbs-name{display:flex;align-items:baseline;gap:6px;min-width:0}
.sbs-title{font-weight:600;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.sbs-id{color:#6b7178;font-size:12px;font-family:ui-monospace,SFMono-Regular,Menlo,monospace}
.sbs-badge{margin-left:auto;padding:0 6px;border-radius:8px;background:#e8efff;color:#1d4ed8;font-size:11px;line-height:18px;white-space:nowrap}
.sbs-meta{color:#5c6068;font-size:12px;line-height:17px}
.sbs-requires{color:#9a6400}
.sbs-state{display:flex;gap:8px;align-items:center;padding:16px 4px;color:#5c6068}
.sbs-notice{font-size:12px}.sbs-error{color:#c02d2e}.sbs-ok{color:#286243}
.sbs-foot{display:flex;flex-wrap:wrap;align-items:center;gap:6px 16px;padding-top:8px;border-top:1px solid #eceef1;font-size:12px}
.sbs-check{display:inline-flex;align-items:center;gap:6px;cursor:pointer}
.sbs-current{display:inline-flex;align-items:center;gap:6px;color:#5c6068}
.sbs-link{padding:0;border:0;background:none;color:#0067c0;font:inherit;cursor:pointer}
.sbs-link:disabled{color:#a6aab0;cursor:default}
.sbs-gallery{display:flex;flex-direction:column;gap:8px;min-width:0}
.sbs-filters{display:flex;flex-wrap:wrap;gap:6px}
.sbs-chip{height:24px;padding:0 10px;border:1px solid #d9dbe0;border-radius:12px;background:#fff;color:#3c4046;font:400 12px/22px ${FONT};cursor:pointer;white-space:nowrap}
.sbs-chip:hover{background:#f5f8ff;border-color:#2f6bff}
.sbs-chip[aria-pressed=true]{background:#2f6bff;border-color:#2f6bff;color:#fff}
.sbs-chip:focus-visible,.sbs-hero:focus-visible,.sbs-action:focus-visible{outline:2px solid #0067c0;outline-offset:1px}
.sbs-previews{display:grid;grid-template-columns:repeat(4,1fr);gap:4px}
.sbs-preview{display:block;width:100%;aspect-ratio:16/9;object-fit:cover;border-radius:3px;background:#f0f1f3;border:1px solid #eceef1;box-sizing:border-box}
.sbs-scenarios{color:#3c4046}
.sbs-reason{color:#3c4046}
.sbs-hero{display:inline-flex;align-items:center;gap:4px;height:28px;padding:0 10px;border:1px solid var(--dsw-alias-separator,#00000026);border-radius:14px;background:transparent;color:var(--dsw-alias-label-secondary,#5c6068);font:400 13px/26px ${FONT};cursor:pointer;white-space:nowrap}
.sbs-hero:hover,.sbs-hero[aria-pressed=true]{background:var(--dsw-alias-fill-secondary,#0000000a);color:var(--dsw-alias-label-primary,#222)}
.sbs-dock{box-sizing:border-box;display:flex;flex-direction:column;gap:8px;max-height:min(420px,50vh);overflow:hidden;margin:0 0 8px;padding:10px 12px;background:#fff;border:1px solid #d9dbe0;border-radius:12px;box-shadow:0 2px 8px #00000014;color:#242424;font:400 13px/20px ${FONT};text-align:left}
.sbs-dock .sbs-gallery{flex:1 1 auto;min-height:0}
.sbs-dock .sbs-grid{min-height:0;overflow:auto;padding:2px}
.sbs-dock .sbs-grid{grid-template-columns:repeat(auto-fill,minmax(180px,1fr))}
.sbs-dock-suggest .sbs-grid{grid-template-columns:repeat(auto-fill,minmax(150px,1fr))}
.sbs-actions{display:flex;flex-wrap:wrap;gap:8px}
.sbs-action{height:28px;padding:0 12px;border:1px solid #d9dbe0;border-radius:8px;background:#fff;color:#242424;font:400 13px/26px ${FONT};cursor:pointer}
.sbs-action:hover:not(:disabled){background:#f5f8ff;border-color:#2f6bff}
.sbs-action:disabled{color:#a6aab0;cursor:default}
/* Inspiration library v2.0 (inspiration-library.mjs, style-picker.mjs): the input box's template prefix, then the
   full page behind the sidebar entry. Values: handoff OfficeSDK-DSH-Inspiration-v2.0 source design-overrides.css and
   template-detail.css. Colours follow DSH's theme (body[data-ds-dark-theme]); the page never changes it. */
.sbs-tag{position:relative;display:inline-flex;align-items:center;gap:2px;min-width:0;max-width:220px;height:24px;margin-left:4px;padding:0 4px 0 2px;box-sizing:border-box;border:0;background:transparent;color:#245dcc;font:400 14px/24px ${FONT};white-space:nowrap}
[data-ds-dark-theme] .sbs-tag{color:#98baff}
@media (prefers-color-scheme:dark){:root:not([data-ds-theme-source]) body:not([data-ds-dark-theme]) .sbs-tag{color:#98baff}}
.sbs-tag .sbs-bulb{flex:0 0 14px;width:14px;height:14px}
.sbs-tag-name{min-width:0;overflow:hidden;text-overflow:ellipsis}
.sbs-tag-x{position:absolute;top:-6px;right:-9px;display:grid;place-items:center;width:16px;height:16px;padding:0;border:0;border-radius:50%;background:var(--dsw-alias-bg-layer-1,#fff);color:inherit;font:400 14px/16px ${FONT};cursor:pointer;opacity:0;pointer-events:none;box-shadow:0 0 0 1px var(--dsw-alias-border-l2,#0002)}
.sbs-tag:hover .sbs-tag-x,.sbs-tag:focus-within .sbs-tag-x{opacity:1;pointer-events:auto}
.sbs-tag-failed{color:var(--dsw-alias-state-error-primary,#c02d2e);font-size:12px;gap:6px;max-width:320px;padding-right:12px}
.sbs-tag-failed .sbs-link{font-size:12px}
.sbs-tag-failed .sbs-tag-x,.sbs-tag-x:focus-visible{opacity:1;pointer-events:auto}
@media (hover:none){.sbs-tag-x{position:static;opacity:1;pointer-events:auto}}
.sbi-glyph{display:inline-grid;place-items:center;line-height:0}
.sbi-glyph .sbi-bulb{display:block}
.sbi-page{--sbi-bg:radial-gradient(ellipse at 55% -12%,rgba(168,199,219,.25),transparent 58%),#eef1f4;--sbi-surface:rgba(255,255,255,.42);--sbi-raised:rgba(255,255,255,.68);--sbi-line:rgba(48,61,75,.12);--sbi-highlight:rgba(255,255,255,.8);--sbi-text:#252b32;--sbi-muted:#717a85;--sbi-hover-fade:linear-gradient(transparent,rgba(232,235,240,.9));--sbi-quick-bg:#3b3c3d;--sbi-quick-hover:#525354;--sbi-quick-text:#fff;--sbi-d-bg:#fafbfc;--sbi-d-text:#222529;--sbi-d-muted:#7d838b;--sbi-d-line:#e0e3e6;--sbi-d-chip:#eceef0;--sbi-d-cat:#626b75;--sbi-d-rule:#8a929b;--sbi-d-action-bg:#24282c;--sbi-d-action-text:#fff;--sbi-focus:#2f6bff;position:relative;display:grid;grid-template-rows:clamp(300px,40dvh,380px) minmax(240px,1fr);height:100%;min-height:0;overflow-y:auto;background:var(--sbi-bg);color:var(--sbi-text);font:400 13px/1.5 ${FONT};box-sizing:border-box}
[data-ds-dark-theme] .sbi-page,[data-ds-dark-theme] .sbi-detail,[data-ds-dark-theme] .sbi-lightbox{--sbi-bg:radial-gradient(ellipse at 55% -12%,rgba(133,170,194,.14),transparent 56%),radial-gradient(ellipse at 95% 78%,rgba(131,150,163,.05),transparent 45%),#121416;--sbi-surface:rgba(255,255,255,.045);--sbi-raised:rgba(255,255,255,.085);--sbi-line:rgba(255,255,255,.14);--sbi-highlight:rgba(255,255,255,.2);--sbi-text:#f1f2f4;--sbi-muted:#aeb2b8;--sbi-hover-fade:linear-gradient(transparent,rgba(24,25,26,.82));--sbi-quick-bg:#3b3c3d;--sbi-quick-hover:#525354;--sbi-quick-text:#fff;--sbi-d-bg:#1b1d1d;--sbi-d-text:#f4f4f4;--sbi-d-muted:#a7aaad;--sbi-d-line:#343636;--sbi-d-chip:#2a2c2d;--sbi-d-cat:#c4c7ca;--sbi-d-rule:#ddd;--sbi-d-action-bg:#ededed;--sbi-d-action-text:#191b1b;--sbi-focus:#9bbdff;}
@media (prefers-color-scheme:dark){:root:not([data-ds-theme-source]) body:not([data-ds-dark-theme]) :is(.sbi-page,.sbi-detail,.sbi-lightbox){--sbi-bg:radial-gradient(ellipse at 55% -12%,rgba(133,170,194,.14),transparent 56%),radial-gradient(ellipse at 95% 78%,rgba(131,150,163,.05),transparent 45%),#121416;--sbi-surface:rgba(255,255,255,.045);--sbi-raised:rgba(255,255,255,.085);--sbi-line:rgba(255,255,255,.14);--sbi-highlight:rgba(255,255,255,.2);--sbi-text:#f1f2f4;--sbi-muted:#aeb2b8;--sbi-hover-fade:linear-gradient(transparent,rgba(24,25,26,.82));--sbi-quick-bg:#3b3c3d;--sbi-quick-hover:#525354;--sbi-quick-text:#fff;--sbi-d-bg:#1b1d1d;--sbi-d-text:#f4f4f4;--sbi-d-muted:#a7aaad;--sbi-d-line:#343636;--sbi-d-chip:#2a2c2d;--sbi-d-cat:#c4c7ca;--sbi-d-rule:#ddd;--sbi-d-action-bg:#ededed;--sbi-d-action-text:#191b1b;--sbi-focus:#9bbdff;}}
.sbi-detail,.sbi-lightbox{--sbi-bg:radial-gradient(ellipse at 55% -12%,rgba(168,199,219,.25),transparent 58%),#eef1f4;--sbi-surface:rgba(255,255,255,.42);--sbi-raised:rgba(255,255,255,.68);--sbi-line:rgba(48,61,75,.12);--sbi-highlight:rgba(255,255,255,.8);--sbi-text:#252b32;--sbi-muted:#717a85;--sbi-hover-fade:linear-gradient(transparent,rgba(232,235,240,.9));--sbi-quick-bg:#3b3c3d;--sbi-quick-hover:#525354;--sbi-quick-text:#fff;--sbi-d-bg:#fafbfc;--sbi-d-text:#222529;--sbi-d-muted:#7d838b;--sbi-d-line:#e0e3e6;--sbi-d-chip:#eceef0;--sbi-d-cat:#626b75;--sbi-d-rule:#8a929b;--sbi-d-action-bg:#24282c;--sbi-d-action-text:#fff;--sbi-focus:#2f6bff;}
[data-ds-dark-theme] .sbi-detail,[data-ds-dark-theme] .sbi-lightbox{--sbi-bg:radial-gradient(ellipse at 55% -12%,rgba(133,170,194,.14),transparent 56%),radial-gradient(ellipse at 95% 78%,rgba(131,150,163,.05),transparent 45%),#121416;--sbi-surface:rgba(255,255,255,.045);--sbi-raised:rgba(255,255,255,.085);--sbi-line:rgba(255,255,255,.14);--sbi-highlight:rgba(255,255,255,.2);--sbi-text:#f1f2f4;--sbi-muted:#aeb2b8;--sbi-hover-fade:linear-gradient(transparent,rgba(24,25,26,.82));--sbi-quick-bg:#3b3c3d;--sbi-quick-hover:#525354;--sbi-quick-text:#fff;--sbi-d-bg:#1b1d1d;--sbi-d-text:#f4f4f4;--sbi-d-muted:#a7aaad;--sbi-d-line:#343636;--sbi-d-chip:#2a2c2d;--sbi-d-cat:#c4c7ca;--sbi-d-rule:#ddd;--sbi-d-action-bg:#ededed;--sbi-d-action-text:#191b1b;--sbi-focus:#9bbdff;}
.sbi-page *,.sbi-detail *{box-sizing:border-box}
.sbi-page button,.sbi-detail button,.sbi-lightbox button{font-family:inherit;cursor:pointer}
.sbi-page button:focus-visible,.sbi-detail button:focus-visible,.sbi-lightbox button:focus-visible{outline:2px solid var(--sbi-focus);outline-offset:3px}
.sbi-back{position:absolute;z-index:2;top:14px;left:20px;display:inline-flex;align-items:center;gap:4px;height:28px;padding:0 10px 0 6px;border:1px solid transparent;border-radius:8px;background:transparent;color:var(--sbi-muted);font-size:12px}
.sbi-back:hover{background:var(--sbi-surface);color:var(--sbi-text)}
.sbi-compose{display:flex;flex-direction:column;align-items:center;justify-content:center;gap:20px;min-width:0;padding:52px 36px 24px}
.sbi-title{width:100%;margin:0;text-align:center;font-size:26px;line-height:34px;font-weight:650;color:var(--sbi-text)}
.sbi-title-mark{display:inline-block;width:.5em;text-align:left}
.sbi-switch{position:relative;display:grid;grid-template-columns:repeat(var(--sbi-count,3),1fr);isolation:isolate;flex:0 0 auto;width:288px;height:40px;padding:4px;border:1px solid var(--sbi-line);border-radius:22px;background:var(--sbi-raised);box-shadow:inset 0 1px 0 var(--sbi-highlight),0 4px 16px rgba(0,0,0,.08);-webkit-backdrop-filter:blur(20px) saturate(140%);backdrop-filter:blur(20px) saturate(140%)}
.sbi-switch-slider{position:absolute;top:4px;bottom:4px;left:4px;width:calc((100% - 8px) / var(--sbi-count,3));transform:translateX(calc(var(--sbi-index) * 100%));transform-origin:center;border:1px solid rgba(255,255,255,.48);border-radius:18px;background:linear-gradient(125deg,rgba(255,255,255,.88),rgba(218,231,240,.8));box-shadow:inset 0 1px 0 #fff,0 2px 8px rgba(0,0,0,.12);pointer-events:none;z-index:-1}
.sbi-switch button{display:grid;place-items:center;min-width:0;padding:0;border:0;border-radius:18px;background:transparent;color:var(--sbi-muted);font:400 12px/14px ${FONT}}
.sbi-switch button[aria-checked=true]{color:#242b32}
.sbi-type{display:inline-flex;align-items:center;justify-content:center;gap:2px;flex:0 0 auto;width:max-content;height:14px;white-space:nowrap;font:400 12px/14px ${FONT}}
.sbi-type-icon{display:grid;place-items:center;width:12px;height:14px;flex:0 0 12px}
.sbi-type-icon>img{display:block;width:12px;height:12px;object-fit:contain}
.sbi-type-outline{display:block;width:12px;height:12px;background:currentColor;-webkit-mask:var(--sbi-outline) center/contain no-repeat;mask:var(--sbi-outline) center/contain no-repeat}
.sbi-type-label{display:block;font:inherit;line-height:14px}
.sbi-gallery{display:flex;flex-direction:column;min-width:0;min-height:0}
.sbi-toolbar{display:flex;align-items:center;gap:16px;flex:0 0 60px;min-width:0;padding:6px 32px 18px}
.sbi-filters{display:flex;align-items:center;gap:8px;min-width:0;overflow-x:auto;scrollbar-width:none}
.sbi-chip{flex:0 0 auto;height:28px;padding:0 12px;border:1px solid transparent;border-radius:16px;background:transparent;color:var(--sbi-muted);font:400 12px/20px ${FONT};white-space:nowrap}
.sbi-chip:hover{background:var(--sbi-surface);color:var(--sbi-text)}
.sbi-chip[aria-pressed=true]{color:var(--sbi-text);border-color:var(--sbi-line);background:var(--sbi-raised);box-shadow:inset 0 1px 0 var(--sbi-highlight)}
.sbi-scroll{flex:1 1 auto;min-height:0;overflow:auto;padding:0 32px 24px;scrollbar-width:thin;scrollbar-gutter:stable}
.sbi-grid{--sbi-gap:20px;display:grid;grid-template-columns:repeat(var(--sbi-columns,1),var(--sbi-card-width,280px));gap:var(--sbi-gap);align-items:start;justify-content:start;width:100%;padding:2px}
.sbi-column{display:flex;flex-direction:column;gap:var(--sbi-gap);min-width:0}
.sbi-card{position:relative;display:block;width:100%;min-width:0;margin:0;padding:7px;border:1px solid var(--sbi-line);border-radius:14px;background:var(--sbi-surface);color:var(--sbi-text);box-shadow:inset 0 1px 0 rgba(255,255,255,.06);transition:border-color 180ms,background 180ms}
.sbi-card:hover{background:var(--sbi-raised)}
.sbi-card-open{display:block;width:100%;padding:0;border:0;border-radius:7px;background:transparent;color:inherit;text-align:left}
.sbi-cover{position:relative;display:block;width:100%;overflow:hidden;border-radius:8px;background:#fff}
.sbi-cover>img{position:absolute;inset:0;display:block;width:100%;height:100%;object-fit:contain;opacity:0;transition:opacity 650ms ease-in-out;pointer-events:none}
.sbi-cover>img.sbi-current{opacity:1}
.sbi-cover-empty{display:flex;flex-direction:column;align-items:center;justify-content:center;gap:12px;height:100%;color:#9299a3;font-size:10px;line-height:16px}
.sbi-card-title{display:flex;align-items:center;gap:0;height:18px;margin:10px 0 4px;overflow:hidden;white-space:nowrap;font:400 12px/14px ${FONT};transition:opacity 160ms ease}
.sbi-card-divider{flex:0 0 .5px;width:.5px;height:.9em;margin-inline:8px;background:currentColor;opacity:.3}
.sbi-card-name{flex:1 1 auto;min-width:0;font:inherit;line-height:14px;overflow:hidden;white-space:nowrap;text-overflow:ellipsis}
.sbi-card-actions{position:absolute;inset:auto 7px 5px;height:42px;display:flex;align-items:center;justify-content:center;pointer-events:none;opacity:0;transition:opacity 160ms ease;background:var(--sbi-hover-fade);border-radius:0 0 7px 7px}
.sbi-quick-use{display:inline-flex;align-items:center;justify-content:center;min-height:28px;padding:4px 12px;border:0;border-radius:5px;background:var(--sbi-quick-bg);color:var(--sbi-quick-text);font:600 11px/20px ${FONT};pointer-events:none;transform:translateY(3px);transition:transform 160ms ease,background 160ms ease}
.sbi-card:hover .sbi-card-actions,.sbi-card:focus-within .sbi-card-actions{opacity:1}
.sbi-card:hover .sbi-quick-use,.sbi-card:focus-within .sbi-quick-use{transform:none;pointer-events:auto}
.sbi-card:hover .sbi-card-title,.sbi-card:focus-within .sbi-card-title{opacity:.34}
.sbi-quick-use:hover{background:var(--sbi-quick-hover)}
.sbi-state,.sbi-empty{padding:8px 2px 16px;color:var(--sbi-muted);font-size:12px}
.sbi-empty{padding:60px 16px;text-align:center}
.sbi-link{padding:0;border:0;background:transparent;color:inherit;font:inherit;text-decoration:underline;text-underline-offset:3px;cursor:pointer}
.sbi-create-area{width:min(100%,340px);height:260px;padding:2px}
.sbi-create-wrap{height:100%}
.sbi-create{display:flex;flex-direction:column;align-items:center;justify-content:center;width:100%;height:100%;padding:18px;border:1px dashed var(--sbi-line);border-radius:14px;background:var(--sbi-surface);color:var(--sbi-text);text-align:center}
.sbi-create:hover,.sbi-create.sbi-drag{background:var(--sbi-raised)}
.sbi-create-plus{display:grid;place-items:center;width:40px;height:40px;margin-bottom:10px;border:1px solid var(--sbi-line);border-radius:8px;color:var(--sbi-muted)}
.sbi-create-title{font-size:16px;font-weight:550;line-height:24px}
.sbi-create-hint,.sbi-create-note{font-size:12px;line-height:18px;color:var(--sbi-muted)}
.sbi-create-offline{margin-top:8px;font-size:11px;line-height:18px;color:var(--sbi-muted)}
.sbi-create-error{margin-top:8px;font-size:11px;line-height:18px;color:var(--dsw-alias-state-error-primary,#c02d2e)}
.sbi-create-file{font-size:13px;font-weight:550;line-height:20px;overflow-wrap:anywhere}
.sbi-foot{flex:0 0 36px;display:flex;align-items:center;justify-content:space-between;gap:10px;min-width:0;padding:0 32px;color:var(--sbi-muted);font-size:12px}
.sbi-check{display:inline-flex;align-items:center;gap:7px;flex-shrink:0;white-space:nowrap;cursor:pointer}
.sbi-check input{width:13px;height:13px;margin:0}
.sbi-check.sbi-off{opacity:.45;cursor:default}
.sbi-status{min-width:0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;text-align:right}
.sbi-status.sbi-error{color:var(--dsw-alias-state-error-primary,#c02d2e)}
.sbi-status.sbi-ok{color:var(--dsw-alias-state-success-primary,#286243)}
.sbi-detail{position:fixed;inset:0;width:min(770px,calc(100vw - 48px));max-width:none;height:min(720px,calc(100dvh - 64px));max-height:none;margin:auto;padding:0;border:1px solid var(--sbi-d-line);border-radius:12px;background:var(--sbi-d-bg);color:var(--sbi-d-text);box-shadow:0 30px 90px #0005;overflow:hidden;font:400 13px/1.75 ${FONT};animation:sbi-appear 160ms ease-out}
.sbi-detail::backdrop{background:rgba(0,0,0,.76)}
.sbi-detail-layout{position:relative;display:grid;grid-template-rows:minmax(0,1fr) auto;height:100%;min-height:0}
.sbi-detail-close{position:absolute;z-index:3;top:10px;right:12px;display:grid;place-items:center;width:28px;height:28px;padding:0;border:0;border-radius:5px;background:transparent;color:#979b9f}
.sbi-detail-close:hover{color:var(--sbi-d-text);background:var(--sbi-d-chip)}
.sbi-detail-scroll{min-height:0;overflow-y:auto;overscroll-behavior:contain;padding:46px 25px 28px;scrollbar-width:thin}
.sbi-detail-intro{display:grid;grid-template-columns:minmax(0,1fr) minmax(0,334px);gap:24px;align-items:start;margin-bottom:26px}
.sbi-detail-copy{display:flex;flex-direction:column;align-items:stretch;min-width:0;align-self:start}
.sbi-detail-category{display:block;align-self:flex-start;padding:2px 9px;border-radius:4px;background:var(--sbi-d-chip);color:var(--sbi-d-cat);font-size:10px;line-height:17px;letter-spacing:.45px;font-weight:650;text-transform:uppercase}
.sbi-detail h2{margin:12px 0 10px;font-size:30px;line-height:1.14;font-weight:720;letter-spacing:-.6px;overflow-wrap:anywhere}
.sbi-summary-area{position:relative;min-width:0}
.sbi-summary{position:relative;margin:0;padding:0 0 0 11px;border-left:2px solid transparent;font-size:13px;line-height:1.55;font-style:italic;font-weight:500}
.sbi-summary::before{content:'';position:absolute;inset:0 auto 0 0;width:2px;background:var(--sbi-d-rule);pointer-events:none}
.sbi-summary.sbi-clamped{display:-webkit-box;-webkit-box-orient:vertical;-webkit-line-clamp:var(--sbi-lines,3);overflow:hidden}
.sbi-summary-measure{position:absolute;inset:0 0 auto;height:0;overflow:hidden;visibility:hidden;pointer-events:none}
.sbi-summary-more{display:block;margin:4px 0 0 13px;padding:0;height:20px;border:0;border-radius:2px;background:transparent;color:var(--sbi-d-muted);font:500 12px/20px ${FONT}}
.sbi-summary-more:hover{color:var(--sbi-d-text);text-decoration:underline;text-underline-offset:3px}
.sbi-carousel{position:relative;min-width:0}
.sbi-carousel-image{display:flex;align-items:center;justify-content:center;width:100%;aspect-ratio:var(--sbi-ratio);margin:0;padding:0;border:0;border-radius:7px;overflow:hidden;background:transparent;cursor:zoom-in}
.sbi-carousel-image>img{display:block;width:100%;height:100%;object-fit:contain}
.sbi-carousel[data-kind=word]{width:min(100%,210px);justify-self:center}
.sbi-carousel-arrow{position:absolute;z-index:1;top:50%;display:grid;place-items:center;width:30px;height:30px;margin-top:-15px;padding:0;border:1px solid #777a7d;border-radius:50%;background:rgba(255,255,255,.8);color:#383b3d}
.sbi-carousel-arrow.sbi-prev{left:5px}.sbi-carousel-arrow.sbi-next{right:5px}
.sbi-carousel-arrow:hover:not(:disabled){background:#fff}
.sbi-carousel-arrow:disabled{opacity:.35;cursor:default}
.sbi-carousel-counter{position:absolute;right:5px;bottom:6px;padding:1px 7px;border-radius:12px;background:#3e4142b3;color:#fff;font-size:10px;line-height:16px;pointer-events:none}
.sbi-no-preview{display:flex;flex-direction:column;align-items:center;justify-content:center;gap:10px;width:100%;height:100%;background:#fff;color:#969da5;font-size:12px}
.sbi-feature{margin:0 0 24px}
.sbi-feature>img{display:block;width:100%;height:auto}
.sbi-feature[data-kind=word]{max-width:420px;margin-inline:auto}
.sbi-pages{margin-top:34px}
.sbi-pages-heading{display:flex;align-items:center;justify-content:space-between;gap:16px;padding-bottom:12px;margin-bottom:13px;border-bottom:1px solid var(--sbi-d-line)}
.sbi-pages-heading h3{margin:0;font-size:14px;line-height:20px;font-weight:650}
.sbi-pages-heading>span{color:var(--sbi-d-muted);font-size:10px;line-height:18px}
.sbi-thumbs{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:12px;align-items:start}
.sbi-thumb{display:block;position:relative;min-width:0;padding:0;border:0;border-radius:5px;overflow:hidden;background:transparent;cursor:zoom-in}
.sbi-thumb>img{display:block;width:100%;height:auto}
.sbi-thumb-badge{position:absolute;top:4px;left:4px;min-width:18px;padding:0 3px;border-radius:3px;background:#333a;color:#fff;font:10px/13px ui-monospace,SFMono-Regular,Menlo,monospace}
.sbi-detail-foot{display:flex;align-items:center;justify-content:flex-end;gap:10px;min-height:53px;padding:10px 16px;border-top:1px solid var(--sbi-d-line);background:var(--sbi-d-bg)}
.sbi-detail-use{display:inline-flex;align-items:center;justify-content:center;flex:0 0 auto;height:34px;padding:0 12px;border:1px solid transparent;border-radius:6px;background:var(--sbi-d-action-bg);color:var(--sbi-d-action-text);font-size:12px;font-weight:550;line-height:20px}
.sbi-detail-use:hover{filter:brightness(.94)}
.sbi-lightbox{position:fixed;inset:0;width:100vw;height:100dvh;max-width:none;max-height:none;margin:0;padding:60px 80px 80px;border:0;background:transparent;overflow:hidden;color:#fff}
.sbi-lightbox[open]{display:grid;place-items:center}
.sbi-lightbox::backdrop{background:rgba(0,0,0,.94)}
.sbi-lightbox-image{display:flex;align-items:center;justify-content:center;width:min(1120px,100%);height:100%;min-height:0;pointer-events:none}
.sbi-lightbox-image>img{display:block;max-width:100%;max-height:100%;width:auto;height:auto;object-fit:contain;pointer-events:auto}
.sbi-lightbox-close,.sbi-lightbox-arrow{position:absolute;display:grid;place-items:center;padding:0;border:0;border-radius:50%;color:#fff;background:#3339}
.sbi-lightbox-close{top:20px;right:20px;width:36px;height:36px}
.sbi-lightbox-arrow{top:calc(50% - 22px);width:44px;height:44px}
.sbi-lightbox-arrow.sbi-prev{left:22px}.sbi-lightbox-arrow.sbi-next{right:22px}
.sbi-lightbox-arrow[hidden]{display:none}
.sbi-lightbox-close:hover,.sbi-lightbox-arrow:hover:not(:disabled){background:#555b}
.sbi-lightbox-arrow:disabled{opacity:.2;cursor:default}
.sbi-lightbox-counter{position:absolute;bottom:24px;left:50%;transform:translateX(-50%);padding:5px 14px;border-radius:18px;background:#2b2d2ed9;color:#fff;font:13px/22px ${FONT}}
@keyframes sbi-appear{from{opacity:0;transform:translateY(5px) scale(.985)}to{opacity:1;transform:none}}
@media (hover:none){.sbi-card-actions{position:static;height:34px;margin-top:3px;opacity:1;background:none}.sbi-quick-use{pointer-events:auto;transform:none}}
@media (max-width:900px){.sbi-compose{padding-inline:24px}.sbi-toolbar{padding-inline:24px;gap:8px}.sbi-scroll,.sbi-foot{padding-inline:24px}.sbi-grid{--sbi-gap:16px}}
@media (max-width:700px){.sbi-title{font-size:23px}.sbi-detail-intro{grid-template-columns:minmax(0,1fr) minmax(0,46%);gap:18px}.sbi-detail h2{font-size:26px}}
@media (max-width:560px){.sbi-detail{width:calc(100vw - 24px);height:calc(100dvh - 32px)}.sbi-detail-scroll{padding:40px 18px 24px}.sbi-detail-intro{grid-template-columns:1fr;gap:22px}.sbi-carousel{width:100%}.sbi-carousel[data-kind=word]{width:180px}.sbi-thumbs{grid-template-columns:repeat(2,minmax(0,1fr))}.sbi-lightbox{padding:65px 46px 76px}.sbi-lightbox-arrow{width:32px;height:32px}.sbi-lightbox-arrow.sbi-prev{left:6px}.sbi-lightbox-arrow.sbi-next{right:6px}}
@media (prefers-reduced-motion:reduce){.sbi-switch-slider,.sbi-cover>img,.sbi-card,.sbi-card-actions,.sbi-quick-use,.sbi-card-title{transition:none}.sbi-detail{animation:none}}
@media (prefers-reduced-transparency:reduce){.sbi-page{--sbi-surface:#f8fafc;--sbi-raised:#fff}[data-ds-dark-theme] .sbi-page{--sbi-surface:#23272c;--sbi-raised:#30353b}.sbi-switch{-webkit-backdrop-filter:none;backdrop-filter:none}}
@supports not ((backdrop-filter:blur(1px)) or (-webkit-backdrop-filter:blur(1px))){.sbi-switch{background:#f8fafc}[data-ds-dark-theme] .sbi-switch{background:#282c31}}
`;

/**
 * Adds the stylesheet once per document. HMR can briefly run two copies of this module, so the user count lives on
 * the element itself and the last one out removes it; the newest copy's CSS wins.
 */
export function mountStyles(document) {
  if (!document?.head) return () => {};
  const element = document.head.querySelector('style[data-officesdk-p]') ?? document.createElement('style');
  element.textContent = CSS;
  element.setAttribute('data-officesdk-p', String(Number(element.getAttribute('data-officesdk-p') || 0) + 1));
  if (!element.parentNode) document.head.appendChild(element);
  let done = false;
  return () => {
    if (done) return;
    done = true;
    const left = Number(element.getAttribute('data-officesdk-p') || 1) - 1;
    if (left > 0) element.setAttribute('data-officesdk-p', String(left));
    else element.remove();
  };
}
