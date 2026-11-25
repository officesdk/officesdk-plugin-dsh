// Used in both the editor iframe and the host document, which have separate stylesheets.
export const PROMPT_CSS = `
.office-prompt{box-sizing:border-box;display:flex;align-items:center;gap:8px;min-width:0;height:48px;padding:6px 6px 6px 17px;border:1px solid #d9dde1;border-radius:9999px;corner-shape:round;background:#fff;box-shadow:0 3px 12px #1b232c12,0 1px 3px #1b232c08}
.office-prompt textarea{display:block;flex:1;min-width:0;height:20px;min-height:0;box-sizing:border-box;margin:0;padding:0;border:0;outline:0;resize:none;overflow:auto;scrollbar-width:none;background:transparent;color:#222529;font:400 14px/20px -apple-system,BlinkMacSystemFont,"PingFang SC",sans-serif}
.office-prompt textarea::-webkit-scrollbar{display:none;width:0;height:0}
.office-prompt textarea:placeholder-shown{text-overflow:ellipsis}
.office-prompt textarea::placeholder{color:#818790}
.office-prompt button.office-prompt-send{display:grid;flex:none;place-items:center;box-sizing:border-box;width:34px;height:34px;margin:0;padding:0;border:0;border-radius:50%;corner-shape:round;background:#151618;color:white;cursor:pointer}
.office-prompt button.office-prompt-send:disabled{background:#e3e5e7;color:#999fa6;cursor:default}
.office-prompt-send svg{display:block;width:21px;height:21px}
`;
