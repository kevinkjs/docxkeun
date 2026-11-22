/**
 * Shared look of the viewer's chrome (toolbar, navigation, comments, menu). Apps theme it by
 * setting the public `--docxkeun-ui-*` variables (bg, fg, muted, line, hover, accent,
 * accent-soft, radius, shadow, font) on any ancestor, e.g. the container or `:root`. The library
 * never declares those; it reads them into internal `--dk-ui-*` variables that components use.
 */
const STYLE_ID = 'docxkeun-ui-style';
const CSS = `
.docxkeun-frame,.docxkeun-viewer,.docxkeun-ui-root{--dk-ui-bg:var(--docxkeun-ui-bg,var(--docxkeun-toolbar-bg,#f6f6f4));--dk-ui-fg:var(--docxkeun-ui-fg,var(--docxkeun-toolbar-fg,#222));--dk-ui-muted:var(--docxkeun-ui-muted,#5f6368);--dk-ui-line:var(--docxkeun-ui-line,var(--docxkeun-toolbar-border,#dcdcd6));--dk-ui-hover:var(--docxkeun-ui-hover,var(--docxkeun-toolbar-hover,#e8e8e3));--dk-ui-accent:var(--docxkeun-ui-accent,#2e74b5);--dk-ui-accent-soft:var(--docxkeun-ui-accent-soft,var(--docxkeun-toolbar-active,#dde6f2));--dk-ui-radius:var(--docxkeun-ui-radius,6px);--dk-ui-shadow:var(--docxkeun-ui-shadow,0 1px 3px rgba(60,64,67,.3),0 4px 8px rgba(60,64,67,.15));--dk-ui-font:var(--docxkeun-ui-font,13px/1.4 system-ui,sans-serif)}
.docxkeun-frame{display:flex;flex-direction:column;width:100%;height:100%;min-height:0;box-sizing:border-box}
.docxkeun-toolbar-slot{flex:none}
.docxkeun-toolbar-slot:empty{display:none}
.docxkeun-body{display:flex;flex-direction:row;flex:1 1 auto;min-height:0;min-width:0}
.docxkeun-body>.docxkeun-viewer{flex:1 1 auto;min-width:0;height:auto}
.docxkeun-ui-btn{display:inline-flex;align-items:center;justify-content:center;height:28px;min-width:28px;border:1px solid transparent;border-radius:4px;background:transparent;color:var(--dk-ui-fg);font:var(--dk-ui-font);padding:0 6px;cursor:pointer}
.docxkeun-ui-btn:hover:not(:disabled){background:var(--dk-ui-hover)}
.docxkeun-ui-btn:focus-visible{outline:2px solid var(--dk-ui-accent);outline-offset:1px}
.docxkeun-ui-btn[aria-pressed=true]{background:var(--dk-ui-accent-soft)}
.docxkeun-ui-btn:disabled{opacity:.4;cursor:default}
`;

/** Inject the shared variables and base classes once. */
export function ensureUiStyle(): void {
  if (typeof document === 'undefined' || document.getElementById(STYLE_ID)) return;
  const style = document.createElement('style');
  style.id = STYLE_ID;
  style.textContent = CSS;
  document.head.appendChild(style);
}
