// @vitest-environment happy-dom
import { describe, expect, it } from 'vitest';
import { DocxViewer } from '../src/render/viewer.js';
import { ensureUiStyle } from '../src/render/ui/theme.js';
import { icon } from '../src/render/ui/icons.js';

describe('viewer frame', () => {
  it('leaves the DOM alone until a part asks for the frame', () => {
    const container = document.createElement('div');
    const viewer = new DocxViewer(container);
    expect(container.firstElementChild).toBe(viewer.root);
  });

  it('wraps the root on demand, once', () => {
    const container = document.createElement('div');
    const viewer = new DocxViewer(container);
    const a = viewer.frame();
    const b = viewer.frame();
    expect(b.frame).toBe(a.frame);
    expect(b.body).toBe(a.body);
    expect(b.toolbarSlot).toBe(a.toolbarSlot);
    expect(container.firstElementChild).toBe(a.frame);
    expect(a.frame.classList.contains('docxkeun-frame')).toBe(true);
    expect(a.frame.firstElementChild).toBe(a.toolbarSlot);
    expect(viewer.root.parentElement).toBe(a.body);
    expect(a.body.classList.contains('docxkeun-body')).toBe(true);
    expect(a.body.parentElement).toBe(a.frame);
    expect(viewer.root.className).toBe('docxkeun-viewer');
  });

  it('injects the shared style once and draws icons', () => {
    ensureUiStyle();
    ensureUiStyle();
    expect(document.querySelectorAll('#docxkeun-ui-style').length).toBe(1);
    expect(icon('fold', 'Fold')).toContain('<svg');
    expect(icon('fold', 'Fold')).toContain('aria-label="Fold"');
  });

  it('reads public --docxkeun-ui-* inputs and never declares them', () => {
    ensureUiStyle();
    const css = document.getElementById('docxkeun-ui-style')!.textContent!;
    expect(css).not.toMatch(/--docxkeun-ui-[a-z-]+\s*:/);
    const rule = css.match(/\.docxkeun-frame,\.docxkeun-viewer,\.docxkeun-ui-root\{([^}]*)\}/)![1]!;
    for (const n of ['bg', 'fg', 'muted', 'line', 'hover', 'accent', 'accent-soft', 'radius', 'shadow', 'font']) {
      expect(rule).toContain(`--dk-ui-${n}:var(--docxkeun-ui-${n},`);
    }
    expect(rule).toContain('--dk-ui-accent:var(--docxkeun-ui-accent,#2e74b5)');
  });
});
