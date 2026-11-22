// @vitest-environment happy-dom
import { afterEach, describe, expect, it } from 'vitest';
import { AiActivity } from '../src/render/activity.js';
import type { DocxViewer } from '../src/render/viewer.js';

/** Just what AiActivity touches when no document is loaded. */
function host(): { viewer: DocxViewer; root: HTMLDivElement } {
  const root = document.createElement('div');
  root.className = 'docxkeun-viewer';
  document.body.appendChild(root);
  const viewer = { root, scale: 1, document: undefined, on: () => () => {} } as unknown as DocxViewer;
  return { viewer, root };
}

const label = (root: HTMLElement) => root.querySelector('.docxkeun-ai-caret span')!.textContent;
const prop = (root: HTMLElement, name: string) => root.style.getPropertyValue(name);

afterEach(() => document.body.replaceChildren());

describe('AI activity: name and colours', () => {
  it('names the caret from the option and renames it at run time', () => {
    const { viewer, root } = host();
    const activity = new AiActivity(viewer, { name: 'Ada' });
    activity.idle();
    expect(label(root)).toBe('Ada');
    activity.name = 'Helper';
    expect(activity.name).toBe('Helper');
    expect(label(root)).toBe('Helper');
  });

  it('defaults to Claude and the purple-to-orange colours', () => {
    const { viewer, root } = host();
    const activity = new AiActivity(viewer);
    expect(activity.name).toBe('Claude');
    expect(activity.color).toBeUndefined();
    expect(prop(root, '--docxkeun-ai-color-1')).toBe('');
  });

  it('takes one colour for both ends, or two for a gradient', () => {
    const { viewer, root } = host();
    const activity = new AiActivity(viewer, { color: '#0b57d0' });
    expect(prop(root, '--docxkeun-ai-color-1')).toBe('#0b57d0');
    expect(prop(root, '--docxkeun-ai-color-2')).toBe('#0b57d0');
    activity.color = ['#0b57d0', '#34a853'];
    expect(activity.color).toEqual(['#0b57d0', '#34a853']);
    expect(prop(root, '--docxkeun-ai-color-1')).toBe('#0b57d0');
    expect(prop(root, '--docxkeun-ai-color-2')).toBe('#34a853');
  });

  it('goes back to the defaults with undefined', () => {
    const { viewer, root } = host();
    const activity = new AiActivity(viewer, { color: ['#ffeb3b', '#fff59d'] });
    activity.color = undefined;
    expect(activity.color).toBeUndefined();
    for (const name of ['--docxkeun-ai-color-1', '--docxkeun-ai-color-2', '--docxkeun-ai-text']) expect(prop(root, name)).toBe('');
  });

  it('keeps the label readable: dark text on light colours, white on dark ones', () => {
    const { viewer, root } = host();
    const activity = new AiActivity(viewer, { color: ['#ffeb3b', '#fff59d'] });
    expect(prop(root, '--docxkeun-ai-text')).toBe('#1f1f1f');
    activity.color = '#1a237e';
    expect(prop(root, '--docxkeun-ai-text')).toBe('');
    // The default purple-to-orange keeps its white label.
    activity.color = ['#7c5cff', '#d97757'];
    expect(prop(root, '--docxkeun-ai-text')).toBe('');
  });

  it('refuses something that is not a colour and keeps the previous one', () => {
    const { viewer, root } = host();
    const activity = new AiActivity(viewer, { color: '#0b57d0' });
    expect(() => (activity.color = '')).toThrow(TypeError);
    expect(() => (activity.color = ['#0b57d0', ' '])).toThrow(TypeError);
    if (typeof CSS !== 'undefined' && typeof CSS.supports === 'function' && !CSS.supports('color', 'banana')) {
      expect(() => (activity.color = 'banana')).toThrow(/not a CSS colour/);
    }
    expect(activity.color).toBe('#0b57d0');
    expect(prop(root, '--docxkeun-ai-color-1')).toBe('#0b57d0');
  });

  it('sets the name as text, never as markup', () => {
    const { viewer, root } = host();
    const activity = new AiActivity(viewer);
    activity.name = '<img src=x onerror=alert(1)>';
    expect(root.querySelector('.docxkeun-ai-caret img')).toBeNull();
    expect(label(root)).toBe('<img src=x onerror=alert(1)>');
  });
});
