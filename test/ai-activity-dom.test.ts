// @vitest-environment happy-dom
import { afterEach, describe, expect, it } from 'vitest';
import { DocxEditor, paragraphsOf, paragraphText } from '../src/api/index.js';
import type { Page } from '../src/layout/index.js';
import { DocxDocument } from '../src/model/index.js';
import { findAll, textContent } from '../src/xml/index.js';
import { ACTIVITY_STYLES, AiActivity, type ActivityStyle } from '../src/render/activity.js';
import type { DocxViewer } from '../src/render/viewer.js';
import { buildDocx, p } from './helpers/docx.js';
import { inline } from './helpers/drawings.js';

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

/**
 * A real document with a hand-made layout (the font files can't be read here): one line per
 * paragraph, 10 points apart, and a text item per line so word-level marks have positions.
 */
async function laidOut(texts: string[], setup?: (editor: DocxEditor) => Promise<void>, body = texts.map((t) => p(t)).join('')) {
  const doc = await DocxDocument.open(await buildDocx({ body }));
  if (setup) await setup(new DocxEditor(doc));
  const root = document.createElement('div');
  root.className = 'docxkeun-viewer';
  document.body.appendChild(root);
  const pageEl = document.createElement('div');
  root.appendChild(pageEl);
  const layout = { pages: [] as Page[] };
  const relayout = () => {
    const paras = paragraphsOf(doc.body().blocks);
    const page = { index: 0, width: 612, height: 792, content: { x: 72, y: 72, w: 468, h: 648 }, lines: [], items: [] } as unknown as Page & { lines: unknown[]; items: unknown[] };
    paras.forEach((para, i) => {
      const y = 72 + i * 20;
      page.lines.push({ paragraph: para.el, x: 72, y, w: 400, h: 14 });
      for (const seg of paragraphText(para).segments) {
        const chars = [...seg.node.children.map((c) => ('value' in c ? c.value : '')).join('')];
        page.items.push({
          type: 'text',
          y: y + 11,
          size: 11,
          font: { ascent: 0.9, descent: 0.25 },
          chars,
          xs: chars.map((_, k) => 72 + (seg.start + k) * 6),
          widths: chars.map(() => 6),
          source: { inline: { node: seg.node }, offsets: chars.map((_, k) => k) },
        });
      }
    });
    layout.pages = [page];
  };
  relayout();
  const viewer = {
    root,
    scale: 1,
    get document() {
      return doc;
    },
    get layout() {
      return layout;
    },
    pageElement: (i: number) => (i === 0 ? pageEl : undefined),
    on: () => () => {},
  } as unknown as DocxViewer;
  return { doc, root, viewer, relayout };
}

const CLASSES: Record<ActivityStyle, string> = {
  highlight: 'docxkeun-ai-shimmer',
  words: 'docxkeun-ai-words',
  underline: 'docxkeun-ai-underline',
  edge: 'docxkeun-ai-edge',
  spotlight: 'docxkeun-ai-spot',
  quiet: 'docxkeun-ai-pulse',
  sparkles: 'docxkeun-ai-sparkle',
};
const count = (root: HTMLElement, cls: string) => root.querySelectorAll(`.${cls}`).length;

/** The text items (by index on the page) a paint layer colours. */
function painted(root: HTMLElement, layer: 'work' | 'marks' | 'flash'): number[] {
  const css = root.querySelector(`style[data-docxkeun-ai-paint="${layer}"]`)!.textContent ?? '';
  return [...css.matchAll(/data-i="(\d+)"/g)].map((m) => Number(m[1]));
}

/** Where a gradient's stops in the AI's colours start and end, in page points. */
function colouredSpan(gradient: Element): [number, number] {
  const x1 = Number(gradient.getAttribute('x1'));
  const x2 = Number(gradient.getAttribute('x2'));
  const at = [...gradient.querySelectorAll('stop')].filter((s) => s.getAttribute('style')?.includes('--dk-ai')).map((s) => x1 + Number(s.getAttribute('offset')) * (x2 - x1));
  return [Math.min(...at), Math.max(...at)];
}

describe('AI activity: effect styles', () => {
  it('lists every style once, highlight first, each with a label and a line', () => {
    expect(ACTIVITY_STYLES.map((s) => s.style)).toEqual(['highlight', 'words', 'underline', 'edge', 'spotlight', 'quiet', 'sparkles']);
    for (const s of ACTIVITY_STYLES) expect(s.label && s.description).toBeTruthy();
  });

  it('draws each style while working, and only that style, with the named caret', async () => {
    const { viewer, root } = await laidOut(['First paragraph.', 'Second paragraph being changed.', 'Third.']);
    const activity = new AiActivity(viewer);
    expect(activity.style).toBe('highlight');
    activity.work([2], { scroll: false });
    for (const { style } of ACTIVITY_STYLES) {
      activity.style = style;
      for (const [other, cls] of Object.entries(CLASSES)) {
        if (other === style) expect(count(root, cls), style).toBeGreaterThan(0);
        else expect(count(root, cls), `${style} left ${other}`).toBe(0);
      }
      expect(root.querySelector('.docxkeun-ai-caret span')!.textContent).toBe('Claude is editing…');
    }
    activity.idle();
    for (const cls of Object.values(CLASSES)) expect(count(root, cls)).toBe(0);
  });

  it('shimmers the letters of the word style and lays nothing over the page', async () => {
    const { viewer, root } = await laidOut(['First paragraph.', 'Second paragraph being changed.', 'Third.']);
    const activity = new AiActivity(viewer, { style: 'words' });
    activity.work([2], { scroll: false });
    // Only the second paragraph's text (item 1) takes the gradient; no box covers the paper or shading.
    expect(painted(root, 'work')).toEqual([1]);
    expect([...root.querySelector('.docxkeun-ai-layer')!.children].every((el) => el.classList.contains('docxkeun-ai-caret'))).toBe(true);
    const defs = () => root.querySelector('svg[data-paint="work"]')!.innerHTML;
    expect(defs()).toMatch(/animateTransform/i);
    // Held still: the letters keep the colours without the sweep.
    activity.animate = false;
    expect(painted(root, 'work')).toEqual([1]);
    expect(defs()).not.toMatch(/animateTransform/i);
    activity.idle();
    expect(painted(root, 'work')).toEqual([]);
  });

  it('takes the style as an option and refuses unknown ones', async () => {
    const { viewer } = await laidOut(['Only paragraph.']);
    const activity = new AiActivity(viewer, { style: 'words' });
    expect(activity.style).toBe('words');
    expect(() => (activity.style = 'glitter' as ActivityStyle)).toThrow(/Unknown AI effect style "glitter"/);
    expect(activity.style).toBe('words');
    expect(() => new AiActivity(viewer, { style: 'nope' as ActivityStyle })).toThrow(TypeError);
  });

  it('leaves the mark of the style after a change: tint, coloured letters, underline, or the margin bar', async () => {
    const { viewer, root, doc, relayout } = await laidOut(['Payment is due within thirty days.', 'Unchanged.']);
    const activity = new AiActivity(viewer, { style: 'highlight', animate: false });
    const before = ['Payment is due within twenty days.', 'Unchanged.'];
    // The document already says "thirty": reveal compares it with the text before.
    expect(paragraphText(paragraphsOf(doc.body().blocks)[0]!).text).toContain('thirty');
    relayout();
    const diff = await activity.reveal(before, { scroll: false });
    expect(diff.changes).toHaveLength(1);
    const marks = () => ({
      tint: count(root, 'docxkeun-ai-mark-tint'),
      words: count(root, 'docxkeun-ai-mark-words'),
      underline: count(root, 'docxkeun-ai-mark-underline'),
      bar: count(root, 'docxkeun-ai-bar'),
    });
    expect(marks()).toEqual({ tint: 1, words: 0, underline: 0, bar: 0 });
    activity.style = 'words';
    expect(marks()).toEqual({ tint: 0, words: 1, underline: 0, bar: 0 });
    // Only the letters of "thirty" (characters 22-28, 6 points each from x = 72) take the colours.
    expect(painted(root, 'marks')).toEqual([0]);
    const [from, to] = colouredSpan(root.querySelector('.docxkeun-ai-mark-words')!);
    expect(from).toBeCloseTo(72 + 22 * 6, 1);
    expect(to).toBeCloseTo(72 + 28 * 6, 1);
    activity.style = 'underline';
    expect(marks()).toEqual({ tint: 0, words: 0, underline: 1, bar: 0 });
    expect(painted(root, 'marks')).toEqual([]);
    for (const style of ['edge', 'spotlight', 'quiet', 'sparkles'] as const) {
      activity.style = style;
      expect(marks(), style).toEqual({ tint: 0, words: 0, underline: 0, bar: 1 });
    }
    // The change map's tick says what was there before, whatever the style.
    expect(root.querySelector<HTMLElement>('.docxkeun-ai-map i')!.title).toBe('Changed by Claude. Before: Payment is due within twenty days.');
    activity.marksShown = false;
    expect(marks()).toEqual({ tint: 0, words: 0, underline: 0, bar: 0 });
    expect(root.querySelectorAll('.docxkeun-ai-map i')).toHaveLength(0);
  });

  it('lights up the changed letters for a moment when the word style keeps no marks', async () => {
    const t = await laidOut(['Payment is due within thirty days.']);
    const activity = new AiActivity(t.viewer, { style: 'words', marks: false });
    const before = activity.snapshot();
    new DocxEditor(t.doc).replaceAll('thirty', 'forty');
    t.relayout();
    await activity.reveal(before, { scroll: false });
    // No tinted glow over the page, no mark that stays: the letters of "forty" flash instead.
    expect(count(t.root, 'docxkeun-ai-glow')).toBe(0);
    expect(count(t.root, 'docxkeun-ai-mark-words')).toBe(0);
    expect(count(t.root, 'docxkeun-ai-flash-words')).toBeGreaterThan(0);
    expect(t.root.querySelector('.docxkeun-ai-flash')!.getAttribute('style')).toContain('--from:#000');
  });

  it('marks a paragraph with only deletions with the margin bar, even in a word style', async () => {
    const { viewer, root, relayout } = await laidOut(['Payment is due within thirty days.']);
    const activity = new AiActivity(viewer, { style: 'words', animate: false });
    relayout();
    await activity.reveal(['Payment is due within thirty calendar days.'], { scroll: false });
    expect(count(root, 'docxkeun-ai-mark-words')).toBe(0);
    expect(count(root, 'docxkeun-ai-bar')).toBe(1);
  });

  it('follows the page rows when the viewer rebuilds them (zoom)', async () => {
    const { viewer, root, relayout } = await laidOut(['Payment is due within thirty days.']);
    const activity = new AiActivity(viewer, { style: 'highlight', animate: false });
    relayout();
    await activity.reveal(['Payment is due within twenty days.'], { scroll: false });
    const tint = () => root.querySelector<HTMLElement>('.docxkeun-ai-mark-tint')!;
    const before = parseFloat(tint().style.width);
    (viewer as unknown as { scale: number }).scale = 2;
    activity.rowsBuilt();
    await new Promise((ok) => requestAnimationFrame(() => ok(undefined)));
    // The tint is the line plus 1.5 points each side, so it doubles with the scale.
    expect(parseFloat(tint().style.width)).toBeCloseTo(before * 2, 0);
  });
});

describe('AI activity: marks over a run of changes', () => {
  /** Make a change with the editor, lay out again, and reveal it as the AI's. */
  async function change(t: Awaited<ReturnType<typeof laidOut>>, activity: AiActivity, edit: (e: DocxEditor) => void) {
    const before = activity.snapshot();
    edit(new DocxEditor(t.doc));
    t.relayout();
    return activity.reveal(before, { scroll: false });
  }
  // Lines are 20 points apart from y = 72; marks sit up to 2 points above their line (padding).
  const lineOf = (el: HTMLElement) => Math.round((parseFloat(el.style.top) - 70) / 20);
  /** The lines whose letters the marks colour (text baselines are 11 points below the line's top). */
  const markedLines = (t: Awaited<ReturnType<typeof laidOut>>) => painted(t.root, 'marks').map((i) => Math.round(((t.viewer.layout!.pages[0]!.items[i] as { y: number }).y - 83) / 20));

  it('keeps the marks of every change, not only the last one', async () => {
    const t = await laidOut(['First paragraph.', 'Second paragraph.', 'Third paragraph.']);
    const activity = new AiActivity(t.viewer, { style: 'highlight' });
    await change(t, activity, (e) => e.replaceAll('First', 'Opening'));
    await change(t, activity, (e) => e.replaceAll('Third', 'Closing'));
    const tints = [...t.root.querySelectorAll<HTMLElement>('.docxkeun-ai-mark-tint')];
    expect(tints.map(lineOf)).toEqual([0, 2]);
    // Only the latest change's mark plays its entrance; the earlier one just stays.
    expect(tints.map((el) => el.classList.contains('docxkeun-ai-intro'))).toEqual([false, true]);
    expect(t.root.querySelectorAll('.docxkeun-ai-map i')).toHaveLength(2);
  });

  it('fades in only the latest change\'s coloured letters', async () => {
    const t = await laidOut(['First paragraph.', 'Second paragraph.', 'Third paragraph.']);
    const activity = new AiActivity(t.viewer, { style: 'words' });
    await change(t, activity, (e) => e.replaceAll('First', 'Opening'));
    await change(t, activity, (e) => e.replaceAll('Third', 'Closing'));
    const marks = [...t.root.querySelectorAll('.docxkeun-ai-mark-words')];
    expect(markedLines(t)).toEqual([0, 2]);
    expect(marks.map((g) => g.querySelectorAll('stop.docxkeun-ai-in').length > 0)).toEqual([false, true]);
  });

  it('moves earlier marks with their paragraphs when paragraphs are added or removed', async () => {
    const t = await laidOut(['Alpha.', 'Payment is due within thirty days.', 'Omega.']);
    const activity = new AiActivity(t.viewer, { style: 'words', animate: false });
    await change(t, activity, (e) => e.replaceAll('thirty', 'forty'));
    await change(t, activity, (e) => void e.insertParagraph('start', 'A new opening paragraph.'));
    // The new paragraph (line 0) and "forty", now on line 2.
    expect(markedLines(t).sort()).toEqual([0, 2]);
    await change(t, activity, (e) => e.deleteParagraph(paragraphsOf(t.doc.body().blocks)[0]!.el));
    expect(markedLines(t)).toEqual([1]);
  });

  it('marks a paragraph changed twice against its text before the first change', async () => {
    const t = await laidOut(['Payment is due within thirty days.']);
    const activity = new AiActivity(t.viewer, { style: 'words', animate: false });
    await change(t, activity, (e) => e.replaceAll('thirty', 'forty'));
    await change(t, activity, (e) => e.replaceAll('Payment', 'The fee'));
    // Both edits stay marked: "The fee" and "forty".
    expect(t.root.querySelectorAll('.docxkeun-ai-mark-words')).toHaveLength(2);
    expect(t.root.querySelector<HTMLElement>('.docxkeun-ai-map i')!.title).toBe('Changed by Claude. Before: Payment is due within thirty days.');
  });

  it('drops the mark when a later change puts the text back', async () => {
    const t = await laidOut(['Payment is due within thirty days.', 'Other.']);
    const activity = new AiActivity(t.viewer, { style: 'underline', animate: false });
    await change(t, activity, (e) => e.replaceAll('thirty', 'forty'));
    expect(t.root.querySelectorAll('.docxkeun-ai-mark-underline')).toHaveLength(1);
    await change(t, activity, (e) => e.replaceAll('forty', 'thirty'));
    expect(t.root.querySelectorAll('.docxkeun-ai-mark-underline')).toHaveLength(0);
    expect(t.root.querySelectorAll('.docxkeun-ai-map i')).toHaveLength(0);
  });

  it('starts over after clearMarks()', async () => {
    const t = await laidOut(['One.', 'Two.']);
    const activity = new AiActivity(t.viewer, { style: 'edge', animate: false });
    await change(t, activity, (e) => e.replaceAll('One', 'First'));
    activity.clearMarks();
    await change(t, activity, (e) => e.replaceAll('Two', 'Second'));
    expect([...t.root.querySelectorAll<HTMLElement>('.docxkeun-ai-bar')].map(lineOf)).toEqual([1]);
  });
});

describe('AI activity: changed objects', () => {
  const endOfText = (editor: DocxEditor) => {
    const t = findAll(editor.paragraphs()[0]!.el, 'w:t').at(-1)!;
    return { node: t, offset: textContent(t).length };
  };
  const PICTURE = 'http://schemas.openxmlformats.org/drawingml/2006/picture';
  const withPicture = `<w:p><w:r><w:t>Logo</w:t></w:r><w:r>${inline('', 100, 50, PICTURE)}</w:r></w:p>${p('Unchanged.')}`;

  it('marks a paragraph whose chart changed', async () => {
    const t = await laidOut(['Sales by quarter', 'Unchanged.'], async (e) => {
      await e.insertChart(endOfText(e), { type: 'column', categories: ['Q1', 'Q2'], series: [{ name: 'Sales', values: [1, 2] }] });
    });
    const activity = new AiActivity(t.viewer, { style: 'highlight', animate: false });
    const before = activity.snapshot();
    const editor = new DocxEditor(t.doc);
    await editor.updateChart(editor.objects()[0]!, { series: [{ name: 'Sales', values: [5, 9] }] });
    t.relayout();
    const diff = await activity.reveal(before, { scroll: false });
    expect(diff.changes).toMatchObject([{ para: 0, kind: 'changed', formatting: true }]);
    expect(count(t.root, 'docxkeun-ai-mark-tint')).toBe(1);
    expect(t.root.querySelectorAll('.docxkeun-ai-map i')).toHaveLength(1);
  });

  it('marks a paragraph whose picture changed in size or alt text, and none when nothing changed', async () => {
    const t = await laidOut(['Logo', 'Unchanged.'], undefined, withPicture);
    const activity = new AiActivity(t.viewer, { style: 'highlight', animate: false });
    const editor = new DocxEditor(t.doc);
    let before = activity.snapshot();
    t.relayout();
    expect((await activity.reveal(before, { scroll: false })).changes).toEqual([]);
    editor.resizeObject(editor.objects()[0]!, { width: 200 });
    t.relayout();
    expect((await activity.reveal(before, { scroll: false })).changes).toMatchObject([{ para: 0, formatting: true }]);
    before = activity.snapshot();
    editor.setObjectAlt(editor.objects()[0]!, 'A logo');
    t.relayout();
    expect((await activity.reveal(before, { scroll: false })).changes).toMatchObject([{ para: 0, formatting: true }]);
  });
});
