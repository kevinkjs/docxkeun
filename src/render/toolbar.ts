/**
 * The formatting toolbar for in-page editing: undo/redo, styles, font size, character formatting,
 * alignment, lists, text direction and Track Changes. It drives `viewer.editing`, mirrors the
 * formatting at the caret (the viewer's `caretchange` event), folds into a thin strip, and moves
 * groups that do not fit into a "more" menu. `viewer.toolbar` is its controller.
 */
import type { DocxViewer } from './viewer.js';
import type { CaretFormat } from './editing.js';
import { icon, type IconName } from './ui/icons.js';
import { fitGroups } from './ui/overflow.js';
import { positionMenu } from './ui/menu.js';
import { ensureUiStyle } from './ui/theme.js';

const STYLE_ID = 'docxkeun-toolbar-style';
const CSS = `
.docxkeun-toolbar{position:relative;background:var(--dk-ui-bg);color:var(--dk-ui-fg);border-bottom:1px solid var(--dk-ui-line);font:var(--dk-ui-font);box-sizing:border-box}
.docxkeun-toolbar[hidden]{display:none}
.docxkeun-toolbar .dk-tb-row{display:flex;flex-wrap:nowrap;align-items:center;gap:2px;padding:4px 6px;min-height:36px;box-sizing:border-box;overflow:hidden}
.docxkeun-toolbar.dk-collapsed .dk-tb-row,.docxkeun-toolbar.dk-collapsed .dk-tb-pop{display:none}
.docxkeun-toolbar .dk-tb-group{display:inline-flex;align-items:center;gap:2px;flex:none}
.docxkeun-toolbar .dk-tb-sep{width:1px;height:20px;background:var(--dk-ui-line);margin:0 4px}
.docxkeun-toolbar .dk-tb-spacer{flex:1 1 auto;min-width:0}
.docxkeun-toolbar button,.docxkeun-toolbar select{height:28px;min-width:28px;border:1px solid transparent;border-radius:4px;background:transparent;color:inherit;font:inherit;padding:0 6px;cursor:pointer;display:inline-flex;align-items:center;justify-content:center}
.docxkeun-toolbar select{border-color:var(--dk-ui-line);background:var(--dk-ui-bg);display:inline-block}
.docxkeun-toolbar button:hover:not(:disabled){background:var(--dk-ui-hover)}
.docxkeun-toolbar button:focus-visible,.docxkeun-toolbar select:focus-visible,.docxkeun-toolbar input:focus-visible+button{outline:2px solid var(--dk-ui-accent);outline-offset:1px}
.docxkeun-toolbar button[aria-pressed=true]{background:var(--dk-ui-accent-soft)}
.docxkeun-toolbar button:disabled,.docxkeun-toolbar select:disabled{opacity:.4;cursor:default}
.docxkeun-toolbar .color{position:relative;display:inline-flex;align-items:center}
.docxkeun-toolbar .color input{position:absolute;inset:0;opacity:0;cursor:pointer;width:100%;height:100%}
.docxkeun-toolbar .swatch{display:block;height:3px;margin-top:1px}
.docxkeun-toolbar .story{font-size:12px;color:var(--dk-ui-muted);padding:0 6px;white-space:nowrap}
.docxkeun-toolbar .dk-tb-more[hidden]{display:none}
.docxkeun-toolbar .dk-tb-pop{position:fixed;left:0;top:0;z-index:20;display:flex;flex-direction:column;gap:4px;padding:6px;background:var(--dk-ui-bg);border:1px solid var(--dk-ui-line);border-radius:var(--dk-ui-radius);box-shadow:var(--dk-ui-shadow)}
.docxkeun-toolbar .dk-tb-pop[hidden]{display:none}
.docxkeun-toolbar .dk-tb-pop .dk-tb-sep{display:none}
.docxkeun-toolbar .dk-tb-strip{display:none;width:100%;height:18px;min-height:0;border-radius:0;font-size:11px;color:var(--dk-ui-muted);gap:4px}
.docxkeun-toolbar.dk-collapsed .dk-tb-strip{display:flex}
.docxkeun-toolbar.dk-collapsed .dk-tb-strip[hidden]{display:none}
.docxkeun-toolbar .dk-tb-strip svg{width:12px;height:12px}
`;

function ensureStyle(): void {
  ensureUiStyle();
  if (typeof document === 'undefined' || document.getElementById(STYLE_ID)) return;
  const style = document.createElement('style');
  style.id = STYLE_ID;
  style.textContent = CSS;
  document.head.appendChild(style);
}

function opt(text: string, value: string): HTMLOptionElement {
  const o = document.createElement('option');
  o.value = value;
  o.textContent = text;
  return o;
}

const SIZES = [8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 24, 28, 36, 48, 72];
const HIGHLIGHTS: Array<[string, string]> = [
  ['none', 'No highlight'],
  ['yellow', 'Yellow'],
  ['green', 'Bright green'],
  ['cyan', 'Turquoise'],
  ['magenta', 'Pink'],
  ['lightGray', 'Gray'],
];
const STORY_NAMES: Record<string, string> = { header: 'Editing header', footer: 'Editing footer', footnote: 'Editing footnote', endnote: 'Editing endnote', textbox: 'Editing text box' };


export interface ToolbarOptions {
  /** Put the toolbar in this element instead of above the pages. */
  container?: HTMLElement;
  /** Start folded into the thin "Show toolbar" strip. */
  collapsed?: boolean;
  /** Paragraph styles offered (id → name). Default: Normal, Title, Heading 1–3 when the document has them. */
  styles?: Array<{ id: string; name: string }>;
}

/** What the toolbar needs from the viewer. */
export type ToolbarHost = Pick<DocxViewer, 'root' | 'on' | 'emit' | 'editing' | 'document' | 'frame'>;

type Control = HTMLButtonElement | HTMLSelectElement | HTMLInputElement;

/** Controls the toolbar of a viewer (`viewer.toolbar`); the DOM is built on the first `show()`. */
export class ToolbarController {
  private el: HTMLDivElement | undefined;
  private shown = false;
  private fold = false;
  private restore = false;
  private format: CaretFormat | undefined;
  private readonly controls = new Map<string, Control>();
  private readonly cleanup: Array<() => void> = [];
  private groups: HTMLElement[] = [];
  private row!: HTMLDivElement;
  private pop!: HTMLDivElement;
  private moreWrap!: HTMLSpanElement;
  private moreButton!: HTMLButtonElement;
  private foldButton!: HTMLButtonElement;
  private strip!: HTMLButtonElement;
  private storyLabel!: HTMLSpanElement;
  private resizeObserver: ResizeObserver | undefined;

  constructor(private readonly viewer: ToolbarHost, private readonly options: ToolbarOptions = {}) {
    this.fold = !!options.collapsed;
  }

  get element(): HTMLElement | undefined {
    return this.el;
  }

  get visible(): boolean {
    return this.shown;
  }

  get collapsed(): boolean {
    return this.fold;
  }

  set collapsed(value: boolean) {
    if (value === this.fold) return;
    this.fold = value;
    if (!this.el) return;
    this.applyFold();
    this.announce();
  }

  show(): void {
    if (!this.el) this.build();
    this.restore = false;
    if (this.shown) return;
    this.shown = true;
    this.el!.hidden = false;
    this.announce();
    this.relayout();
  }

  hide(): void {
    this.restore = false;
    if (!this.shown) return;
    this.shown = false;
    this.closeMenu();
    if (this.el) this.el.hidden = true;
    this.announce();
  }

  /** Show the toolbar as soon as editing turns on (used when it was asked for while read-only). */
  deferShow(): void {
    this.restore = true;
  }

  /** The viewer's `setEditable`: hide with editing, and bring the previous state back with it. */
  setEditable(on: boolean): void {
    if (on) {
      if (this.restore) this.show();
    } else {
      const was = this.shown || this.restore;
      this.hide();
      this.restore = was;
    }
  }

  private announce(): void {
    this.viewer.emit('chromechange', { part: 'toolbar', visible: this.shown, collapsed: this.fold });
  }

  private get editing() {
    return this.viewer.editing;
  }

  private build(): void {
    ensureStyle();
    const el = (this.el = document.createElement('div'));
    el.className = 'docxkeun-toolbar docxkeun-ui-root';
    el.setAttribute('role', 'toolbar');
    el.setAttribute('aria-label', 'Formatting');
    // Clicking a button must not take focus (and the selection) away from the page.
    el.addEventListener('mousedown', (e) => {
      if ((e.target as Element).closest('button')) e.preventDefault();
    });
    el.addEventListener('keydown', (e) => this.onKey(e));

    const row = (this.row = document.createElement('div'));
    row.className = 'dk-tb-row';

    const edit = () => this.editing;
    this.group('history', row, () => {
      this.button('undo', 'Undo', icon('undo'), () => edit()?.undo());
      this.button('redo', 'Redo', icon('redo'), () => edit()?.redo());
    });
    this.group('text', row, () => {
      this.select('style', 'Paragraph style', [], (v) => void edit()?.setStyle(v));
      this.select('size', 'Font size', SIZES.map((n) => [String(n), String(n)]), (v) => edit()?.format({ size: Number(v) }));
    });
    this.group('font', row, () => {
      this.button('bold', 'Bold', '<b>B</b>', () => edit()?.toggle('bold'));
      this.button('italic', 'Italic', '<i style="font-family:serif">I</i>', () => edit()?.toggle('italic'));
      this.button('underline', 'Underline', '<u>U</u>', () => edit()?.toggle('underline'));
      this.button('strike', 'Strikethrough', '<s>S</s>', () => edit()?.toggle('strike'));
      this.colorInput();
      this.select('highlight', 'Highlight', HIGHLIGHTS, (v) => edit()?.format({ highlight: v }));
    });
    this.group('align', row, () => {
      const names: Array<['left' | 'center' | 'right' | 'justify', IconName]> = [['left', 'alignLeft'], ['center', 'alignCenter'], ['right', 'alignRight'], ['justify', 'justify']];
      for (const [align, name] of names) this.button(`align-${align}`, `Align ${align}`, icon(name), () => void edit()?.formatParagraphs({ align }));
    });
    this.group('list', row, () => {
      this.button('bullets', 'Bulleted list', icon('bullets'), () => void edit()?.setList(this.format?.list === 'bullet' ? 'none' : 'bullet'));
      this.button('numbers', 'Numbered list', icon('numbers'), () => void edit()?.setList(this.format?.list === 'number' ? 'none' : 'decimal'));
      this.button('outdent', 'Decrease indent', icon('outdent'), () => void edit()?.indent(-1));
      this.button('indent', 'Increase indent', icon('indent'), () => void edit()?.indent(1));
    });
    this.group('direction', row, () => {
      this.button('ltr', 'Left-to-right text', icon('ltr'), () => void edit()?.formatParagraphs({ direction: 'ltr' }));
      this.button('rtl', 'Right-to-left text', icon('rtl'), () => void edit()?.formatParagraphs({ direction: 'rtl' }));
    });
    this.group('track', row, () => {
      this.button('track', 'Track changes', `${icon('track')}<span style="margin-left:4px">Track</span>`, () => {
        const ed = edit();
        if (ed) ed.trackChanges = !ed.trackChanges;
      });
    });

    this.moreWrap = document.createElement('span');
    this.moreWrap.className = 'dk-tb-more';
    this.moreWrap.hidden = true;
    this.moreButton = this.rawButton('Show more formatting', icon('more'), () => this.toggleMenu());
    this.moreButton.setAttribute('aria-haspopup', 'true');
    this.moreButton.setAttribute('aria-expanded', 'false');
    this.moreWrap.appendChild(this.moreButton);
    const spacer = document.createElement('span');
    spacer.className = 'dk-tb-spacer';
    this.storyLabel = document.createElement('span');
    this.storyLabel.className = 'story';
    this.storyLabel.setAttribute('aria-live', 'polite');
    this.foldButton = this.rawButton('Hide toolbar', icon('fold'), () => (this.collapsed = true));
    row.append(this.moreWrap, spacer, this.storyLabel, this.foldButton);

    this.pop = document.createElement('div');
    this.pop.className = 'dk-tb-pop';
    this.pop.setAttribute('role', 'group');
    this.pop.setAttribute('aria-label', 'More formatting');
    this.pop.hidden = true;

    this.strip = document.createElement('button');
    this.strip.type = 'button';
    this.strip.className = 'dk-tb-strip';
    this.strip.setAttribute('aria-label', 'Show toolbar');
    this.strip.innerHTML = `${icon('unfold')}<span>Show toolbar</span>`;
    this.strip.addEventListener('click', () => (this.collapsed = false));

    el.append(row, this.pop, this.strip);
    el.hidden = true;
    this.applyFold();

    const container = this.options.container;
    if (container) container.appendChild(el);
    else this.viewer.frame().toolbarSlot.appendChild(el);

    this.cleanup.push(this.viewer.on('caretchange', ({ format }) => this.update(format)));
    this.cleanup.push(this.viewer.on('layout', () => this.fillStyles()));
    const onDoc = (e: Event) => {
      if (!this.pop.hidden && !el.contains(e.target as Node)) this.closeMenu();
    };
    document.addEventListener('mousedown', onDoc);
    this.cleanup.push(() => {
      window.removeEventListener('resize', this.onWinResize);
      window.removeEventListener('scroll', this.onWinScroll, true);
    });
    this.cleanup.push(() => document.removeEventListener('mousedown', onDoc));
    if (typeof ResizeObserver !== 'undefined') {
      this.resizeObserver = new ResizeObserver(() => this.relayout());
      this.resizeObserver.observe(el);
    }
    this.fillStyles();
    this.update(this.editing?.caretFormat());
    this.roving();
  }

  private applyFold(): void {
    const el = this.el!;
    el.classList.toggle('dk-collapsed', this.fold);
    this.strip.hidden = !this.fold;
    this.foldButton.setAttribute('aria-expanded', String(!this.fold));
    this.closeMenu();
    if (!this.fold) this.relayout();
  }

  private group(name: string, parent: HTMLElement, fill: () => void): void {
    const g = document.createElement('span');
    g.className = 'dk-tb-group';
    g.dataset.group = name;
    if (this.groups.length) {
      const sep = document.createElement('span');
      sep.className = 'dk-tb-sep';
      sep.setAttribute('role', 'separator');
      g.appendChild(sep);
    }
    this.groups.push(g);
    this.target = g;
    fill();
    this.target = undefined;
    parent.appendChild(g);
  }

  private target: HTMLElement | undefined;

  private rawButton(label: string, html: string, run: () => void): HTMLButtonElement {
    const b = document.createElement('button');
    b.type = 'button';
    b.innerHTML = html;
    b.title = label;
    b.dataset.nav = '';
    b.setAttribute('aria-label', label);
    b.addEventListener('click', run);
    return b;
  }

  private button(key: string, label: string, html: string, run: () => void): void {
    const b = this.rawButton(label, html, () => {
      run();
      this.closeMenu();
      this.editing?.focus();
    });
    this.controls.set(key, b);
    this.target?.appendChild(b);
  }

  private select(key: string, label: string, options: Array<[string, string]>, run: (value: string) => void): void {
    const s = document.createElement('select');
    s.title = label;
    s.dataset.nav = '';
    s.setAttribute('aria-label', label);
    for (const [value, text] of options) s.add(opt(text, value));
    s.addEventListener('change', () => {
      run(s.value);
      this.editing?.focus();
    });
    this.controls.set(key, s);
    this.target?.appendChild(s);
  }

  private colorInput(): void {
    const wrap = document.createElement('label');
    wrap.className = 'color';
    wrap.title = 'Text colour';
    const face = document.createElement('button');
    face.type = 'button';
    face.tabIndex = -1;
    face.innerHTML = '<span style="font-weight:600">A</span><span class="swatch" style="background:#c00000"></span>';
    face.setAttribute('aria-hidden', 'true');
    const input = document.createElement('input');
    input.type = 'color';
    input.value = '#c00000';
    input.dataset.nav = '';
    input.setAttribute('aria-label', 'Text colour');
    input.addEventListener('change', () => {
      (face.querySelector('.swatch') as HTMLElement).style.background = input.value;
      this.editing?.format({ color: input.value.slice(1).toUpperCase() });
      this.editing?.focus();
    });
    wrap.append(face, input);
    this.controls.set('color', input);
    this.target?.appendChild(wrap);
  }

  /** Offer the document's own Normal, Title and Heading styles (or the ones given). */
  private fillStyles(): void {
    const select = this.controls.get('style') as HTMLSelectElement;
    const doc = this.viewer.document;
    let styles = this.options.styles;
    if (!styles) {
      const wanted = ['Normal', 'Title', 'Subtitle', 'Heading1', 'Heading2', 'Heading3', 'Heading4', 'Quote', 'ListParagraph'];
      styles = wanted.filter((id) => !doc || id === 'Normal' || doc.styles.get(id)).map((id) => ({ id, name: doc?.styles.get(id)?.name?.replace(/^heading/, 'Heading') ?? id }));
    }
    const key = styles.map((s) => s.id).join();
    if (select.dataset.key === key) return;
    select.dataset.key = key;
    select.replaceChildren(...styles.map((s) => opt(s.name, s.id)));
  }

  private update(format: CaretFormat | undefined): void {
    this.format = format;
    const on = !!format;
    for (const c of this.controls.values()) c.disabled = !on;
    const pressed = (key: string, value: boolean) => this.controls.get(key)?.setAttribute('aria-pressed', String(value));
    if (!format) {
      this.storyLabel.textContent = '';
      this.roving();
      return;
    }
    pressed('bold', format.bold);
    pressed('italic', format.italic);
    pressed('underline', format.underline);
    pressed('strike', format.strike);
    for (const a of ['left', 'center', 'right', 'justify']) pressed(`align-${a}`, format.align === a);
    pressed('bullets', format.list === 'bullet');
    pressed('numbers', format.list === 'number');
    pressed('ltr', format.direction === 'ltr');
    pressed('rtl', format.direction === 'rtl');
    pressed('track', format.trackChanges);
    (this.controls.get('undo') as HTMLButtonElement).disabled = !(this.editing?.canUndo ?? format.canUndo);
    (this.controls.get('redo') as HTMLButtonElement).disabled = !(this.editing?.canRedo ?? format.canRedo);
    const size = this.controls.get('size') as HTMLSelectElement;
    const sizeText = format.size !== undefined ? String(format.size) : '';
    if (sizeText && ![...size.options].some((o) => o.value === sizeText)) size.add(opt(sizeText, sizeText));
    size.value = sizeText;
    (this.controls.get('style') as HTMLSelectElement).value = format.styleId ?? 'Normal';
    (this.controls.get('highlight') as HTMLSelectElement).value = format.highlight ?? 'none';
    this.storyLabel.textContent = STORY_NAMES[format.story] ?? '';
    this.roving();
  }

  // Overflow ----------------------------------------------------------------------------------

  /**
   * Fit the groups into the row: those that do not fit move into the "more" menu, rightmost
   * first. Widths and room are measured unless given (tests, hosts without layout).
   */
  relayout(widths?: number[], available?: number): void {
    if (!this.el || !this.shown || this.fold) return;
    for (const g of this.groups) this.row.insertBefore(g, this.moreWrap);
    this.moreWrap.hidden = true;
    let w = widths;
    let room = available;
    if (!w || room === undefined) {
      const rowWidth = this.row.clientWidth;
      if (!rowWidth) {
        this.closeMenu();
        return;
      }
      w = this.groups.map((g) => g.offsetWidth + 2);
      room = rowWidth - 12 - (this.foldButton.offsetWidth + 2) - this.storyLabel.offsetWidth;
    }
    const keep = fitGroups(w, room, (this.moreButton.offsetWidth || 28) + 2);
    if (keep >= this.groups.length) {
      this.closeMenu();
      this.roving();
      return;
    }
    this.moreWrap.hidden = false;
    for (const g of this.groups.slice(keep)) this.pop.appendChild(g);
    if (!this.pop.hidden) this.placePop();
    this.roving();
  }

  private readonly onWinResize = (): void => {
    if (!this.pop.hidden) this.placePop();
  };

  private readonly onWinScroll = (e: Event): void => {
    if (this.pop.hidden || this.pop.contains(e.target as Node)) return;
    this.closeMenu();
  };

  private toggleMenu(): void {
    if (this.pop.hidden) {
      this.pop.hidden = false;
      this.placePop();
      window.addEventListener('resize', this.onWinResize);
      window.addEventListener('scroll', this.onWinScroll, true);
      this.moreButton.setAttribute('aria-expanded', 'true');
      this.roving();
    } else this.closeMenu();
  }

  /** Flip/clamp the overflow popover inside the viewport, anchored under the "more" button. */
  private placePop(): void {
    const b = this.moreButton.getBoundingClientRect();
    const r = this.pop.getBoundingClientRect();
    const vp = { w: window.innerWidth, h: window.innerHeight };
    // Right-align under the button; positionMenu flips up/clamps when it would overflow.
    const p = positionMenu({ x: b.right - r.width, y: b.bottom }, { w: r.width, h: r.height }, vp);
    if (p.y < b.bottom && p.y + r.height > b.top) p.y = Math.max(4, b.top - r.height);
    this.pop.style.left = `${p.x}px`;
    this.pop.style.top = `${p.y}px`;
  }

  private closeMenu(): void {
    window.removeEventListener('resize', this.onWinResize);
    window.removeEventListener('scroll', this.onWinScroll, true);
    if (!this.el) return;
    this.pop.hidden = true;
    this.moreButton.setAttribute('aria-expanded', 'false');
    this.roving();
  }

  // Keyboard ----------------------------------------------------------------------------------

  private navItems(): HTMLElement[] {
    const row = [...this.row.querySelectorAll<HTMLElement>('[data-nav]')];
    const more = this.pop.hidden ? [] : [...this.pop.querySelectorAll<HTMLElement>('[data-nav]')];
    // The "more" button comes after the row's controls and before the fold button.
    const fold = row.indexOf(this.foldButton);
    const items = [...row.slice(0, fold), ...more, this.foldButton];
    return items.filter((n) => !(n as Control).disabled && !(this.moreWrap.hidden && n === this.moreButton));
  }

  /** The one tab stop: the current one if still usable, else the first enabled visible control. */
  private roving(active?: HTMLElement): void {
    if (!this.el || !this.row) return;
    const usable = this.navItems();
    const all = [...this.el.querySelectorAll<HTMLElement>('[data-nav]')];
    const current = [active, all.find((n) => n.tabIndex === 0)].find((n) => n && usable.includes(n)) ?? usable[0];
    for (const n of all) n.tabIndex = n === current ? 0 : -1;
  }

  private onKey(e: KeyboardEvent): void {
    if (e.key === 'Escape' && !this.pop.hidden) {
      this.closeMenu();
      this.moreButton.focus();
      return;
    }
    const keys = ['ArrowRight', 'ArrowLeft', 'Home', 'End'];
    const target = e.target as HTMLElement;
    if (!keys.includes(e.key) || !target.hasAttribute('data-nav')) return;
    const items = this.navItems();
    const at = items.indexOf(target);
    if (at < 0) return;
    const next = e.key === 'Home' ? 0 : e.key === 'End' ? items.length - 1 : (at + (e.key === 'ArrowRight' ? 1 : -1) + items.length) % items.length;
    const to = items[next]!;
    e.preventDefault();
    this.roving(to);
    to.focus();
  }

  destroy(): void {
    for (const fn of this.cleanup) fn();
    this.resizeObserver?.disconnect();
    this.el?.remove();
  }
}

/** The earlier name of the toolbar. */
export { ToolbarController as FormatToolbar };
