/**
 * The right-click menu (`viewer.contextMenu`, off by default). `hitTarget` says what was clicked,
 * `defaultItems` builds the menu for it (guard refusals show as disabled items with the reason),
 * and `ContextMenuController` wires the `contextmenu` event, moves the caret, and runs the
 * actions, each of them one undo step through `viewer.editing.transaction`.
 */
import { DocxEditor, objectNode, type EditAction } from '../api/editor.js';
import type { SelectionInfo } from '../api/selection.js';
import { readComments } from '../comments/read.js';
import type { ImageInline } from '../model/content.js';
import type { DocxDocument } from '../model/document.js';
import { linkSource, type LinkTarget } from '../model/fields.js';
import type { Page } from '../layout/types.js';
import type { XmlElement } from '../xml/index.js';
import type { CommentsController } from './comments.js';
import type { InlineEditing } from './editing.js';
import type { MenuHandle, MenuItem } from './ui/menu.js';
import { openMenu } from './ui/menu.js';
import { promptPopover } from './ui/popover.js';
import type { DocxViewer } from './viewer.js';

export type MenuTargetKind = 'text' | 'link' | 'comment' | 'image' | 'header' | 'footer';

export interface MenuTarget {
  kind: MenuTargetKind;
  link?: LinkTarget;
  commentId?: string;
  image?: ImageInline;
  /** The header or footer part. */
  part?: string;
}

/** What was clicked. `x` and `y` are viewport coordinates (the pointer), `page` is 0-based. */
export interface MenuContext {
  x: number;
  y: number;
  page: number;
  selection: SelectionInfo | undefined;
  target: MenuTarget;
}

export type FormatChoice = 'bold' | 'italic' | 'underline' | 'strike' | 'superscript' | 'subscript' | 'clear';
export type ParagraphChoice =
  | { align: 'left' | 'center' | 'right' | 'justify' }
  | { list: 'bullet' | 'decimal' }
  | { indent: 1 | -1 }
  | { direction: 'ltr' | 'rtl' }
  | { style: string };

type Run = () => void | Promise<void>;

/** What the menu items do when chosen. */
export interface MenuActions {
  copy: Run;
  cut: Run;
  paste: Run;
  pastePlain: Run;
  comment: Run;
  link: Run;
  format(f: FormatChoice): void | Promise<void>;
  paragraph(f: ParagraphChoice): void | Promise<void>;
  openLink: Run;
  copyLink: Run;
  editLink: Run;
  removeLink: Run;
  reply: Run;
  resolve: Run;
  deleteComment: Run;
  altText: Run;
  copyImage: Run;
  deleteImage: Run;
  editStory: Run;
  removeStory: Run;
}

export interface MenuEnv {
  editable: boolean;
  /** The refusal reason for an action, or `undefined` when it is allowed. */
  can(action: 'edit' | 'format' | 'comment' | 'link' | 'reply' | 'resolveComment' | 'deleteComment', detail?: { commentAuthor?: string }): string | undefined;
  /** Whether the async clipboard can be read (otherwise Paste just focuses the page for ⌘V). */
  clipboardRead: boolean;
  act: MenuActions;
  /** Author and state of a comment, for the comment group's header and Resolve/Reopen. */
  commentInfo?(id: string): { author: string; resolved: boolean } | undefined;
  /** Paragraph styles offered (id and name). */
  styles?: Array<{ id: string; name: string }>;
}

const hasRect = (r: { x: number; y: number; w: number; h: number }, x: number, y: number) => x >= r.x && x <= r.x + r.w && y >= r.y && y <= r.y + r.h;

/** What lies at a point of a laid-out page (points): link, comment, picture, header/footer, else text. */
export function hitTarget(page: Page, x: number, y: number): MenuTarget {
  const link = page.links.find((l) => hasRect(l, x, y));
  if (link) return { kind: 'link', link: link.target };
  const comment = page.comments.find((c) => c.w > 0 && hasRect(c, x, y));
  if (comment) return { kind: 'comment', commentId: comment.id };
  for (let i = page.items.length - 1; i >= 0; i--) {
    const item = page.items[i]!;
    if (item.type === 'image' && item.inline && !item.frameOnly && hasRect(item, x, y)) return { kind: 'image', image: item.inline };
  }
  const line = page.storyLines.find((l) => (l.story === 'header' || l.story === 'footer') && hasRect({ x: l.x - 2, y: l.y - 2, w: l.w + 4, h: l.h + 4 }, x, y));
  if (line && (line.story === 'header' || line.story === 'footer')) return { kind: line.story, ...(line.part ? { part: line.part } : {}) };
  return { kind: 'text' };
}

const DEFAULT_STYLES = [
  { id: 'Normal', name: 'Normal' },
  { id: 'Title', name: 'Title' },
  { id: 'Heading1', name: 'Heading 1' },
  { id: 'Heading2', name: 'Heading 2' },
  { id: 'Heading3', name: 'Heading 3' },
];

type Item = Exclude<MenuItem, '-'>;

const FIELD_LINK = 'This link cannot be edited here';

/** The default menu for what was clicked (see the spec's list); the app's items come last. */
export function defaultItems(at: MenuContext, env: MenuEnv): MenuItem[] {
  const { act } = env;
  const refuse = (action: Parameters<MenuEnv['can']>[0], detail?: { commentAuthor?: string }): string | undefined =>
    env.can(action, detail) ?? (env.editable ? undefined : 'The document is read-only');
  const off = (reason: string | undefined, hint?: string): boolean | string => reason ?? hint ?? false;
  const nothingSelected = !at.selection || at.selection.collapsed;
  const groups: Item[][] = [];

  const clipboard: Item[] = [];
  if (env.editable) clipboard.push({ id: 'cut', label: 'Cut', icon: 'cut', shortcut: '⌘X', disabled: off(refuse('edit'), at.selection?.collapsed ? 'Nothing selected' : undefined), run: act.cut });
  clipboard.push({ id: 'copy', label: 'Copy', icon: 'copy', shortcut: '⌘C', disabled: at.selection?.collapsed ? 'Nothing selected' : false, run: act.copy });
  if (env.editable) {
    clipboard.push({ id: 'paste', label: env.clipboardRead ? 'Paste' : 'Paste (⌘V)', icon: 'paste', ...(env.clipboardRead ? { shortcut: '⌘V' } : {}), disabled: off(refuse('edit')), run: act.paste });
    if (env.clipboardRead) clipboard.push({ id: 'paste-plain', label: 'Paste as plain text', shortcut: '⇧⌘V', disabled: off(refuse('edit')), run: act.pastePlain });
  }
  groups.push(clipboard);

  const kind = at.target.kind;
  if (kind === 'text' || kind === 'link' || kind === 'comment') {
    const formatReason = refuse('format');
    const choice = (id: FormatChoice, label: string, shortcut?: string): Item => ({ id: `format-${id}`, label, ...(shortcut ? { shortcut } : {}), disabled: off(formatReason), run: () => act.format(id) });
    const para = (id: string, label: string, p: ParagraphChoice): Item => ({ id: `paragraph-${id}`, label, disabled: off(formatReason), run: () => act.paragraph(p) });
    const styles = env.styles ?? DEFAULT_STYLES;
    groups.push([
      { id: 'comment', label: 'Comment', icon: 'comment', disabled: off(refuse('comment'), nothingSelected ? 'Select some text to comment on' : undefined), run: act.comment },
      { id: 'link', label: 'Link…', icon: 'link', disabled: off(refuse('link'), nothingSelected && kind !== 'link' ? 'Select some text to link' : undefined), run: act.link },
      {
        id: 'format',
        label: 'Format',
        disabled: off(formatReason),
        items: [
          choice('bold', 'Bold', '⌘B'),
          choice('italic', 'Italic', '⌘I'),
          choice('underline', 'Underline', '⌘U'),
          choice('strike', 'Strikethrough'),
          choice('superscript', 'Superscript'),
          choice('subscript', 'Subscript'),
          choice('clear', 'Clear formatting'),
        ],
      },
      {
        id: 'paragraph',
        label: 'Paragraph',
        disabled: off(formatReason),
        items: [
          para('left', 'Align left', { align: 'left' }),
          para('center', 'Align center', { align: 'center' }),
          para('right', 'Align right', { align: 'right' }),
          para('justify', 'Justify', { align: 'justify' }),
          '-',
          para('bullets', 'Bullets', { list: 'bullet' }),
          para('numbering', 'Numbering', { list: 'decimal' }),
          para('indent', 'Increase indent', { indent: 1 }),
          para('outdent', 'Decrease indent', { indent: -1 }),
          '-',
          para('ltr', 'Left to right', { direction: 'ltr' }),
          para('rtl', 'Right to left', { direction: 'rtl' }),
          '-',
          { id: 'styles', label: 'Styles', disabled: off(formatReason), items: styles.map((s) => para(`style-${s.id}`, s.name, { style: s.id })) },
        ],
      },
    ]);
  }

  if (kind === 'link' && at.target.link) {
    const link = at.target.link;
    // A link made by a HYPERLINK field has no w:hyperlink to change.
    const linkReason = refuse('link') ?? (linkSource(link) ? undefined : FIELD_LINK);
    groups.push([
      { id: 'link-address', label: link.url ?? `#${link.anchor ?? ''}`, icon: 'link', disabled: true },
      { id: 'open-link', label: 'Open link', run: act.openLink },
      { id: 'copy-link', label: 'Copy link address', run: act.copyLink },
      { id: 'edit-link', label: 'Edit link…', disabled: off(linkReason), run: act.editLink },
      { id: 'remove-link', label: 'Remove link', disabled: off(linkReason), run: act.removeLink },
    ]);
  }

  if (kind === 'comment' && at.target.commentId !== undefined) {
    const info = env.commentInfo?.(at.target.commentId);
    groups.push([
      { id: 'comment-header', label: info ? `Comment by ${info.author}` : 'Comment', icon: 'comment', disabled: true },
      { id: 'reply', label: 'Reply', disabled: off(refuse('reply')), run: act.reply },
      { id: 'resolve', label: info?.resolved ? 'Reopen' : 'Resolve', disabled: off(refuse('resolveComment')), run: act.resolve },
      { id: 'delete-comment', label: 'Delete comment', disabled: off(refuse('deleteComment', info ? { commentAuthor: info.author } : undefined)), run: act.deleteComment },
    ]);
  }

  if (kind === 'image') {
    const reason = refuse('edit');
    groups.push([
      { id: 'alt-text', label: 'Alt text…', disabled: off(reason), run: act.altText },
      { id: 'copy-picture', label: 'Copy picture', icon: 'copy', run: act.copyImage },
      { id: 'delete-picture', label: 'Delete', disabled: off(reason), run: act.deleteImage },
    ]);
  }

  if (kind === 'header' || kind === 'footer') {
    const reason = refuse('edit');
    groups.push([
      { id: 'edit-story', label: 'Edit header/footer', disabled: off(reason), run: act.editStory },
      { id: 'remove-story', label: 'Remove header/footer', disabled: off(reason), run: act.removeStory },
    ]);
  }

  const out: MenuItem[] = [];
  for (const group of groups) {
    if (out.length) out.push('-');
    out.push(...group);
  }
  return out;
}

/** A link field's text as a hyperlink target: `#name` is a bookmark, anything else an address. */
export function parseLinkTarget(value: string): { url: string } | { bookmark: string } | undefined {
  const v = value.trim();
  if (!v) return undefined;
  if (v.startsWith('#')) return v.length > 1 ? { bookmark: v.slice(1) } : undefined;
  return { url: /^[a-z][a-z0-9+.-]*:/i.test(v) ? v : `https://${v}` };
}

export interface ContextMenuOptions {
  /** Change the menu: receives what was clicked and the default items, returns the items to show. */
  items?: (at: MenuContext, defaults: MenuItem[]) => MenuItem[];
}

/** What the controller needs from the viewer. */
export type ContextMenuHost = Pick<DocxViewer, 'root' | 'container' | 'emit' | 'editing' | 'comments' | 'scale' | 'layout' | 'document' | 'followLink' | 'imageBlob'>;

const STYLE_WANTED = ['Normal', 'Title', 'Heading1', 'Heading2', 'Heading3'];

/** Controls the right-click menu of a viewer (`viewer.contextMenu`). Off until `enable()`. */
export class ContextMenuController {
  private on = false;
  private dead = false;
  private handle: MenuHandle | undefined;
  /** Closes the link and alt-text prompts the menu opened (on destroy). */
  private readonly prompts = new AbortController();

  constructor(private readonly host: ContextMenuHost, private readonly options: ContextMenuOptions = {}) {}

  get enabled(): boolean {
    return this.on;
  }

  enable(): void {
    if (this.on) return;
    this.on = true;
    this.host.root.addEventListener('contextmenu', this.onContextMenu);
    this.host.emit('chromechange', { part: 'contextMenu', visible: true });
  }

  disable(): void {
    if (!this.on) return;
    this.on = false;
    this.host.root.removeEventListener('contextmenu', this.onContextMenu);
    this.close();
    this.host.emit('chromechange', { part: 'contextMenu', visible: false });
  }

  /** Open the menu at a viewport point (from a "⋯" button, a long-press or code). */
  open(x: number, y: number): void {
    const under = typeof document.elementFromPoint === 'function' ? document.elementFromPoint(x, y) : null;
    let pageEl = under?.closest?.('.docxkeun-page') as HTMLElement | null | undefined;
    if (pageEl && !this.host.root.contains(pageEl)) pageEl = undefined;
    if (pageEl) {
      this.openOn(pageEl, x, y);
      return;
    }
    // Not over a page (a button, say): use the selection's page and line, else the first page.
    const spot = this.host.editing?.isEnabled ? this.host.editing.selectionSpot() : undefined;
    if (spot) this.openAt(spot.page, spot.x, spot.y + spot.h / 2, x, y, true);
    else this.openAt(0, 0, 0, x, y, false);
  }

  close(): void {
    this.handle?.close();
    this.handle = undefined;
  }

  destroy(): void {
    this.on = false;
    this.dead = true;
    this.host.root.removeEventListener('contextmenu', this.onContextMenu);
    this.close();
    // An open prompt resolves as cancelled, and anything still queued does nothing (see edit()).
    this.prompts.abort();
  }

  private readonly onContextMenu = (e: MouseEvent): void => {
    if (e.shiftKey) return;
    const pageEl = (e.target as Element | null)?.closest?.('.docxkeun-page') as HTMLElement | null | undefined;
    if (!pageEl || !this.host.root.contains(pageEl)) return;
    // No items (the app's `items` returned none): the browser's own menu shows.
    if (this.openOn(pageEl, e.clientX, e.clientY)) e.preventDefault();
  };

  private openOn(pageEl: HTMLElement, clientX: number, clientY: number): boolean {
    const rect = pageEl.getBoundingClientRect();
    const s = this.host.scale;
    return this.openAt(Number(pageEl.dataset.page ?? 0), (clientX - rect.left) / s, (clientY - rect.top) / s, clientX, clientY, true);
  }

  /** Build and show the menu for a point on a page (points), with the menu itself at a viewport point. False when there was nothing to show. */
  private openAt(page: number, px: number, py: number, clientX: number, clientY: number, hasPoint: boolean): boolean {
    const editing = this.host.editing;
    const layoutPage = this.host.layout?.pages[page];
    const target: MenuTarget = layoutPage && hasPoint ? hitTarget(layoutPage, px, py) : { kind: 'text' };
    const live = !!editing?.isEnabled;
    // As in Word, a right-click outside the selection first moves the caret there (not into a header).
    if (live && hasPoint && target.kind !== 'header' && target.kind !== 'footer' && !editing!.selectionContains(page, px, py)) editing!.placeCaretAt(page, px, py);
    const at: MenuContext = { x: clientX, y: clientY, page, selection: live ? editing!.selectionInfo() : undefined, target };
    const defaults = defaultItems(at, this.env(at));
    const items = this.options.items ? this.options.items(at, defaults) : defaults;
    if (!items.length) return false;
    this.close();
    this.handle = openMenu(items, { x: clientX, y: clientY }, { label: 'Document menu', ...(this.host.container ? { container: this.host.container } : {}), onClose: () => (this.handle = undefined) });
    return true;
  }

  // ---- what the items need -----------------------------------------------------------------

  private env(at: MenuContext): MenuEnv {
    const editing = this.host.editing;
    const editable = !!editing?.isEnabled;
    const clipboard = typeof navigator === 'undefined' ? undefined : navigator.clipboard;
    return {
      editable,
      can: (action, detail) => {
        if (!editing?.isEnabled) return 'The document is read-only';
        const guard = editing.guard;
        if (!guard) return undefined;
        try {
          guard.assert(action as EditAction, undefined, detail);
          return undefined;
        } catch (error) {
          return (error instanceof Error && error.message) || 'Not allowed';
        }
      },
      clipboardRead: !!(clipboard && (clipboard.read || clipboard.readText)),
      commentInfo: (id) => {
        const doc = this.host.document;
        const c = doc ? readComments(doc).find((x) => x.id === id) : undefined;
        return c ? { author: c.author, resolved: c.resolved } : undefined;
      },
      styles: this.styles(),
      act: this.actions(at),
    };
  }

  private styles(): Array<{ id: string; name: string }> {
    const doc = this.host.document;
    return STYLE_WANTED.filter((id) => !doc || id === 'Normal' || doc.styles.get(id)).map((id) => ({ id, name: doc?.styles.get(id)?.name?.replace(/^heading ?/i, 'Heading ') ?? DEFAULT_STYLES.find((s) => s.id === id)?.name ?? id }));
  }

  /** Run an edit as one undo step; a failure has already been reported through `editerror`. */
  private edit<T>(label: string, fn: (editor: DocxEditor) => T | Promise<T>): Promise<void> {
    const editing = this.host.editing;
    if (this.dead || !editing?.isEnabled) return Promise.resolve();
    return editing.transaction(label, fn).then(
      () => undefined,
      () => undefined,
    );
  }

  private async write(text: string): Promise<void> {
    try {
      await navigator.clipboard?.writeText(text);
    } catch {
      // Refused by the browser: nothing to copy to.
    }
  }

  private actions(at: MenuContext): MenuActions {
    const host = this.host;
    const target = at.target;
    const point = { x: at.x, y: at.y };
    const container = host.container;
    const signal = this.prompts.signal;
    const ask = (title: string, value: string, placeholder?: string, validate?: (v: string) => string | undefined) =>
      promptPopover({ title, value, ...(placeholder ? { placeholder } : {}), at: point, ...(validate ? { validate } : {}), ...(container ? { container } : {}), signal });
    const linkValidate = (v: string) => (parseLinkTarget(v) ? undefined : 'Enter a web address, or #bookmark');
    const withEditing = (): InlineEditing | undefined => (host.editing?.isEnabled ? host.editing : undefined);
    const selectedRange = (editing: InlineEditing) => {
      const sel = editing.selection!;
      return editing.rangeOf(sel.anchor, sel.focus);
    };
    /** The w:hyperlink that was right-clicked (not whatever the selection touches). */
    const clickedHyperlink = (): XmlElement => {
      const source = target.link && linkSource(target.link);
      if (!source) throw new Error(FIELD_LINK);
      return source as XmlElement;
    };
    /**
     * The editor for the part holding `node`: a link or picture in a header, footer or note is
     * changed through that part (its relationships, its dirty flag), with the editing layer's
     * guard, Track Changes and author. The transaction's own editor when it is that part.
     */
    const ownerEditor = (editor: DocxEditor, node: XmlElement): DocxEditor => {
      const doc = host.document;
      const editing = host.editing;
      const part = doc ? partHolding(doc, node) : undefined;
      if (!doc || !editing || !part || part === editor.part) return editor;
      return new DocxEditor(doc, editing.guard, { part, trackChanges: editing.trackChanges, author: editing.author });
    };
    const copyFromPage = async (): Promise<void> => {
      const text = window.getSelection?.()?.toString() ?? '';
      if (text) await this.write(text);
    };
    const paste = async (plain: boolean): Promise<void> => {
      const editing = withEditing();
      if (!editing) return;
      editing.focus();
      const clipboard = navigator.clipboard;
      if (!clipboard || (!clipboard.read && !clipboard.readText)) return; // ⌘V works from the focused page
      try {
        if (!plain && clipboard.read) {
          let text: string | undefined;
          for (const item of await clipboard.read()) {
            if (item.types.includes('text/html') && editing.pasteHtml(await (await item.getType('text/html')).text())) return;
            if (text === undefined && item.types.includes('text/plain')) text = await (await item.getType('text/plain')).text();
          }
          if (text !== undefined) return editing.pasteText(text);
        }
        if (clipboard.readText) return editing.pasteText(await clipboard.readText());
      } catch {
        editing.focus();
      }
    };

    return {
      copy: () => (withEditing()?.selection ? withEditing()!.copySelection(false) : copyFromPage()),
      cut: () => withEditing()?.copySelection(true),
      paste: () => paste(false),
      pastePlain: () => paste(true),
      comment: () => host.comments.compose(),
      link: async () => {
        const value = await ask('Link', '', 'https://… or #bookmark', linkValidate);
        const parsed = value === undefined ? undefined : parseLinkTarget(value);
        const editing = withEditing();
        if (!parsed || !editing?.selection) return;
        await this.edit('Add link', (editor) => editor.addHyperlink(selectedRange(editing), parsed));
      },
      format: (f) => {
        const editing = withEditing();
        if (!editing) return;
        switch (f) {
          case 'bold':
          case 'italic':
          case 'underline':
          case 'strike':
            return editing.toggle(f);
          case 'superscript':
          case 'subscript':
            return editing.format({ vertAlign: f });
          case 'clear': {
            if (!editing.selection) return;
            return this.edit('Clear formatting', (editor) => editor.clearFormatting(selectedRange(editing)));
          }
        }
      },
      paragraph: (f) => {
        const editing = withEditing();
        if (!editing) return;
        if ('align' in f) return editing.formatParagraphs({ align: f.align });
        if ('direction' in f) return editing.formatParagraphs({ direction: f.direction });
        if ('list' in f) return editing.setList(f.list);
        if ('indent' in f) return editing.indent(f.indent);
        return editing.setStyle(f.style);
      },
      openLink: () => {
        if (target.link) host.followLink(target.link);
      },
      copyLink: () => {
        const l = target.link;
        return l ? this.write(l.url ?? `#${l.anchor ?? ''}`) : undefined;
      },
      editLink: async () => {
        const link = target.link;
        const value = await ask('Edit link', link?.url ?? (link?.anchor ? `#${link.anchor}` : ''), 'https://… or #bookmark', linkValidate);
        const parsed = value === undefined ? undefined : parseLinkTarget(value);
        const editing = withEditing();
        if (!parsed || !editing) return;
        await this.edit('Edit link', (editor) => {
          const link = clickedHyperlink();
          return ownerEditor(editor, link).setHyperlinkTarget(link, parsed);
        });
      },
      removeLink: async () => {
        const editing = withEditing();
        if (!editing) return;
        await this.edit('Remove link', (editor) => {
          const link = clickedHyperlink();
          return ownerEditor(editor, link).removeHyperlink(link);
        });
      },
      reply: () => {
        if (target.commentId !== undefined) host.comments.reply(target.commentId);
      },
      resolve: async () => {
        const id = target.commentId;
        if (id === undefined) return;
        const resolved = !readComments(host.document!).find((c) => c.id === id)?.resolved;
        await this.edit(resolved ? 'Resolve comment' : 'Reopen comment', (editor) => editor.resolveComment(id, resolved));
        if (resolved) host.comments.close();
      },
      deleteComment: async () => {
        const id = target.commentId;
        if (id === undefined) return;
        await this.edit('Delete comment', (editor) => editor.deleteComment(id));
        host.comments.close();
      },
      altText: async () => {
        const image = target.image;
        if (!image) return;
        const text = await ask('Alt text', image.alt ?? '', 'Describe the picture');
        if (text === undefined) return;
        await this.edit('Alt text', (editor) => ownerEditor(editor, image.drawing).setImageAlt(image.drawing, text));
      },
      copyImage: async () => {
        const blob = target.image ? await host.imageBlob(target.image) : undefined;
        if (!blob || typeof ClipboardItem === 'undefined') return;
        try {
          await navigator.clipboard.write([new ClipboardItem({ [blob.type]: blob })]);
        } catch {
          // Refused: the browser does not take pictures on the clipboard.
        }
      },
      deleteImage: async () => {
        const image = target.image;
        const node = image && objectNode(image);
        if (!node) return;
        // A picture in a header, footer or note is deleted through that part's editor (still one undo step).
        await this.edit('Delete picture', (editor) => ownerEditor(editor, node).deleteRange({ start: { node, offset: 0 }, end: { node, offset: 1 } }));
      },
      editStory: () => {
        const editing = withEditing();
        if (editing && target.part && (target.kind === 'header' || target.kind === 'footer')) editing.enterStory({ kind: target.kind, part: target.part });
      },
      removeStory: async () => {
        const part = target.part;
        if (!part || (target.kind !== 'header' && target.kind !== 'footer')) return;
        const section = (host.layout?.pages[at.page]?.sectionIndex ?? 0) + 1;
        await this.edit('Remove header/footer', async (editor) => {
          const info = editor.headersAndFooters().find((h) => h.part === part);
          if (!info) throw new Error('This header or footer was not found');
          // Its own section when it has a reference there; otherwise the sections it is declared in.
          for (const s of info.sections.includes(section) ? [section] : info.sections) await editor.removeHeaderFooter(info.kind, { type: info.type, section: s });
        });
      },
    };
  }
}

function nodeIn(root: XmlElement, node: XmlElement): boolean {
  for (const c of root.children) if (c.type === 'element' && (c === node || nodeIn(c, node))) return true;
  return false;
}

/** The part whose XML holds `node`: the main document, a header, a footer or a notes part. */
function partHolding(doc: DocxDocument, node: XmlElement): string | undefined {
  if (nodeIn(doc.documentXml.root, node)) return doc.mainPart;
  return doc.storyParts().find((sp) => {
    const root = doc.xmlSync(sp.part)?.root;
    return !!root && nodeIn(root, node);
  })?.part;
}
