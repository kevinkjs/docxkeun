// @vitest-environment happy-dom
import { afterEach, describe, expect, it, vi } from 'vitest';
import { CommentsController, relativeTime, type CommentThread } from '../src/render/comments.js';
import type { CommentInfo } from '../src/comments/read.js';

const info = (id: string, author: string, text: string, extra: Partial<CommentInfo> = {}): CommentInfo => ({ id, author, text, blocks: [], resolved: false, el: {} as never, ...extra });
const thread = (id: string, y: number, extra: { author?: string; text?: string; replies?: CommentInfo[]; resolved?: boolean } = {}): CommentThread => ({
  comment: info(id, extra.author ?? 'Ada Lovelace', extra.text ?? `Text of ${id}`, { resolved: !!extra.resolved, date: '2020-01-01T10:00:00Z' }),
  replies: extra.replies ?? [],
  anchor: { page: 0, y, rects: [{ x: 10, y, w: 50, h: 12 }] },
});

function stub(opts: { threads?: CommentThread[]; editable?: boolean; deny?: string[]; width?: number; selection?: boolean } = {}) {
  const listeners = new Map<string, Array<(p: any) => void>>();
  const emitted: Array<{ event: string; payload: any }> = [];
  const root = document.createElement('div');
  document.body.appendChild(root);
  Object.defineProperty(root, 'clientWidth', { value: opts.width ?? 1200, configurable: true });
  const state = { threads: opts.threads ?? [], active: [] as string[], doc: {} as object, bottom: 0 };
  const transaction = vi.fn(async (_label: string, fn: (editor: any) => unknown) => fn(editor));
  const editor = {
    replyToComment: vi.fn(async () => '9'),
    resolveComment: vi.fn(async () => undefined),
    deleteComment: vi.fn(async () => undefined),
    addComment: vi.fn(async () => '10'),
  };
  const rebuild = vi.fn();
  const host: any = {
    root,
    threads: () => state.threads,
    anchorOf: (t: CommentThread) => ({ top: 100 + t.anchor.y, pageRight: 400, height: 12, bottom: 100 + t.anchor.y + 12 + state.bottom }),
    currentDocument: () => state.doc,
    columnX: () => 420,
    widestPage: () => 816,
    editable: () => opts.editable ?? true,
    can: (action: string) => !(opts.deny ?? []).includes(action),
    author: () => 'Grace',
    transaction,
    selection: () => (opts.selection === false ? undefined : { range: { start: 'a', end: 'b' }, top: 300, pageRight: 400, height: 12 }),
    setActive: (ids: string[]) => (state.active = ids),
    rebuild,
    on: (event: string, fn: (p: any) => void) => {
      (listeners.get(event) ?? listeners.set(event, []).get(event)!).push(fn);
      return () => {};
    },
    emit: (event: string, payload: any) => {
      emitted.push({ event, payload });
      for (const l of listeners.get(event) ?? []) l(payload);
    },
  };
  return { host, root, state, transaction, editor, emitted, rebuild, fire: (e: string, p?: any) => host.emit(e, p) };
}

const column = (root: HTMLElement) => root.querySelector('.docxkeun-comments') as HTMLElement | null;
const cards = (root: HTMLElement) => [...root.querySelectorAll('.dk-card')] as HTMLElement[];
const card = (root: HTMLElement, id: string) => root.querySelector(`.dk-card[data-comment="${id}"]`) as HTMLElement;
const buttons = (el: Element) => [...el.querySelectorAll('button')].map((b) => b.textContent!.trim());
const click = (el: Element) => el.dispatchEvent(new MouseEvent('click', { bubbles: true }));
const flush = () => new Promise((r) => setTimeout(r, 0));

function make(opts: Parameters<typeof stub>[0] = {}, options: ConstructorParameters<typeof CommentsController>[1] = {}) {
  const s = stub(opts);
  const ctl = new CommentsController(s.host, options);
  ctl.refresh();
  return { s, ctl };
}

afterEach(() => {
  document.body.replaceChildren();
});

describe('CommentsController cards', () => {
  it('shows resting cards with initials, a 2-line clamp, replies folded and resolved ones hidden', () => {
    const { s, ctl } = make({
      threads: [
        thread('1', 10, { replies: [info('2', 'Bob', 'r1', { parentId: '1' }), info('3', 'Cy', 'r2', { parentId: '1' })] }),
        thread('4', 200, { resolved: true }),
      ],
    });
    expect(ctl.effectiveMode).toBe('cards');
    expect(column(s.root)).not.toBeNull();
    expect(cards(s.root).map((c) => c.dataset.comment)).toEqual(['1']);
    const c = card(s.root, '1');
    expect(c.querySelector('.dk-avatar')!.textContent).toBe('AL');
    expect(c.querySelector('.dk-text')!.classList.contains('dk-clamp')).toBe(true);
    expect(c.querySelector('.dk-replies')!.textContent).toBe('2 replies');
    expect(c.getAttribute('aria-expanded')).toBe('false');
    expect(c.querySelector('.dk-reply')).toBeNull();
    const sw = s.root.querySelector('.dk-resolved-switch')!;
    expect(sw.textContent).toBe('Show resolved (1)');
    expect(sw.getAttribute('aria-checked')).toBe('false');
  });

  it('has no resolved switch without resolved threads, and says "1 reply" for one', () => {
    const { s } = make({ threads: [thread('1', 10, { replies: [info('2', 'Bob', 'r', { parentId: '1' })] })] });
    expect(s.root.querySelector('.dk-resolved-switch')).toBeNull();
    expect(card(s.root, '1').querySelector('.dk-replies')!.textContent).toBe('1 reply');
  });

  it('shows resolved threads, faded, when the switch is turned on', () => {
    const { s, ctl } = make({ threads: [thread('1', 10), thread('4', 200, { resolved: true })] });
    click(s.root.querySelector('.dk-resolved-switch')!);
    expect(ctl.showResolved).toBe(true);
    expect(cards(s.root).map((c) => c.dataset.comment)).toEqual(['1', '4']);
    expect(card(s.root, '4').classList.contains('resolved')).toBe(true);
    expect(s.root.querySelector('.dk-resolved-switch')!.getAttribute('aria-checked')).toBe('true');
  });

  it('starts with resolved threads shown when the option says so', () => {
    const { s } = make({ threads: [thread('4', 200, { resolved: true })] }, { showResolved: true });
    expect(cards(s.root)).toHaveLength(1);
  });

  it('opens a thread: class, aria-expanded, replies, connector, active highlight, commentclick', () => {
    const { s, ctl } = make({ threads: [thread('1', 10, { replies: [info('2', 'Bob', 'second voice', { parentId: '1' })] }), thread('5', 60)] });
    ctl.open('1');
    const c = card(s.root, '1');
    expect(c.classList.contains('open')).toBe(true);
    expect(c.getAttribute('aria-expanded')).toBe('true');
    expect(c.querySelector('.dk-connector')).not.toBeNull();
    expect(c.querySelector('.dk-text')!.classList.contains('dk-clamp')).toBe(false);
    expect(c.textContent).toContain('second voice');
    expect(c.querySelector('.dk-replies')).toBeNull();
    expect(card(s.root, '5').classList.contains('open')).toBe(false);
    expect(card(s.root, '5').querySelector('.dk-connector')).toBeNull();
    expect(s.state.active).toEqual(['1', '2']);
    expect(s.emitted.filter((e) => e.event === 'commentclick').map((e) => e.payload.comment.id)).toEqual(['1']);
    // The open card sits at its anchor.
    expect(c.style.top).toBe('110px');
  });

  it('opens by clicking a card, not again when it is already open, and a reply id opens its thread', () => {
    const { s, ctl } = make({ threads: [thread('1', 10, { replies: [info('2', 'Bob', 'r', { parentId: '1' })] }), thread('5', 60)] });
    click(card(s.root, '5'));
    expect(card(s.root, '5').classList.contains('open')).toBe(true);
    click(card(s.root, '5'));
    expect(s.emitted.filter((e) => e.event === 'commentclick')).toHaveLength(1);
    ctl.open('2');
    expect(card(s.root, '1').classList.contains('open')).toBe(true);
    expect(card(s.root, '5').classList.contains('open')).toBe(false);
  });

  it('opens from the keyboard with Enter on a focused card', () => {
    const { s } = make({ threads: [thread('1', 10)] });
    card(s.root, '1').dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
    expect(card(s.root, '1').classList.contains('open')).toBe(true);
  });

  it('closes on Esc and on a click elsewhere, but not on a click inside the column', () => {
    const { s, ctl } = make({ threads: [thread('1', 10)] });
    ctl.open('1');
    s.root.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
    expect(card(s.root, '1').classList.contains('open')).toBe(false);
    expect(s.state.active).toEqual([]);
    ctl.open('1');
    click(card(s.root, '1'));
    expect(card(s.root, '1').classList.contains('open')).toBe(true);
    document.body.dispatchEvent(new MouseEvent('click', { bubbles: true }));
    expect(card(s.root, '1').classList.contains('open')).toBe(false);
  });

  it('opens a thread when its highlighted text is clicked', () => {
    const { s } = make({ threads: [thread('1', 10, { replies: [info('2', 'Bob', 'r', { parentId: '1' })] })] });
    const page = document.createElement('div');
    const hl = document.createElement('div');
    hl.className = 'docxkeun-comment-hl';
    hl.dataset.comment = '2';
    hl.getBoundingClientRect = () => ({ left: 10, top: 10, right: 60, bottom: 22, width: 50, height: 12, x: 10, y: 10, toJSON() {} });
    page.appendChild(hl);
    s.root.appendChild(page);
    page.dispatchEvent(new MouseEvent('click', { bubbles: true, clientX: 30, clientY: 15 }));
    expect(card(s.root, '1').classList.contains('open')).toBe(true);
    // A click beside the highlight closes it.
    page.dispatchEvent(new MouseEvent('click', { bubbles: true, clientX: 300, clientY: 15 }));
    expect(card(s.root, '1').classList.contains('open')).toBe(false);
  });

  it('opens a resolved thread even while resolved ones are hidden', () => {
    const { s, ctl } = make({ threads: [thread('4', 10, { resolved: true })] });
    expect(cards(s.root)).toHaveLength(0);
    ctl.open('4');
    expect(card(s.root, '4').classList.contains('open')).toBe(true);
  });
});

describe('CommentsController connector and Esc', () => {
  it('draws the connector only in the gap, from the page edge to the card', () => {
    const { s, ctl } = make({ threads: [thread('1', 10)] });
    ctl.open('1');
    const line = card(s.root, '1').querySelector('.dk-connector') as HTMLElement;
    // Column at x=420, page right edge at 400: the line starts at 400 and never reaches into the page.
    const left = 420 - parseFloat(line.style.width);
    expect(left).toBeGreaterThanOrEqual(400);
    expect(parseFloat(line.style.width)).toBe(20);
  });

  it('closes on Esc from the editing input after opening from the highlight', () => {
    const { s } = make({ threads: [thread('1', 10)] });
    const page = document.createElement('div');
    const hl = document.createElement('div');
    hl.className = 'docxkeun-comment-hl';
    hl.dataset.comment = '1';
    hl.getBoundingClientRect = () => ({ left: 10, top: 10, right: 60, bottom: 22, width: 50, height: 12, x: 10, y: 10, toJSON() {} });
    const input = document.createElement('textarea');
    input.className = 'docxkeun-input';
    // The editor swallows Esc itself.
    input.addEventListener('keydown', (e) => {
      e.preventDefault();
      e.stopPropagation();
    });
    page.append(hl, input);
    s.root.appendChild(page);
    page.dispatchEvent(new MouseEvent('click', { bubbles: true, clientX: 30, clientY: 15 }));
    expect(card(s.root, '1').classList.contains('open')).toBe(true);
    input.focus();
    input.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true, cancelable: true }));
    expect(card(s.root, '1').classList.contains('open')).toBe(false);
  });

  it('ignores Esc from outside the viewer', () => {
    const { s, ctl } = make({ threads: [thread('1', 10)] });
    ctl.open('1');
    const other = document.createElement('input');
    document.body.appendChild(other);
    other.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
    expect(card(s.root, '1').classList.contains('open')).toBe(true);
  });
});

describe('CommentsController actions', () => {
  const open = (opts: Parameters<typeof stub>[0]) => {
    const m = make({ threads: [thread('1', 10, { author: 'Ada' })], ...opts });
    m.ctl.open('1');
    return m;
  };

  it('shows Reply, Resolve and Delete on an open card when editable and allowed', () => {
    const { s } = open({});
    expect(buttons(card(s.root, '1'))).toEqual(['Reply', 'Resolve', 'Delete']);
  });

  it('shows no actions when the viewer is read-only', () => {
    const { s } = open({ editable: false });
    expect(card(s.root, '1').querySelector('.dk-actions')).toBeNull();
  });

  it('hides each action the guard refuses', () => {
    expect(buttons(card(open({ deny: ['reply'] }).s.root, '1'))).toEqual(['Resolve', 'Delete']);
    expect(buttons(card(open({ deny: ['resolveComment'] }).s.root, '1'))).toEqual(['Reply', 'Delete']);
    expect(buttons(card(open({ deny: ['deleteComment'] }).s.root, '1'))).toEqual(['Reply', 'Resolve']);
  });

  it('asks the guard about deleting with the comment author', () => {
    const s = stub({ threads: [thread('1', 10, { author: 'Ada' })] });
    const can = vi.fn(() => true);
    s.host.can = can;
    const ctl = new CommentsController(s.host);
    ctl.refresh();
    ctl.open('1');
    expect(can).toHaveBeenCalledWith('deleteComment', { commentAuthor: 'Ada' });
  });

  it('keeps actions off resting cards', () => {
    const { s, ctl } = make({ threads: [thread('1', 10)] });
    expect(card(s.root, '1').querySelector('.dk-actions')).toBeNull();
    ctl.open('1');
    expect(card(s.root, '1').querySelector('.dk-actions')).not.toBeNull();
  });

  it('posts a reply through a transaction, with Post disabled while the box is empty', async () => {
    const { s } = open({});
    click([...card(s.root, '1').querySelectorAll('button')].find((b) => b.textContent === 'Reply')!);
    const box = card(s.root, '1').querySelector('textarea') as HTMLTextAreaElement;
    expect(box).not.toBeNull();
    const post = [...card(s.root, '1').querySelectorAll('button')].find((b) => b.textContent === 'Post') as HTMLButtonElement;
    expect(post.disabled).toBe(true);
    box.value = 'Agreed.';
    box.dispatchEvent(new Event('input', { bubbles: true }));
    expect(post.disabled).toBe(false);
    click(post);
    await flush();
    expect(s.transaction).toHaveBeenCalledTimes(1);
    expect(s.transaction.mock.calls[0]![0]).toBe('Reply');
    expect(s.editor.replyToComment).toHaveBeenCalledWith('1', { author: 'Grace', text: 'Agreed.' });
    expect(card(s.root, '1').querySelector('textarea')).toBeNull();
  });

  it('cancels a reply box without writing, and keeps typed text when the layout refreshes', () => {
    const { s, ctl } = open({});
    click([...card(s.root, '1').querySelectorAll('button')].find((b) => b.textContent === 'Reply')!);
    const box = card(s.root, '1').querySelector('textarea') as HTMLTextAreaElement;
    box.value = 'half typed';
    s.fire('layout', { pages: [] });
    expect((card(s.root, '1').querySelector('textarea') as HTMLTextAreaElement).value).toBe('half typed');
    expect(ctl).toBeDefined();
    click([...card(s.root, '1').querySelectorAll('button')].find((b) => b.textContent === 'Cancel')!);
    expect(card(s.root, '1').querySelector('textarea')).toBeNull();
    expect(s.transaction).not.toHaveBeenCalled();
  });

  it('resolves and deletes through transactions', async () => {
    const a = open({});
    click([...card(a.s.root, '1').querySelectorAll('button')].find((b) => b.textContent === 'Resolve')!);
    await flush();
    expect(a.s.editor.resolveComment).toHaveBeenCalledWith('1', true);
    expect(a.s.transaction.mock.calls[0]![0]).toBe('Resolve comment');
    document.body.replaceChildren();
    const b = open({});
    click([...card(b.s.root, '1').querySelectorAll('button')].find((x) => x.textContent === 'Delete')!);
    await flush();
    expect(b.s.editor.deleteComment).toHaveBeenCalledWith('1');
  });

  it('offers Reopen on a resolved thread', () => {
    const s = stub({ threads: [thread('1', 10, { resolved: true })] });
    const ctl = new CommentsController(s.host);
    ctl.refresh();
    ctl.open('1');
    expect(buttons(card(s.root, '1'))).toEqual(['Reply', 'Reopen', 'Delete']);
  });
});

describe('CommentsController compose', () => {
  it('puts a draft card with a text box beside the selection; Cancel removes it', () => {
    const { s, ctl } = make({ threads: [thread('1', 10)] });
    ctl.compose();
    const draft = s.root.querySelector('.dk-card.dk-draft') as HTMLElement;
    expect(draft).not.toBeNull();
    expect(draft.querySelector('textarea')).not.toBeNull();
    click([...draft.querySelectorAll('button')].find((b) => b.textContent === 'Cancel')!);
    expect(s.root.querySelector('.dk-draft')).toBeNull();
    expect(s.transaction).not.toHaveBeenCalled();
  });

  it('Post adds the comment through a transaction with the host author', async () => {
    const { s, ctl } = make({ threads: [thread('1', 10)] });
    ctl.compose();
    const draft = s.root.querySelector('.dk-draft') as HTMLElement;
    const box = draft.querySelector('textarea') as HTMLTextAreaElement;
    box.value = 'New thought';
    box.dispatchEvent(new Event('input', { bubbles: true }));
    click([...draft.querySelectorAll('button')].find((b) => b.textContent === 'Post')!);
    await flush();
    expect(s.transaction.mock.calls[0]![0]).toBe('Add comment');
    expect(s.editor.addComment).toHaveBeenCalledWith({ start: 'a', end: 'b' }, { author: 'Grace', text: 'New thought' });
    expect(s.root.querySelector('.dk-draft')).toBeNull();
  });

  it('keeps the draft when the transaction fails', async () => {
    const { s, ctl } = make({ threads: [thread('1', 10)] });
    s.transaction.mockRejectedValueOnce(new Error('refused'));
    ctl.compose();
    const draft = s.root.querySelector('.dk-draft') as HTMLElement;
    const box = draft.querySelector('textarea') as HTMLTextAreaElement;
    box.value = 'x';
    box.dispatchEvent(new Event('input', { bubbles: true }));
    click([...draft.querySelectorAll('button')].find((b) => b.textContent === 'Post')!);
    await flush();
    expect(s.root.querySelector('.dk-draft textarea')).not.toBeNull();
  });

  it('does nothing without a selection, when read-only, or when the guard refuses', () => {
    for (const opts of [{ selection: false }, { editable: false }, { deny: ['comment'] }]) {
      const { s, ctl } = make({ threads: [thread('1', 10)], ...opts });
      ctl.compose();
      expect(s.root.querySelector('.dk-draft')).toBeNull();
      document.body.replaceChildren();
    }
  });

  it('works in a document without comments yet: the column exists while the draft does', () => {
    const { s, ctl } = make({ threads: [] });
    expect(ctl.width).toBe(0);
    ctl.compose();
    expect(ctl.width).toBe(272);
    expect(s.rebuild).toHaveBeenCalled();
    expect(s.root.querySelector('.dk-draft')).not.toBeNull();
    click([...s.root.querySelectorAll('.dk-draft button')].find((b) => b.textContent === 'Cancel')!);
    expect(ctl.width).toBe(0);
    expect(column(s.root)).toBeNull();
  });
});

describe('CommentsController bubbles', () => {
  const two = () => [thread('1', 10), thread('2', 14), thread('3', 200)];

  it('draws one bubble per group, "2" for a merged group, in a 36px strip', () => {
    const { s, ctl } = make({ threads: two() });
    ctl.mode = 'bubbles';
    expect(ctl.effectiveMode).toBe('bubbles');
    expect(ctl.width).toBe(48);
    const bubbles = [...s.root.querySelectorAll('.dk-bubble')];
    expect(bubbles).toHaveLength(2);
    expect(bubbles[0]!.textContent).toBe('2');
    expect(bubbles[1]!.querySelector('svg')).not.toBeNull();
    expect(column(s.root)!.style.width).toBe('36px');
    expect(s.root.querySelector('.dk-card')).toBeNull();
  });

  it('opens a popover card from a bubble and steps between the threads with the arrows', () => {
    const { s, ctl } = make({ threads: two() });
    ctl.mode = 'bubbles';
    click(s.root.querySelectorAll('.dk-bubble')[0]!);
    const pop = s.root.querySelector('.dk-pop') as HTMLElement;
    expect(pop).not.toBeNull();
    expect(pop.classList.contains('docxkeun-ui-root')).toBe(true);
    expect(pop.querySelector('.dk-card')!.getAttribute('data-comment')).toBe('1');
    expect(pop.querySelector('.dk-count')!.textContent).toBe('1 of 2');
    const prev = pop.querySelector('.dk-prev') as HTMLButtonElement;
    const next = pop.querySelector('.dk-next') as HTMLButtonElement;
    expect(prev.disabled).toBe(true);
    click(next);
    expect(s.root.querySelector('.dk-pop .dk-card')!.getAttribute('data-comment')).toBe('2');
    expect(s.root.querySelector('.dk-pop .dk-count')!.textContent).toBe('2 of 2');
    expect((s.root.querySelector('.dk-next') as HTMLButtonElement).disabled).toBe(true);
    click(s.root.querySelector('.dk-prev')!);
    expect(s.root.querySelector('.dk-pop .dk-card')!.getAttribute('data-comment')).toBe('1');
    s.root.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
    expect(s.root.querySelector('.dk-pop')).toBeNull();
  });

  it('keeps a reply being typed (text and focus) across a layout refresh in bubble mode', () => {
    const { s, ctl } = make({ threads: two() });
    ctl.mode = 'bubbles';
    click(s.root.querySelectorAll('.dk-bubble')[0]!);
    click([...s.root.querySelectorAll('.dk-pop button')].find((b) => b.textContent === 'Reply')!);
    const box = s.root.querySelector('.dk-pop textarea') as HTMLTextAreaElement;
    box.value = 'half typed';
    box.focus();
    const bubble = s.root.querySelector('.dk-bubble');
    s.state.threads = two();
    s.fire('layout', { pages: [] });
    window.dispatchEvent(new Event('resize'));
    ctl.refresh();
    expect(s.root.querySelector('.dk-pop textarea')).toBe(box);
    expect(box.value).toBe('half typed');
    expect(document.activeElement).toBe(box);
    expect(s.root.querySelector('.dk-bubble')).toBe(bubble);
  });

  it('shows a single "30" bubble for 30 threads on the same spot', () => {
    const many = Array.from({ length: 30 }, (_, i) => thread(String(i + 1), 50));
    const { s, ctl } = make({ threads: many });
    ctl.mode = 'bubbles';
    const bubbles = [...s.root.querySelectorAll('.dk-bubble')];
    expect(bubbles).toHaveLength(1);
    expect(bubbles[0]!.textContent).toBe('30');
  });

  it('chooses the mode from the viewer width when set to auto, with hysteresis', () => {
    const s = stub({ threads: two(), width: 1200 });
    const ctl = new CommentsController(s.host);
    ctl.refresh();
    expect(ctl.mode).toBe('auto');
    expect(ctl.effectiveMode).toBe('cards');
    // 0.75 * 816 + 300 = 912
    Object.defineProperty(s.root, 'clientWidth', { value: 800, configurable: true });
    ctl.refresh();
    expect(ctl.effectiveMode).toBe('bubbles');
    Object.defineProperty(s.root, 'clientWidth', { value: 930, configurable: true });
    ctl.refresh();
    expect(ctl.effectiveMode).toBe('bubbles');
    Object.defineProperty(s.root, 'clientWidth', { value: 960, configurable: true });
    ctl.refresh();
    expect(ctl.effectiveMode).toBe('cards');
  });

  it('asks the host to rebuild its rows and announces the change when the mode switches', () => {
    const { s, ctl } = make({ threads: two() });
    s.rebuild.mockClear();
    s.emitted.length = 0;
    ctl.mode = 'bubbles';
    expect(s.rebuild).toHaveBeenCalledTimes(1);
    expect(s.emitted).toContainEqual({ event: 'chromechange', payload: { part: 'comments', visible: true, collapsed: true } });
  });
});

describe('CommentsController bubble popover placement', () => {
  const two = () => [thread('1', 10), thread('2', 14), thread('3', 400)];
  let restore: (() => void) | undefined;
  /** happy-dom does no layout: the popover is 120px tall and the viewer shows 600px. */
  function geometry(root: HTMLElement, view = { height: 600, width: 1200 }) {
    Object.defineProperty(root, 'clientHeight', { value: view.height, configurable: true });
    Object.defineProperty(root, 'clientWidth', { value: view.width, configurable: true });
    const desc = Object.getOwnPropertyDescriptor(HTMLElement.prototype, 'offsetHeight');
    Object.defineProperty(HTMLElement.prototype, 'offsetHeight', {
      configurable: true,
      get(this: HTMLElement) {
        return this.classList.contains('dk-pop') ? 120 : 0;
      },
    });
    restore = () => {
      if (desc) Object.defineProperty(HTMLElement.prototype, 'offsetHeight', desc);
      else delete (HTMLElement.prototype as { offsetHeight?: number }).offsetHeight;
    };
  }
  afterEach(() => {
    restore?.();
    restore = undefined;
  });

  it('opens below the last line of the commented text, right-aligned to the bubble strip', () => {
    const { s, ctl } = make({ threads: two() });
    geometry(s.root);
    s.state.bottom = 30; // the comment spans three lines
    ctl.mode = 'bubbles';
    click(s.root.querySelectorAll('.dk-bubble')[0]!);
    const pop = s.root.querySelector('.dk-pop') as HTMLElement;
    const anchorBottom = 100 + 10 + 12 + 30;
    expect(parseFloat(pop.style.top)).toBeGreaterThanOrEqual(anchorBottom);
    expect(parseFloat(pop.style.top)).toBeLessThan(anchorBottom + 12);
    // Its right edge is the strip's right edge (the strip is 36px wide at x = 420).
    expect(parseFloat(pop.style.left) + 260).toBe(36);
  });

  it('flips above the anchor when there is no room below, and stays inside the viewer horizontally', () => {
    const { s, ctl } = make({ threads: two() });
    geometry(s.root, { height: 560, width: 1200 });
    ctl.mode = 'bubbles';
    click(s.root.querySelectorAll('.dk-bubble')[1]!);
    const pop = s.root.querySelector('.dk-pop') as HTMLElement;
    // Thread 3: anchor top 500, bottom 512; 512 + 120 does not fit in 560.
    expect(parseFloat(pop.style.top) + 120).toBeLessThanOrEqual(500);
    expect(parseFloat(pop.style.top)).toBeGreaterThanOrEqual(0);
    s.host.columnX = () => 100;
    ctl.refresh();
    expect(100 + parseFloat(pop.style.left)).toBeGreaterThanOrEqual(0);
  });

  it('puts the draft below the selection too', () => {
    const { s, ctl } = make({ threads: two() });
    geometry(s.root);
    ctl.mode = 'bubbles';
    ctl.compose();
    const pop = s.root.querySelector('.dk-pop') as HTMLElement;
    expect(parseFloat(pop.style.top)).toBeGreaterThanOrEqual(300 + 12);
  });
});

describe('CommentsController and a new document', () => {
  it('drops the draft when another document is loaded', () => {
    const { s, ctl } = make({ threads: [] });
    ctl.compose();
    expect(s.root.querySelector('.dk-draft')).not.toBeNull();
    s.state.doc = {};
    s.fire('layout', { pages: [] });
    expect(s.root.querySelector('.dk-draft')).toBeNull();
    expect(ctl.width).toBe(0);
  });

  it('closes a reply box when another document is loaded', () => {
    const { s, ctl } = make({ threads: [thread('1', 10)] });
    ctl.reply('1');
    expect(s.root.querySelector('.dk-card textarea')).not.toBeNull();
    s.state.doc = {};
    s.state.threads = [thread('1', 10)];
    s.fire('layout', { pages: [] });
    expect(s.root.querySelector('.dk-card textarea')).toBeNull();
  });

  it('forgets a new comment that got no anchor by the next layout', async () => {
    const { s, ctl } = make({ threads: [thread('1', 10)] });
    ctl.compose();
    const box = s.root.querySelector('.dk-draft textarea') as HTMLTextAreaElement;
    box.value = 'Floating';
    box.dispatchEvent(new Event('input', { bubbles: true }));
    click([...s.root.querySelectorAll('.dk-draft button')].find((b) => b.textContent === 'Post')!);
    await flush();
    // The new comment ('10') has no anchor in this layout...
    s.fire('layout', { pages: [] });
    expect(ctl.openThread).toBeUndefined();
    // ...and does not pop open later when one appears (after an undo and redo, say).
    s.state.threads = [thread('1', 10), thread('10', 300)];
    s.fire('layout', { pages: [] });
    expect(ctl.openThread).toBeUndefined();
  });
});

describe('CommentsController visibility', () => {
  it('has no width and no column element without threads, and nothing throws', () => {
    const { s, ctl } = make({ threads: [] });
    expect(ctl.width).toBe(0);
    expect(column(s.root)).toBeNull();
    expect(() => {
      ctl.open('x');
      ctl.close();
      s.fire('layout', { pages: [] });
    }).not.toThrow();
  });

  it('hide() removes the column and the width; show() brings them back; both announce', () => {
    const { s, ctl } = make({ threads: [thread('1', 10)] });
    expect(ctl.width).toBe(272);
    s.emitted.length = 0;
    ctl.hide();
    expect(ctl.visible).toBe(false);
    expect(ctl.width).toBe(0);
    expect(column(s.root)).toBeNull();
    ctl.show();
    expect(ctl.visible).toBe(true);
    expect(column(s.root)).not.toBeNull();
    expect(s.emitted.filter((e) => e.event === 'chromechange').map((e) => e.payload)).toEqual([
      { part: 'comments', visible: false, collapsed: false },
      { part: 'comments', visible: true, collapsed: false },
    ]);
  });

  it('starts hidden when told so, and re-places when the layout changes', () => {
    const s = stub({ threads: [thread('1', 10)] });
    const ctl = new CommentsController(s.host, { visible: false });
    ctl.refresh();
    expect(column(s.root)).toBeNull();
    ctl.show();
    s.state.threads = [thread('1', 10), thread('2', 300)];
    s.fire('layout', { pages: [] });
    expect(cards(s.root).map((c) => c.dataset.comment)).toEqual(['1', '2']);
    s.state.threads = [thread('2', 300)];
    s.fire('layout', { pages: [] });
    expect(cards(s.root).map((c) => c.dataset.comment)).toEqual(['2']);
  });

  it('lays cards out without overlap using their measured heights', () => {
    const s = stub({ threads: [thread('1', 10), thread('2', 12), thread('3', 14)] });
    Object.defineProperty(HTMLElement.prototype, 'offsetHeight', { configurable: true, get: () => 60 });
    try {
      const ctl = new CommentsController(s.host);
      ctl.refresh();
      const tops = cards(s.root).map((c) => parseFloat(c.style.top));
      expect(tops).toEqual([110, 178, 246]);
    } finally {
      delete (HTMLElement.prototype as any).offsetHeight;
    }
  });
});

describe('relativeTime', () => {
  const now = Date.parse('2024-06-10T12:00:00Z');
  it('words recent times and falls back to a date', () => {
    expect(relativeTime('2024-06-10T11:59:50Z', now)).toBe('just now');
    expect(relativeTime('2024-06-10T11:55:00Z', now)).toBe('5m ago');
    expect(relativeTime('2024-06-10T09:00:00Z', now)).toBe('3h ago');
    expect(relativeTime('2024-06-08T12:00:00Z', now)).toBe('2d ago');
    expect(relativeTime('2023-01-05T12:00:00Z', now)).toMatch(/2023/);
    expect(relativeTime(undefined, now)).toBe('');
  });
});
