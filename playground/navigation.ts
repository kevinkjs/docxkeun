/** Playground navigation pane: headings, bookmarks and links, with next/previous and back/forward. */
import type { DocxViewer, NavKind, NavTarget } from '../src/render/index.js';

const SECTIONS: Array<{ kind: NavKind; title: string; key: 'headings' | 'bookmarks' | 'links' }> = [
  { kind: 'heading', title: 'Headings', key: 'headings' },
  { kind: 'bookmark', title: 'Bookmarks', key: 'bookmarks' },
  { kind: 'link', title: 'Links', key: 'links' },
];

function el<K extends keyof HTMLElementTagNameMap>(tag: K, props: Partial<HTMLElementTagNameMap[K]> = {}, children: Array<Node | string> = []): HTMLElementTagNameMap[K] {
  const node = Object.assign(document.createElement(tag), props);
  node.append(...children);
  return node;
}

export function mountNavigation(pane: HTMLElement, viewer: DocxViewer, status: (text: string) => void): void {
  let showHidden = false;
  let active: NavTarget | undefined;
  const items = new Map<NavTarget, HTMLLIElement>();

  const describe = (t: NavTarget): string => {
    if (t.kind === 'link') return t.link?.anchor ? `→ #${t.link.anchor}` : (t.link?.url ?? '');
    return `p. ${t.page + 1}`;
  };

  const render = () => {
    items.clear();
    const back = el('button', { textContent: '← Back', disabled: !viewer.canGoBack, title: 'Back (Alt+←)' });
    const forward = el('button', { textContent: 'Forward →', disabled: !viewer.canGoForward, title: 'Forward (Alt+→)' });
    back.onclick = () => viewer.back();
    forward.onclick = () => viewer.forward();
    const parts: Node[] = [el('div', { className: 'history' }, [back, forward])];

    for (const { kind, title, key } of SECTIONS) {
      let list = viewer.navigation[key];
      if (kind === 'bookmark' && !showHidden) list = list.filter((b) => !b.hidden);
      const prev = el('button', { textContent: '↑', title: `Previous ${kind}`, ariaLabel: `Previous ${kind}` });
      const next = el('button', { textContent: '↓', title: `Next ${kind}`, ariaLabel: `Next ${kind}` });
      prev.onclick = () => viewer.previous(kind, { includeHidden: showHidden }) ?? status(`No earlier ${kind}`);
      next.onclick = () => viewer.next(kind, { includeHidden: showHidden }) ?? status(`No later ${kind}`);
      const ul = el('ul');
      for (const t of list) {
        const button = el('button', { title: t.link?.url ?? t.label }, [
          (t.level ? ' '.repeat(t.level) : '') + t.label,
          el('span', { className: 'meta', textContent: describe(t) }),
        ]);
        // Clicking a link entry shows the link itself; double-click follows it.
        button.onclick = () => viewer.goTo(t, { history: true });
        if (t.kind === 'link') button.ondblclick = () => viewer.followLink(t.link!);
        const li = el('li', {}, [button]);
        items.set(t, li);
        ul.appendChild(li);
      }
      const section = el('section', {}, [
        el('h2', {}, [title, el('span', { className: 'count', textContent: `(${list.length})` }), el('span', { className: 'step' }, [prev, next])]),
        list.length ? ul : el('div', { className: 'empty', textContent: 'None' }),
      ]);
      if (kind === 'bookmark') {
        const toggle = el('input', { type: 'checkbox', checked: showHidden });
        toggle.onchange = () => {
          showHidden = toggle.checked;
          render();
        };
        section.appendChild(el('label', {}, [toggle, 'Show hidden bookmarks (_Toc, _Hlk…)']));
      }
      parts.push(section);
    }
    pane.replaceChildren(...parts);
    markActive();
  };

  const markActive = () => {
    for (const [t, li] of items) li.classList.toggle('active', t === active);
    items.get(active!)?.scrollIntoView({ block: 'nearest' });
  };

  viewer.on('layout', () => {
    active = undefined;
    render();
  });
  viewer.on('navigate', ({ target, location }) => {
    active = target;
    render();
    status(target ? `${target.kind}: ${target.label} (page ${target.page + 1})` : `Page ${location.page + 1}`);
  });
  document.addEventListener('keydown', (e) => {
    if (!e.altKey || e.metaKey || e.ctrlKey) return;
    if (e.key === 'ArrowLeft' && viewer.back()) e.preventDefault();
    else if (e.key === 'ArrowRight' && viewer.forward()) e.preventDefault();
  });
  render();
}
