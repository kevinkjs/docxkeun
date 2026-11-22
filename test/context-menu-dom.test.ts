// @vitest-environment happy-dom
import { afterEach, describe, expect, it, vi } from 'vitest';
import { DocxEditor } from '../src/api/index.js';
import { DocxDocument } from '../src/model/index.js';
import { findAll } from '../src/xml/index.js';
import { setLinkSource } from '../src/model/fields.js';
import { buildDocx, p, W_NS } from './helpers/docx.js';
import { ContextMenuController, parseLinkTarget } from '../src/render/context-menu.js';

function stub(opts: { editable?: boolean; items?: any } = {}) {
  const emitted: Array<{ event: string; payload: any }> = [];
  const container = document.createElement('div');
  const root = document.createElement('div');
  const pageEl = document.createElement('div');
  pageEl.className = 'docxkeun-page';
  pageEl.dataset.page = '0';
  const outside = document.createElement('div');
  root.appendChild(pageEl);
  container.append(root, outside);
  document.body.appendChild(container);
  const placed: number[][] = [];
  const editing: any = {
    isEnabled: opts.editable ?? true,
    guard: undefined,
    author: 'Tester',
    trackChanges: false,
    selection: undefined,
    selectionInfo: () => undefined,
    selectionContains: () => false,
    placeCaretAt: (...a: number[]) => placed.push(a),
    selectionSpot: () => undefined,
  };
  const host: any = {
    root,
    container,
    emit: (event: string, payload: any) => emitted.push({ event, payload }),
    editing,
    comments: { compose: vi.fn(), reply: vi.fn(), close: vi.fn() },
    scale: 1,
    layout: { pages: [{ links: [], comments: [], items: [], storyLines: [] }] },
    document: undefined,
    followLink: vi.fn(),
    imageBlob: vi.fn(),
  };
  const ctl = new ContextMenuController(host, opts.items ? { items: opts.items } : {});
  const fire = (target: Element, init: MouseEventInit = {}) => {
    const e = new MouseEvent('contextmenu', { bubbles: true, cancelable: true, clientX: 40, clientY: 50, ...init });
    target.dispatchEvent(e);
    return e;
  };
  return { ctl, host, container, root, pageEl, outside, emitted, placed, fire };
}

afterEach(() => {
  document.body.innerHTML = '';
});

const flush = () => new Promise((r) => setTimeout(r, 0));

describe('ContextMenuController', () => {
  it('is off until enabled, then opens a menu on a page and moves the caret', () => {
    const s = stub();
    expect(s.ctl.enabled).toBe(false);
    expect(s.fire(s.pageEl).defaultPrevented).toBe(false);
    expect(document.querySelector('[role=menu]')).toBeNull();
    s.ctl.enable();
    expect(s.emitted).toEqual([{ event: 'chromechange', payload: { part: 'contextMenu', visible: true } }]);
    const e = s.fire(s.pageEl);
    expect(e.defaultPrevented).toBe(true);
    expect(s.container.querySelector('[role=menu]')).not.toBeNull();
    expect(s.placed).toEqual([[0, 40, 50]]);
  });

  it('leaves Shift + right-click and clicks outside the pages alone', () => {
    const s = stub();
    s.ctl.enable();
    expect(s.fire(s.pageEl, { shiftKey: true }).defaultPrevented).toBe(false);
    expect(s.fire(s.outside).defaultPrevented).toBe(false);
    expect(document.querySelector('[role=menu]')).toBeNull();
  });

  it('stops after disable() and closes an open menu', () => {
    const s = stub();
    s.ctl.enable();
    s.fire(s.pageEl);
    s.ctl.disable();
    expect(document.querySelector('[role=menu]')).toBeNull();
    expect(s.fire(s.pageEl).defaultPrevented).toBe(false);
    expect(s.emitted.at(-1)).toEqual({ event: 'chromechange', payload: { part: 'contextMenu', visible: false } });
  });

  it('lets the app change the items', () => {
    const run = vi.fn();
    const s = stub({ items: (_at: any, defaults: any[]) => [...defaults, '-', { label: 'Ask Claude', run }] });
    s.ctl.enable();
    s.fire(s.pageEl);
    const labels = [...document.querySelectorAll('[role=menuitem]')].map((n) => n.textContent);
    expect(labels.some((l) => l?.includes('Ask Claude'))).toBe(true);
    expect(labels.some((l) => l?.includes('Copy'))).toBe(true);
  });

  it('lets the browser menu show when the app returns no items', () => {
    const s = stub({ items: () => [] });
    s.ctl.enable();
    const e = s.fire(s.pageEl);
    expect(e.defaultPrevented).toBe(false);
    expect(document.querySelector('[role=menu]')).toBeNull();
  });

  it('read-only: no caret move, no Cut or Paste, Comment disabled with a reason', () => {
    const s = stub({ editable: false });
    s.ctl.enable();
    s.fire(s.pageEl);
    expect(s.placed).toEqual([]);
    const rows = [...document.querySelectorAll<HTMLElement>('[role=menuitem]')];
    const text = rows.map((r) => r.textContent ?? '');
    expect(text.some((t) => t.startsWith('Cut'))).toBe(false);
    expect(text.some((t) => t.startsWith('Paste'))).toBe(false);
    const comment = rows.find((r) => r.textContent?.startsWith('Comment'))!;
    expect(comment.getAttribute('aria-disabled')).toBe('true');
    expect(comment.title).toBe('The document is read-only');
  });
});

describe('parseLinkTarget', () => {
  it('reads addresses and #bookmarks', () => {
    expect(parseLinkTarget('https://a.co/x')).toEqual({ url: 'https://a.co/x' });
    expect(parseLinkTarget('example.com')).toEqual({ url: 'https://example.com' });
    expect(parseLinkTarget('mailto:a@b.co')).toEqual({ url: 'mailto:a@b.co' });
    expect(parseLinkTarget('#top')).toEqual({ bookmark: 'top' });
    expect(parseLinkTarget('#')).toBeUndefined();
    expect(parseLinkTarget('  ')).toBeUndefined();
  });
});

describe('Delete picture', () => {
  const EMU = 12700;
  const pic = (id: number) => `<w:r><w:drawing><wp:inline><wp:extent cx="${50 * EMU}" cy="${50 * EMU}"/><wp:docPr id="${id}" name="Pic${id}"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:blipFill><a:blip r:embed="rIdX"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;
  const SECT = '<w:sectPr><w:headerReference w:type="default" r:id="rIdH"/><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>';
  const HDR = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr ${W_NS}><w:p>${pic(2)}</w:p></w:hdr>`;

  it('removes a header picture from the header part only', async () => {
    const doc = await DocxDocument.open(
      await buildDocx({
        body: `<w:p>${pic(1)}</w:p>`,
        sectPr: SECT,
        rels: [['rIdH', 'header', 'header1.xml']],
        parts: { 'word/header1.xml': { data: HDR, contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml' } },
      }),
    );
    const headerPart = doc.storyParts().find((sp) => sp.kind === 'header')!.part;
    const headerDrawings = () => findAll(doc.xmlSync(headerPart)!.root, 'w:drawing');
    // A hand-made layout: the header picture as an image item (fonts are not needed here).
    const drawing = headerDrawings()[0]!;
    const run = findAll(doc.xmlSync(headerPart)!.root, 'w:r')[0]!;
    const image = { type: 'image' as const, x: 100, y: 20, w: 50, h: 50, inline: { kind: 'image', run, drawing, target: undefined, external: false } };
    const layout = { pages: [{ links: [], comments: [], items: [image], storyLines: [] }] };
    const s = stub();
    s.host.document = doc;
    s.host.layout = layout;
    const bodyEditor = new DocxEditor(doc);
    s.host.editing.transaction = async (_l: string, fn: (e: DocxEditor) => unknown) => fn(bodyEditor);
    s.host.editing.author = 'T';
    s.ctl.enable();
    s.fire(s.pageEl, { clientX: image.x + 5, clientY: image.y + 5 });
    const row = [...document.querySelectorAll<HTMLElement>('[role=menuitem]')].find((r) => r.textContent === 'Delete')!;
    expect(row.getAttribute('aria-disabled')).not.toBe('true');
    row.click();
    await new Promise((r) => setTimeout(r, 0));
    expect(headerDrawings()).toHaveLength(0);
    expect(findAll(doc.documentXml.root, 'w:drawing')).toHaveLength(1);
  });
});

describe('links and pictures in headers, footers and notes', () => {
  const EMU = 12700;
  const FTR_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink';
  const FTR = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:ftr ${W_NS}><w:p><w:hyperlink r:id="rIdL"><w:r><w:t>Footer link</w:t></w:r></w:hyperlink></w:p></w:ftr>`;
  const FTR_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdL" Type="${FTR_REL}" Target="https://old.example/" TargetMode="External"/></Relationships>`;
  const pic = `<w:r><w:drawing><wp:inline><wp:extent cx="${50 * EMU}" cy="${50 * EMU}"/><wp:docPr id="2" name="Pic2"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:blipFill><a:blip r:embed="rIdX"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;
  const HDR = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr ${W_NS}><w:p><w:r><w:t>Logo </w:t></w:r>${pic}</w:p></w:hdr>`;
  const SECT = '<w:sectPr><w:headerReference w:type="default" r:id="rIdH"/><w:footerReference w:type="default" r:id="rIdF"/><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>';
  const FOOTER = 'word/footer1.xml';
  const HEADER = 'word/header1.xml';

  async function setup() {
    const doc = await DocxDocument.open(
      await buildDocx({
        body: p('Body text.'),
        sectPr: SECT,
        rels: [
          ['rIdH', 'header', 'header1.xml'],
          ['rIdF', 'footer', 'footer1.xml'],
        ],
        parts: {
          [HEADER]: { data: HDR, contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml' },
          [FOOTER]: { data: FTR, contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml' },
          'word/_rels/footer1.xml.rels': { data: FTR_RELS },
        },
      }),
    );
    const hyperlink = findAll(doc.xmlSync(FOOTER)!.root, 'w:hyperlink')[0]!;
    const drawing = findAll(doc.xmlSync(HEADER)!.root, 'w:drawing')[0]!;
    const run = findAll(doc.xmlSync(HEADER)!.root, 'w:r')[1]!;
    // A hand-made layout: the footer link and the header picture (fonts are not needed here).
    const link = { x: 100, y: 700, w: 80, h: 14, target: setLinkSource({ url: 'https://old.example/' }, hyperlink) };
    const image = { type: 'image' as const, x: 300, y: 20, w: 50, h: 50, inline: { kind: 'image', run, drawing, target: undefined, external: false } };
    const s = stub();
    s.host.document = doc;
    s.host.layout = { pages: [{ links: [link], comments: [], items: [image], storyLines: [] }] };
    // The caret is in the body: the transaction's editor edits the main document.
    const labels: string[] = [];
    s.host.editing.transaction = vi.fn(async (label: string, fn: (e: DocxEditor) => unknown) => {
      labels.push(label);
      return fn(new DocxEditor(doc, s.host.editing.guard, { trackChanges: s.host.editing.trackChanges, author: s.host.editing.author }));
    });
    s.ctl.enable();
    const choose = (at: { x: number; y: number }, label: string) => {
      s.fire(s.pageEl, { clientX: at.x + 5, clientY: at.y + 5 });
      const row = [...document.querySelectorAll<HTMLElement>('[role=menuitem]')].find((r) => r.textContent === label)!;
      expect(row.getAttribute('aria-disabled')).not.toBe('true');
      row.click();
    };
    return { doc, s, link, image, labels, choose };
  }

  const relTargets = (doc: DocxDocument, part: string) => [...doc.rels(part).entries()].map(([id, r]) => ({ id, target: r.target }));

  it('Edit link… on a footer link adds the relationship to the footer, not the body', async () => {
    const { doc, s, link, choose, labels } = await setup();
    const bodyRelsBefore = relTargets(doc, doc.mainPart).length;
    choose(link, 'Edit link…');
    await flush();
    const input = s.container.querySelector<HTMLInputElement>('.dk-popover input')!;
    expect(input.value).toBe('https://old.example/');
    input.value = 'https://new.example/';
    s.container.querySelector<HTMLButtonElement>('.dk-popover [data-role=ok]')!.click();
    await flush();
    expect(labels).toEqual(['Edit link']);
    const again = await DocxDocument.open(await doc.save());
    const footerLink = findAll(again.xmlSync(FOOTER)!.root, 'w:hyperlink')[0]!;
    const id = footerLink.attrs['r:id']!;
    expect(again.rels(FOOTER).get(id)?.target).toBe('https://new.example/');
    // The old relationship is gone and the body did not gain one.
    expect(relTargets(again, FOOTER).map((r) => r.target)).toEqual(['https://new.example/']);
    expect(relTargets(again, again.mainPart)).toHaveLength(bodyRelsBefore);
    expect(relTargets(again, again.mainPart).some((r) => r.target === 'https://new.example/')).toBe(false);
  });

  it('Remove link on a footer link unwraps it and drops the footer relationship', async () => {
    const { doc, s, link, choose } = await setup();
    const errors: unknown[] = [];
    s.host.editing.transaction.mockImplementation(async (_l: string, fn: (e: DocxEditor) => unknown) => {
      try {
        return await fn(new DocxEditor(doc));
      } catch (e) {
        errors.push(e);
        throw e;
      }
    });
    choose(link, 'Remove link');
    await flush();
    expect(errors).toEqual([]);
    const again = await DocxDocument.open(await doc.save());
    expect(findAll(again.xmlSync(FOOTER)!.root, 'w:hyperlink')).toHaveLength(0);
    expect(findAll(again.xmlSync(FOOTER)!.root, 'w:t').map((t) => (t.children[0] as { value: string }).value)).toEqual(['Footer link']);
    expect(again.rels(FOOTER).has('rIdL')).toBe(false);
  });

  it('deletes a header picture as a tracked change when the viewer tracks changes', async () => {
    const { doc, s, image, choose } = await setup();
    s.host.editing.trackChanges = true;
    s.host.editing.author = 'Rina';
    choose(image, 'Delete');
    await flush();
    const header = doc.xmlSync(HEADER)!.root;
    const del = findAll(header, 'w:del');
    expect(del).toHaveLength(1);
    expect(del[0]!.attrs['w:author']).toBe('Rina');
    expect(findAll(del[0]!, 'w:drawing')).toHaveLength(1);
  });

  it('destroy() closes an open link prompt and its action does nothing', async () => {
    const { s, link, choose } = await setup();
    choose(link, 'Edit link…');
    await flush();
    expect(s.container.querySelector('.dk-popover')).not.toBeNull();
    s.ctl.destroy();
    expect(document.querySelector('.dk-popover')).toBeNull();
    await flush();
    expect(s.host.editing.transaction).not.toHaveBeenCalled();
  });
});
