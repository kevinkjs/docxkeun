import { describe, expect, it } from 'vitest';
import { layoutDocument } from '../src/layout/index.js';
import type { Page } from '../src/layout/types.js';
import { DocxDocument } from '../src/model/index.js';
import { defaultItems, hitTarget, type MenuContext, type MenuEnv, type MenuTarget } from '../src/render/context-menu.js';
import type { MenuItem } from '../src/render/ui/menu.js';
import { DocxEditor } from '../src/api/index.js';
import { linkSource, setLinkSource } from '../src/model/fields.js';
import { buildDocx, p, W_NS } from './helpers/docx.js';
import { testRegistry } from './helpers/fonts.js';

const fonts = testRegistry();
const EMU = 12700;
const HDR = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr ${W_NS}><w:p><w:r><w:t>Company header</w:t></w:r></w:p></w:hdr>`;
const SECT = '<w:sectPr><w:headerReference w:type="default" r:id="rIdH"/><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>';
const pic = `<w:p><w:r><w:drawing><wp:inline><wp:extent cx="${50 * EMU}" cy="${50 * EMU}"/><wp:docPr id="1" name="Pic" descr="alt"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:blipFill><a:blip r:embed="rIdX"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`;

async function laidOut(): Promise<Page> {
  const doc = await DocxDocument.open(
    await buildDocx({
      body: p('Some body text here.') + pic,
      sectPr: SECT,
      rels: [['rIdH', 'header', 'header1.xml']],
      parts: { 'word/header1.xml': { data: HDR, contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml' } },
    }),
  );
  return layoutDocument(doc, { fonts }).pages[0]!;
}

describe('hitTarget', () => {
  it('finds a header story line and an inline picture on a laid-out page', async () => {
    const page = await laidOut();
    const line = page.storyLines.find((l) => l.story === 'header')!;
    expect(hitTarget(page, line.x + 2, line.y + line.h / 2)).toMatchObject({ kind: 'header', part: line.part });
    const image = page.items.find((i) => i.type === 'image' && !i.frameOnly && i.inline)!;
    if (image.type !== 'image') throw new Error('no image');
    const t = hitTarget(page, image.x + image.w / 2, image.y + image.h / 2);
    expect(t.kind).toBe('image');
    expect(t.image).toBe(image.inline);
    expect(hitTarget(page, 300, 400).kind).toBe('text');
  });

  it('prefers a link over a comment over a picture, and ignores frame-only items', () => {
    const rect = { x: 10, y: 10, w: 100, h: 20 };
    const inline = { kind: 'image' } as never;
    const page = {
      links: [{ ...rect, target: { url: 'https://example.com' } }],
      comments: [{ ...rect, id: '7' }, { x: 0, y: 100, w: 0, h: 10, id: 'zero' }],
      items: [
        { type: 'image', x: 10, y: 10, w: 100, h: 100, inline },
        { type: 'image', x: 200, y: 200, w: 50, h: 50, inline, frameOnly: true },
      ],
      storyLines: [],
    } as unknown as Page;
    expect(hitTarget(page, 20, 15)).toMatchObject({ kind: 'link', link: { url: 'https://example.com' } });
    page.links = [];
    expect(hitTarget(page, 20, 15)).toMatchObject({ kind: 'comment', commentId: '7' });
    page.comments = [];
    expect(hitTarget(page, 20, 15)).toMatchObject({ kind: 'image', image: inline });
    expect(hitTarget(page, 220, 220).kind).toBe('text');
  });
});

const selection = (collapsed = false) => ({ story: 'body', collapsed, text: collapsed ? '' : 'abc', start: {}, end: {}, paragraphs: [], pages: [1] }) as never;
const noop = () => {};
const act = new Proxy({}, { get: () => noop }) as MenuEnv['act'];
const env = (extra: Partial<MenuEnv> = {}): MenuEnv => ({ editable: true, can: () => undefined, clipboardRead: true, act, ...extra });
const ctx = (target: MenuTarget, extra: Partial<MenuContext> = {}): MenuContext => ({ x: 0, y: 0, page: 0, selection: selection(), target, ...extra });
const labels = (items: MenuItem[]) => items.map((i) => (i === '-' ? '-' : i.label));
const find = (items: MenuItem[], label: string) => items.find((i): i is Exclude<MenuItem, '-'> => i !== '-' && i.label === label)!;

describe('defaultItems', () => {
  it('lists the groups for text when editing', () => {
    const items = defaultItems(ctx({ kind: 'text' }), env());
    expect(labels(items)).toEqual(['Cut', 'Copy', 'Paste', 'Paste as plain text', '-', 'Comment', 'Link…', 'Format', 'Paragraph']);
    expect(find(items, 'Cut').shortcut).toBe('⌘X');
    expect(find(items, 'Paste as plain text').shortcut).toBe('⇧⌘V');
    expect(labels(find(items, 'Format').items!)).toEqual(['Bold', 'Italic', 'Underline', 'Strikethrough', 'Superscript', 'Subscript', 'Clear formatting']);
    const para = labels(find(items, 'Paragraph').items!);
    expect(para).toEqual(expect.arrayContaining(['Align left', 'Align right', 'Bullets', 'Numbering', 'Increase indent', 'Decrease indent', 'Left to right', 'Right to left', 'Styles']));
  });

  it('adds the link group on a link', () => {
    const items = defaultItems(ctx({ kind: 'link', link: { url: 'https://example.com/x' } }), env());
    const l = labels(items);
    expect(l).toEqual(expect.arrayContaining(['https://example.com/x', 'Open link', 'Copy link address', 'Edit link…', 'Remove link']));
    expect(l.indexOf('https://example.com/x')).toBeLessThan(l.indexOf('Open link'));
    expect(find(items, 'https://example.com/x').disabled).toBe(true);
    expect(labels(defaultItems(ctx({ kind: 'link', link: { anchor: 'top' } }), env()))).toContain('#top');
  });

  it('greys Edit link… and Remove link on a link made by a HYPERLINK field', () => {
    const field = defaultItems(ctx({ kind: 'link', link: { url: 'https://example.com/f' } }), env());
    expect(find(field, 'Edit link…').disabled).toBe('This link cannot be edited here');
    expect(find(field, 'Remove link').disabled).toBe('This link cannot be edited here');
    expect(find(field, 'Open link').disabled).toBeFalsy();
    const element = setLinkSource({ url: 'https://example.com/h' }, {});
    const real = defaultItems(ctx({ kind: 'link', link: element }), env());
    expect(find(real, 'Edit link…').disabled).toBe(false);
    expect(find(real, 'Remove link').disabled).toBe(false);
    // A guard refusal still wins.
    const refused = defaultItems(ctx({ kind: 'link', link: element }), env({ can: () => 'Locked' }));
    expect(find(refused, 'Edit link…').disabled).toBe('Locked');
  });

  it('adds comment, picture and header/footer groups', () => {
    const c = defaultItems(ctx({ kind: 'comment', commentId: '3' }), env({ commentInfo: () => ({ author: 'Ada', resolved: false }) }));
    expect(labels(c)).toEqual(expect.arrayContaining(['Comment by Ada', 'Reply', 'Resolve', 'Delete comment']));
    const r = defaultItems(ctx({ kind: 'comment', commentId: '3' }), env({ commentInfo: () => ({ author: 'Ada', resolved: true }) }));
    expect(labels(r)).toContain('Reopen');
    expect(labels(defaultItems(ctx({ kind: 'image', image: {} as never }), env()))).toEqual(expect.arrayContaining(['Alt text…', 'Copy picture', 'Delete']));
    expect(labels(defaultItems(ctx({ kind: 'header', part: 'word/header1.xml' }), env()))).toEqual(expect.arrayContaining(['Edit header/footer', 'Remove header/footer']));
  });

  it('drops Cut and Paste and disables Comment when read-only', () => {
    const items = defaultItems(ctx({ kind: 'text' }), env({ editable: false, can: () => 'The document is read-only' }));
    const l = labels(items);
    expect(l).not.toContain('Cut');
    expect(l).not.toContain('Paste');
    expect(l).toContain('Copy');
    expect(find(items, 'Comment').disabled).toBe('The document is read-only');
    expect(find(items, 'Format').disabled).toBeTruthy();
  });

  it("disables with the guard's reason", () => {
    const items = defaultItems(ctx({ kind: 'text' }), env({ can: (a) => (a === 'comment' ? 'The document only allows comments' : 'Locked') }));
    expect(find(items, 'Comment').disabled).toBe('The document only allows comments');
    expect(find(items, 'Cut').disabled).toBe('Locked');
    expect(find(items, 'Link…').disabled).toBe('Locked');
  });

  it('allows Comment where the guard only allows comments', () => {
    const items = defaultItems(ctx({ kind: 'text' }), env({ can: (a) => (a === 'comment' ? undefined : 'Only comments are allowed') }));
    expect(find(items, 'Comment').disabled).toBeFalsy();
    expect(find(items, 'Cut').disabled).toBe('Only comments are allowed');
  });

  it('shows Paste (⌘V) without the async clipboard', () => {
    const items = defaultItems(ctx({ kind: 'text' }), env({ clipboardRead: false }));
    expect(labels(items)).toContain('Paste (⌘V)');
    expect(labels(items)).not.toContain('Paste as plain text');
  });

  it('asks for a selection before Comment, Link and Copy', () => {
    const items = defaultItems(ctx({ kind: 'text' }, { selection: selection(true) }), env());
    expect(find(items, 'Comment').disabled).toBeTruthy();
    expect(find(items, 'Link…').disabled).toBeTruthy();
    expect(find(items, 'Copy').disabled).toBeTruthy();
  });
});

describe('the clicked link', () => {
  it('is the w:hyperlink the layout area came from, so edits hit that link only', async () => {
    const doc = await DocxDocument.open(await buildDocx({ body: p('First link and second link here') }));
    const editor = new DocxEditor(doc);
    const a = await editor.addHyperlink(editor.find('First link')[0]!, { url: 'https://one.example/' });
    const b = await editor.addHyperlink(editor.find('second link')[0]!, { url: 'https://two.example/' });
    const page = layoutDocument(doc, { fonts }).pages[0]!;
    const areas = page.links.map((l) => ({ url: l.target.url, source: linkSource(l.target) }));
    expect(areas).toEqual([
      { url: 'https://one.example/', source: a },
      { url: 'https://two.example/', source: b },
    ]);
    const second = page.links[1]!;
    const t = hitTarget(page, second.x + 1, second.y + 1);
    expect(linkSource(t.link!)).toBe(b);
    await editor.setHyperlinkTarget(linkSource(t.link!) as never, { url: 'https://three.example/' });
    expect(editor.hyperlinks().map((h) => h.url)).toEqual(['https://one.example/', 'https://three.example/']);
    await editor.removeHyperlink(linkSource(t.link!) as never);
    expect(editor.hyperlinks().map((h) => h.url)).toEqual(['https://one.example/']);
  });
});
