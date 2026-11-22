import { describe, expect, it } from 'vitest';
import { DocxEditor } from '../src/api/index.js';
import { DocxDocument } from '../src/model/index.js';
import { findAll } from '../src/xml/index.js';
import { buildDocx, p } from './helpers/docx.js';

const EMU = 12700;
const pic = `<w:p><w:r><w:drawing><wp:inline><wp:extent cx="${50 * EMU}" cy="${50 * EMU}"/><wp:docPr id="1" name="Pic" descr="old"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:blipFill><a:blip r:embed="rIdX"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`;

async function open(body: string) {
  const doc = await DocxDocument.open(await buildDocx({ body }));
  return { doc, editor: new DocxEditor(doc) };
}

describe('setImageAlt', () => {
  it('writes descr on wp:docPr and survives save', async () => {
    const { doc, editor } = await open(pic);
    const drawing = findAll(doc.documentXml.root, 'w:drawing')[0]!;
    editor.setImageAlt(drawing, 'A red square');
    const again = await DocxDocument.open(await doc.save());
    expect(findAll(again.documentXml.root, 'wp:docPr')[0]!.attrs['descr']).toBe('A red square');
  });

  it('is refused by the guard', async () => {
    const doc = await DocxDocument.open(await buildDocx({ body: pic }));
    const editor = new DocxEditor(doc, {
      assert(action) {
        if (action === 'edit') throw new Error('Read only');
      },
    });
    expect(() => editor.setImageAlt(findAll(doc.documentXml.root, 'w:drawing')[0]!, 'x')).toThrow(/Read only/);
  });
});

describe('setHyperlinkTarget', () => {
  async function linked() {
    const { doc, editor } = await open(p('See the website now'));
    const link = await editor.addHyperlink(editor.find('website')[0]!, { url: 'https://example.com/a' });
    return { doc, editor, link };
  }

  it('points a link at another URL (a new relationship)', async () => {
    const { doc, editor, link } = await linked();
    const before = link.attrs['r:id'];
    await editor.setHyperlinkTarget(link, { url: 'https://example.org/b' });
    expect(link.attrs['r:id']).toBeTruthy();
    expect(link.attrs['r:id']).not.toBe(before);
    const again = new DocxEditor(await DocxDocument.open(await doc.save()));
    expect(again.hyperlinks().map((h) => h.url)).toEqual(['https://example.org/b']);
  });

  it('turns a URL link into a bookmark link', async () => {
    const { doc, editor, link } = await linked();
    await editor.setHyperlinkTarget(link, { bookmark: 'appendix' });
    expect(link.attrs['w:anchor']).toBe('appendix');
    expect(link.attrs['r:id']).toBeUndefined();
    const again = new DocxEditor(await DocxDocument.open(await doc.save()));
    expect(again.hyperlinks().map((h) => [h.url, h.anchor])).toEqual([[undefined, 'appendix']]);
  });

  it('turns a bookmark link back into a URL link, dropping w:anchor', async () => {
    const { editor, link } = await linked();
    await editor.setHyperlinkTarget(link, { bookmark: 'x' });
    await editor.setHyperlinkTarget(link, { url: 'mailto:a@b.co' });
    expect(link.attrs['w:anchor']).toBeUndefined();
    expect(editor.hyperlinks()[0]!.url).toBe('mailto:a@b.co');
  });

  it('refuses javascript: and leaves the link alone', async () => {
    const { editor, link } = await linked();
    const rid = link.attrs['r:id'];
    await expect(editor.setHyperlinkTarget(link, { url: 'javascript:alert(1)' })).rejects.toThrow(/scheme/);
    expect(link.attrs['r:id']).toBe(rid);
  });
});

describe('clearFormatting', () => {
  it('removes direct formatting from a range but keeps the text', async () => {
    const { doc, editor } = await open(p('Plain bold words', '', '<w:b/><w:color w:val="FF0000"/><w:sz w:val="40"/>'));
    editor.clearFormatting(editor.find('bold')[0]!);
    const runs = findAll(doc.documentXml.root, 'w:r');
    const formatted = runs.filter((r) => findAll(r, 'w:b').length);
    expect(formatted.map((r) => findAll(r, 'w:t')[0]!.children.map((c: any) => c.value).join(''))).toEqual(['Plain ', ' words']);
    expect(runs.map((r) => findAll(r, 'w:t')[0]!.children.map((c: any) => c.value).join('')).join('')).toBe('Plain bold words');
    const mid = runs.find((r) => findAll(r, 'w:t')[0]!.children.map((c: any) => c.value).join('') === 'bold')!;
    expect(findAll(mid, 'w:rPr')).toHaveLength(0);
  });
});

describe('relationships of replaced and removed links', () => {
  const relsXml = async (doc: DocxDocument) => {
    const again = await DocxDocument.open(await doc.save());
    return new TextDecoder().decode(await again.pkg.getBytes('word/_rels/document.xml.rels'));
  };
  const urls = (doc: DocxDocument) => [...doc.rels(doc.mainPart).values()].filter((r) => r.external).map((r) => r.target);

  async function twoLinks(sameUrl = false) {
    const { doc, editor } = await open(p('See the website and the other site'));
    const a = await editor.addHyperlink(editor.find('website')[0]!, { url: 'https://old.example/a' }, 'Old tip');
    const b = await editor.addHyperlink(editor.find('other site')[0]!, { url: sameUrl ? 'https://old.example/a' : 'https://keep.example/b' });
    return { doc, editor, a, b };
  }

  it('drops the old relationship, the old URL and the stale tooltip when a link changes URL', async () => {
    const { doc, editor, a } = await twoLinks();
    await editor.setHyperlinkTarget(a, { url: 'https://new.example/c' });
    expect(urls(doc).sort()).toEqual(['https://keep.example/b', 'https://new.example/c']);
    expect(a.attrs['w:tooltip']).toBeUndefined();
    const xml = await relsXml(doc);
    expect(xml).not.toContain('old.example');
    expect(xml).toContain('new.example');
  });

  it('keeps a relationship another hyperlink still uses', async () => {
    const { doc, editor, a, b } = await twoLinks();
    b.attrs['r:id'] = a.attrs['r:id']!;
    await editor.setHyperlinkTarget(a, { url: 'https://new.example/c' });
    expect(urls(doc)).toContain('https://old.example/a');
    expect(await relsXml(doc)).toContain('old.example');
  });

  it('drops the relationship when a link becomes a bookmark link', async () => {
    const { doc, editor, a } = await twoLinks();
    await editor.setHyperlinkTarget(a, { bookmark: 'top' });
    expect(await relsXml(doc)).not.toContain('old.example');
  });

  it('drops an unused relationship when a link is removed', async () => {
    const { doc, editor, a, b } = await twoLinks();
    await editor.removeHyperlink(a);
    expect(await relsXml(doc)).not.toContain('old.example');
    expect(await relsXml(doc)).toContain('keep.example');
    b.attrs['r:id'] = 'rIdNone';
  });
});
