import { describe, expect, it } from 'vitest';
import { layoutDocument, type TextItem } from '../src/layout/index.js';
import { DocxDocument } from '../src/model/index.js';
import { pageToSvg } from '../src/render/svg.js';
import { buildDocx, p } from './helpers/docx.js';
import { testRegistry } from './helpers/fonts.js';

const RTL = '<w:bidi/>';

async function lineChars(body: string) {
  const doc = await DocxDocument.open(await buildDocx({ body }));
  const page = layoutDocument(doc, { fonts: testRegistry() }).pages[0]!;
  const texts = page.items.filter((i): i is TextItem => i.type === 'text');
  // Every character with its page x, left to right.
  const chars = texts.flatMap((t) => t.chars.map((c, i) => ({ c, x: t.xs[i]!, w: t.widths[i]! }))).sort((a, b) => a.x - b.x);
  return { page, texts, chars, visual: chars.map((c) => c.c).join('') };
}

describe('right-to-left text', () => {
  it('lays out a right-to-left paragraph from the right margin', async () => {
    const { chars, visual } = await lineChars(p('שלום עולם', RTL));
    // Visual order is the reverse of the logical order.
    expect(visual).toBe('םלוע םולש');
    const last = chars[chars.length - 1]!;
    // Default (start) alignment in a right-to-left paragraph is the right margin.
    expect(last.x + last.w).toBeCloseTo(612 - 72, 1);
  });

  it('keeps left-to-right words and numbers in order inside right-to-left text', async () => {
    const { visual } = await lineChars(p('שלום abc 123 עולם', RTL));
    expect(visual).toBe('םלוע abc 123 םולש');
  });

  it('reverses a right-to-left word inside a left-to-right paragraph', async () => {
    const { visual, chars } = await lineChars(p('Hello שלום world'));
    expect(visual).toBe('Hello םולש world');
    expect(chars[0]!.x).toBeCloseTo(72, 1);
  });

  it('mirrors brackets in right-to-left text', async () => {
    const { visual } = await lineChars(p('(שלום)', RTL));
    // Logical "(" ends up on the right, drawn as ")" turned around: the pair reads (םולש) on the page.
    expect(visual).toBe('(םולש)');
  });

  it('applies indents from the right in a right-to-left paragraph', async () => {
    const { chars } = await lineChars(p('שלום', `${RTL}<w:ind w:left="1440"/>`));
    const last = chars[chars.length - 1]!;
    expect(last.x + last.w).toBeCloseTo(612 - 72 - 72, 1);
  });

  it('draws right-to-left runs in logical order in the SVG, for the browser to shape', async () => {
    const { page, texts } = await lineChars(p('مرحبا بالعالم', RTL));
    expect(texts.every((t) => t.complex?.rtl)).toBe(true);
    const svg = pageToSvg(page, { fontFamily: (f) => f.family, imageUrl: () => undefined });
    expect(svg).toContain('direction="rtl"');
    expect(svg).toContain('>مرحبا بالعالم</text>');
  });
});

import { existsSync } from 'node:fs';
import { exportPdf } from '../src/pdf/index.js';
import { pageText, pages, parsePdf } from './helpers/pdf.js';
import { readFont } from './helpers/fonts.js';

describe('shaping in layout', () => {
  it('kerns Latin text when Word asks for it (w:kern)', async () => {
    const width = async (rPr: string) => {
      const { texts } = await lineChars(`<w:p><w:r><w:rPr><w:rFonts w:ascii="Liberation Sans" w:hAnsi="Liberation Sans"/><w:sz w:val="48"/>${rPr}</w:rPr><w:t>AVAVAV</w:t></w:r></w:p>`);
      const t = texts[0]!;
      return t.xs[t.xs.length - 1]! + t.widths[t.widths.length - 1]! - t.xs[0]!;
    };
    const plain = await width('');
    const kerned = await width('<w:kern w:val="16"/>');
    expect(kerned).toBeLessThan(plain - 3);
  });

  const arial = '/System/Library/Fonts/Supplemental/Arial.ttf';
  it.skipIf(!existsSync(arial))('joins Arabic letters (and the lam-alef ligature) in PDFs', async () => {
    const fonts = testRegistry();
    fonts.add(readFont(arial), { family: 'Arial', bold: false, italic: false });
    const body = `<w:p><w:pPr><w:bidi/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:rtl/></w:rPr><w:t>سلام</w:t></w:r></w:p>`;
    const doc = await DocxDocument.open(await buildDocx({ body }));
    const layout = layoutDocument(doc, { fonts });
    const item = layout.pages[0]!.items.find((i): i is TextItem => i.type === 'text')!;
    expect(item.complex?.rtl).toBe(true);
    const glyphs = item.clusterGlyphs!.flat();
    // سلام: initial seen, the lam-alef ligature, final meem: three glyphs for four letters.
    expect(glyphs).toHaveLength(3);
    const font = item.font.face!.font;
    expect(glyphs.map((g) => g.id)).not.toContain(font.glyphId('س'.codePointAt(0)!));
    expect(glyphs.find((g) => g.text === 'لا')).toBeDefined();
    // The PDF keeps the text searchable through ToUnicode.
    const pdf = parsePdf(await exportPdf(doc, { fonts, layout }));
    expect(pageText(pdf, pages(pdf)[0]!).replace(/\s/g, '')).toMatch(/[سلام]{4}/);
  });
});
