import { describe, expect, it } from 'vitest';
import { layoutDocument, type ImageItem, type TextItem } from '../src/layout/index.js';
import { DocxDocument } from '../src/model/index.js';
import { buildDocx, p } from './helpers/docx.js';
import { testRegistry } from './helpers/fonts.js';

const EMU = 12700;

/** A floating picture anchored in a paragraph, at (x, y) from the column / paragraph. */
function floating(options: { wrap: string; x?: number; y?: number; w?: number; h?: number; dist?: number; align?: string }): string {
  const { w = 120, h = 90, x = 0, y = 0, dist = 9 } = options;
  const posH = options.align ? `<wp:align>${options.align}</wp:align>` : `<wp:posOffset>${x * EMU}</wp:posOffset>`;
  return (
    `<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="${dist * EMU}" distR="${dist * EMU}" simplePos="0" relativeHeight="1" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">` +
    `<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column">${posH}</wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>${y * EMU}</wp:posOffset></wp:positionV>` +
    `<wp:extent cx="${w * EMU}" cy="${h * EMU}"/>${options.wrap}<wp:docPr id="1" name="Picture 1"/>` +
    `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:blipFill><a:blip r:embed="rIdMissing"/></pic:blipFill></pic:pic></a:graphicData></a:graphic>` +
    `</wp:anchor></w:drawing></w:r>`
  );
}

const words = (n: number, prefix = 'w') => Array.from({ length: n }, (_, i) => `${prefix}${i}`).join(' ');

async function page(body: string) {
  const doc = await DocxDocument.open(await buildDocx({ body }));
  const result = layoutDocument(doc, { fonts: testRegistry() }).pages[0]!;
  const texts = result.items.filter((i): i is TextItem => i.type === 'text');
  const image = result.items.find((i): i is ImageItem => i.type === 'image');
  return { texts, image, lines: result.lines };
}

describe('text wrapping around floating objects', () => {
  it('flows text beside a square-wrapped picture, then back to the full width', async () => {
    const { lines, image } = await page(`<w:p>${floating({ wrap: '<wp:wrapSquare wrapText="bothSides"/>' })}<w:r><w:t>${words(160)}</w:t></w:r></w:p>`);
    expect(image).toMatchObject({ x: 72, y: 72, w: 120, h: 90 });
    const beside = lines.filter((l) => l.y < 72 + 90);
    const below = lines.filter((l) => l.y >= 72 + 90 + 1);
    expect(beside.length).toBeGreaterThan(3);
    // Lines beside the picture start after it and its distance from text.
    for (const l of beside) expect(l.x).toBeCloseTo(72 + 120 + 9, 1);
    expect(below.length).toBeGreaterThan(0);
    for (const l of below) expect(l.x).toBeCloseTo(72, 1);
  });

  it('puts text on the left of a picture aligned right', async () => {
    const { lines, image } = await page(`<w:p>${floating({ wrap: '<wp:wrapSquare wrapText="bothSides"/>', align: 'right' })}<w:r><w:t>${words(120)}</w:t></w:r></w:p>`);
    expect(image!.x + image!.w).toBeCloseTo(612 - 72, 1);
    const beside = lines.filter((l) => l.y < 72 + 90);
    for (const l of beside) {
      expect(l.x).toBeCloseTo(72, 1);
      expect(l.x + l.w).toBeLessThanOrEqual(image!.x - 9 + 0.5);
    }
  });

  it('wraps the following paragraphs too, as long as they meet the picture', async () => {
    const body = `<w:p>${floating({ wrap: '<wp:wrapSquare wrapText="bothSides"/>', h: 150 })}<w:r><w:t>Short.</w:t></w:r></w:p>` + p(words(40, 'a')) + p(words(200, 'b'));
    const { lines } = await page(body);
    const second = lines.filter((l) => l.y > 80 && l.y < 72 + 150);
    expect(second.length).toBeGreaterThan(1);
    for (const l of second) expect(l.x).toBeGreaterThan(190);
    expect(lines.some((l) => l.y > 72 + 160 && Math.abs(l.x - 72) < 0.5)).toBe(true);
  });

  it('moves text below a top-and-bottom picture', async () => {
    const { texts, image } = await page(`<w:p>${floating({ wrap: '<wp:wrapTopAndBottom/>' })}<w:r><w:t>Below the picture</w:t></w:r></w:p>`);
    const first = texts.find((t) => t.text.startsWith('Below'))!;
    expect(first.y).toBeGreaterThan(image!.y + image!.h);
    expect(first.x).toBeCloseTo(72, 1);
  });

  it('leaves text alone for pictures that do not wrap', async () => {
    const { lines } = await page(`<w:p>${floating({ wrap: '<wp:wrapNone/>' })}<w:r><w:t>${words(60)}</w:t></w:r></w:p>`);
    for (const l of lines) expect(l.x).toBeCloseTo(72, 1);
  });
});
