/**
 * Drawings through the whole pipeline: Word markup → model → layout display list, for text-box
 * shapes (and their VML fallback), groups, charts, SmartArt and metafile pictures.
 */
import { describe, expect, it } from 'vitest';
import { layoutDocument, type DisplayItem, type ImageItem, type PathItem, type TextItem } from '../src/layout/index.js';
import { DocxDocument, type ImageInline } from '../src/model/index.js';
import { buildDocx, p } from './helpers/docx.js';
import { EMU, inline, textBoxRun, textBoxShape, vmlTextBoxRun } from './helpers/drawings.js';
import { testRegistry } from './helpers/fonts.js';

async function layout(body: string, extra: Parameters<typeof buildDocx>[0] = { body }) {
  const doc = await DocxDocument.open(await buildDocx({ ...extra, body }));
  const page = layoutDocument(doc, { fonts: testRegistry() }).pages[0]!;
  return { doc, page, items: page.items };
}

const paths = (items: DisplayItem[]) => items.filter((i): i is PathItem => i.type === 'path');
const texts = (items: DisplayItem[]) => items.filter((i): i is TextItem => i.type === 'text');

describe('shapes and text boxes', () => {
  it('draws a text box shape from mc:Choice, with its text and an arrowhead, and records it for editing', async () => {
    const { items, page } = await layout(`<w:p>${textBoxRun('Inside the box')}</w:p>`);
    const shape = paths(items).find((x) => x.fill === 'DDEEFF')!;
    expect(shape).toBeDefined();
    expect(shape.stroke).toBe('2E74B5');
    expect(shape.strokeWidth).toBe(2);
    // A rounded rectangle has curves.
    expect(shape.commands.some((c) => c.op === 'C')).toBe(true);
    const text = texts(items).find((t) => t.text.includes('Inside'))!;
    expect(text).toBeDefined();
    expect(texts(items).some((t) => t.text.includes('Old copy'))).toBe(false);
    // Centred vertically inside the 80pt box.
    const frame = items.find((i): i is ImageItem => i.type === 'image' && !!i.frameOnly)!;
    expect(text.y).toBeGreaterThan(frame.y + 30);
    expect(text.y).toBeLessThan(frame.y + 55);
    expect(page.storyLines.filter((l) => l.story === 'textbox')).toHaveLength(1);
  });

  it('reads an old VML text box', async () => {
    const { items } = await layout(`<w:p>${vmlTextBoxRun('VML text')}</w:p>`);
    expect(paths(items).find((x) => x.fill === 'FFF2CC')?.stroke).toBe('BF9000');
    expect(texts(items).map((t) => t.text).join(' ')).toContain('VML');
  });

  it('places group children in the group coordinate space, using theme colours from the shape style', async () => {
    const child = (x: number, prst: string) =>
      `<wps:wsp><wps:cNvSpPr/><wps:spPr><a:xfrm><a:off x="${x}" y="0"/><a:ext cx="1000" cy="1000"/></a:xfrm><a:prstGeom prst="${prst}"><a:avLst/></a:prstGeom></wps:spPr><wps:style><a:lnRef idx="2"><a:schemeClr val="accent1"><a:shade val="50000"/></a:schemeClr></a:lnRef><a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef><a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></wps:style><wps:bodyPr/></wps:wsp>`;
    const group = `<wpg:wgp><wpg:cNvGrpSpPr/><wpg:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${300 * EMU}" cy="${100 * EMU}"/><a:chOff x="0" y="0"/><a:chExt cx="3000" cy="1000"/></a:xfrm></wpg:grpSpPr>${child(0, 'ellipse')}${child(2000, 'triangle')}</wpg:wgp>`;
    const { items } = await layout(`<w:p><w:r>${inline(group, 300, 100, 'http://schemas.microsoft.com/office/word/2010/wordprocessingGroup')}</w:r></w:p>`);
    const shapes = paths(items).filter((x) => x.fill);
    expect(shapes).toHaveLength(2);
    // Office theme accent1, and its 50% shade for the outline.
    expect(shapes[0]!.fill).toBe('4472C4');
    expect(shapes[0]!.stroke).toMatch(/^[0-9A-F]{6}$/);
    expect(shapes[0]!.strokeWidth).toBe(1);
    const xs = (s: PathItem) => s.commands.filter((c) => c.op !== 'Z').map((c) => (c as { x: number }).x);
    expect(Math.min(...xs(shapes[0]!))).toBeCloseTo(72, 0);
    expect(Math.max(...xs(shapes[0]!))).toBeCloseTo(172, 0);
    expect(Math.min(...xs(shapes[1]!))).toBeCloseTo(272, 0);
  });

  it('rotates a shape about its centre', async () => {
    const rotated = textBoxShape('R', 'rect').replace('<a:xfrm>', '<a:xfrm rot="5400000">');
    const { doc, items } = await layout(`<w:p><w:r>${inline(rotated, 200, 80)}</w:r></w:p>`);
    const inl = doc.body().blocks[0]!.kind === 'paragraph' ? (doc.body().blocks[0] as { inlines: ImageInline[] }).inlines.find((i) => i.kind === 'image')! : undefined;
    expect(inl?.rotation).toBe(90);
    const shape = paths(items).find((x) => x.fill === 'DDEEFF')!;
    const pts = shape.commands.filter((c) => c.op !== 'Z') as Array<{ x: number; y: number }>;
    const w = Math.max(...pts.map((q) => q.x)) - Math.min(...pts.map((q) => q.x));
    const h = Math.max(...pts.map((q) => q.y)) - Math.min(...pts.map((q) => q.y));
    // A 200 × 80 box turned a quarter turn is 80 wide and 200 tall.
    expect(w).toBeCloseTo(80, 0);
    expect(h).toBeCloseTo(200, 0);
  });
});

describe('charts and SmartArt', () => {
  const chartXml = `<?xml version="1.0" encoding="UTF-8"?><c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart><c:title><c:tx><c:rich><a:bodyPr/><a:p><a:r><a:t>Sales</a:t></a:r></a:p></c:rich></c:tx></c:title><c:plotArea><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:v>2026</c:v></c:tx><c:cat><c:strLit><c:ptCount val="2"/><c:pt idx="0"><c:v>North</c:v></c:pt><c:pt idx="1"><c:v>South</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:ptCount val="2"/><c:pt idx="0"><c:v>10</c:v></c:pt><c:pt idx="1"><c:v>30</c:v></c:pt></c:numLit></c:val></c:ser><c:axId val="1"/><c:axId val="2"/></c:barChart><c:catAx><c:axId val="1"/><c:axPos val="b"/><c:crossAx val="2"/></c:catAx><c:valAx><c:axId val="2"/><c:axPos val="l"/><c:crossAx val="1"/></c:valAx></c:plotArea></c:chart></c:chartSpace>`;

  it('draws a chart from its part', async () => {
    const graphic = '<c:chart r:id="rIdC"/>';
    const { items } = await layout(`<w:p><w:r>${inline(graphic, 300, 200, 'http://schemas.openxmlformats.org/drawingml/2006/chart')}</w:r></w:p>`, {
      body: '',
      rels: [['rIdC', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart', 'charts/chart1.xml']],
      parts: { 'word/charts/chart1.xml': { data: chartXml, contentType: 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml' } },
    });
    expect(texts(items).some((t) => t.text === 'Sales')).toBe(true);
    expect(texts(items).some((t) => t.text === 'North')).toBe(true);
    // Two bars: the second three times as tall.
    const bars = paths(items).filter((x) => x.fill === '4472C4');
    expect(bars).toHaveLength(2);
    const height = (b: PathItem) => {
      const ys = b.commands.filter((c) => c.op !== 'Z').map((c) => (c as { y: number }).y);
      return Math.max(...ys) - Math.min(...ys);
    };
    expect(height(bars[1]!) / height(bars[0]!)).toBeCloseTo(3, 1);
  });

  it('draws SmartArt from its saved drawing part', async () => {
    const data = `<?xml version="1.0"?><dgm:dataModel xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram"><dgm:ptLst><dgm:pt modelId="1"><dgm:t><a:p><a:r><a:t>Plan</a:t></a:r></a:p></dgm:t></dgm:pt></dgm:ptLst><dgm:extLst><a:ext uri="http://schemas.microsoft.com/office/drawing/2008/diagram"><dsp:dataModelExt relId="rIdD" minVer="http://schemas.openxmlformats.org/drawingml/2006/diagram"/></a:ext></dgm:extLst></dgm:dataModel>`;
    const drawing = `<?xml version="1.0"?><dsp:drawing xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><dsp:spTree><dsp:nvGrpSpPr><dsp:cNvPr id="0" name=""/><dsp:cNvGrpSpPr/></dsp:nvGrpSpPr><dsp:grpSpPr/>` +
      `<dsp:sp modelId="{1}"><dsp:nvSpPr><dsp:cNvPr id="0" name=""/><dsp:cNvSpPr/></dsp:nvSpPr><dsp:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${120 * EMU}" cy="${60 * EMU}"/></a:xfrm><a:prstGeom prst="chevron"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="70AD47"/></a:solidFill></dsp:spPr><dsp:style><a:lnRef idx="0"><a:scrgbClr r="0" g="0" b="0"/></a:lnRef><a:fillRef idx="1"><a:scrgbClr r="0" g="0" b="0"/></a:fillRef><a:effectRef idx="0"><a:scrgbClr r="0" g="0" b="0"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></dsp:style><dsp:txBody><a:bodyPr anchor="ctr"/><a:p><a:pPr algn="ctr"/><a:r><a:rPr sz="1800"/><a:t>Plan</a:t></a:r></a:p></dsp:txBody><dsp:txXfrm><a:off x="0" y="0"/><a:ext cx="${120 * EMU}" cy="${60 * EMU}"/></dsp:txXfrm></dsp:sp></dsp:spTree></dsp:drawing>`;
    const graphic = '<dgm:relIds r:dm="rIdM" r:lo="rIdL" r:qs="rIdQ" r:cs="rIdS"/>';
    const { items } = await layout(`<w:p><w:r>${inline(graphic, 300, 100, 'http://schemas.openxmlformats.org/drawingml/2006/diagram')}</w:r></w:p>`, {
      body: '',
      rels: [
        ['rIdM', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData', 'diagrams/data1.xml'],
        ['rIdD', 'http://schemas.microsoft.com/office/2007/relationships/diagramDrawing', 'diagrams/drawing1.xml'],
      ],
      parts: {
        'word/diagrams/data1.xml': { data, contentType: 'application/vnd.openxmlformats-officedocument.drawingml.diagramData+xml' },
        'word/diagrams/drawing1.xml': { data: drawing, contentType: 'application/vnd.ms-office.drawingml.diagramDrawing+xml' },
      },
    });
    expect(paths(items).find((x) => x.fill === '70AD47')).toBeDefined();
    const label = texts(items).find((t) => t.text === 'Plan')!;
    expect(label.color).toBe('FFFFFF');
    expect(label.size).toBe(18);
  });
});

describe('metafile pictures', () => {
  /** A tiny EMF: a red filled rectangle over the whole frame. */
  function redEmf(): Uint8Array {
    const u32 = (...v: number[]) => v.flatMap((x) => [x & 0xff, (x >>> 8) & 0xff, (x >>> 16) & 0xff, (x >>> 24) & 0xff]);
    const rec = (type: number, body: number[]) => [...u32(type, 8 + body.length), ...body];
    const records = [
      rec(39, [...u32(1, 0), 0xff, 0, 0, 0, ...u32(0)]), // CREATEBRUSHINDIRECT #1 red
      rec(37, u32(1)), // SELECTOBJECT
      rec(37, u32(0x80000008)), // NULL_PEN
      rec(43, u32(0, 0, 100, 50)), // RECTANGLE
    ];
    const eof = rec(14, u32(0, 16, 20));
    const body = [...records.flat(), ...eof];
    const header = rec(1, [...u32(0, 0, 99, 49), ...u32(0, 0, 2500, 1250), ...u32(0x464d4520, 0x10000, 88 + body.length, records.length + 2), 8, 0, 0, 0, ...u32(0, 0, 0), ...u32(1000, 1000, 250, 250)]);
    return Uint8Array.from([...header, ...body]);
  }

  it('draws an EMF picture as vectors (and keeps its frame for the caret)', async () => {
    const pic = `<w:r><w:drawing><wp:inline><wp:extent cx="${100 * EMU}" cy="${50 * EMU}"/><wp:docPr id="1" name="P"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:blipFill><a:blip r:embed="rIdE"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;
    const { items } = await layout(`<w:p>${pic}</w:p>`, {
      body: '',
      rels: [['rIdE', 'image', 'media/image1.emf']],
      parts: { 'word/media/image1.emf': { data: redEmf(), contentType: 'image/x-emf' } },
    });
    const frame = items.find((i): i is ImageItem => i.type === 'image')!;
    expect(frame.frameOnly).toBe(true);
    const red = paths(items).find((x) => x.fill === 'FF0000')!;
    expect(red).toBeDefined();
    const xs = red.commands.filter((c) => c.op !== 'Z').map((c) => (c as { x: number }).x);
    expect(Math.min(...xs)).toBeCloseTo(frame.x, 0);
    expect(Math.max(...xs)).toBeCloseTo(frame.x + 100, 0);
  });
});
