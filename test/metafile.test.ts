/**
 * EMF and WMF playback: small metafiles written record by record here, played into display
 * items, with assertions on what comes out (shapes, colours, coordinates, text, bitmaps).
 */
import { describe, expect, it } from 'vitest';
import { RegistryFontProvider } from '../src/layout/fonts.js';
import type { DisplayItem, ImageItem, LineItem, PathCommand, PathItem, TextItem } from '../src/layout/types.js';
import { metafileKind, metafileSize, renderMetafile } from '../src/metafile/index.js';
import { testRegistry } from './helpers/fonts.js';

const fonts = new RegistryFontProvider(testRegistry());

/** Little-endian byte writer. Negative values wrap to two's complement. */
class W {
  readonly bytes: number[] = [];
  u8(...v: number[]): this {
    for (const x of v) this.bytes.push(x & 0xff);
    return this;
  }
  u16(...v: number[]): this {
    for (const x of v) this.bytes.push(x & 0xff, (x >> 8) & 0xff);
    return this;
  }
  u32(...v: number[]): this {
    for (const x of v) this.bytes.push(x & 0xff, (x >>> 8) & 0xff, (x >>> 16) & 0xff, (x >>> 24) & 0xff);
    return this;
  }
  f32(...v: number[]): this {
    for (const x of v) this.bytes.push(...new Uint8Array(new Float32Array([x]).buffer));
    return this;
  }
  utf16(s: string): this {
    for (let i = 0; i < s.length; i++) this.u16(s.charCodeAt(i));
    return this;
  }
  ascii(s: string): this {
    for (let i = 0; i < s.length; i++) this.u8(s.charCodeAt(i));
    return this;
  }
  raw(b: ArrayLike<number>): this {
    for (let i = 0; i < b.length; i++) this.bytes.push(b[i]!);
    return this;
  }
  pad(n: number): this {
    while (this.bytes.length % n) this.bytes.push(0);
    return this;
  }
}

// ---- EMF writing ----

function rec(type: number, build?: (w: W) => void): number[] {
  const body = new W();
  build?.(body);
  body.pad(4);
  return new W().u32(type, 8 + body.bytes.length).raw(body.bytes).bytes;
}

/**
 * An EMF whose reference device has 4 pixels per mm (1000 px over 250 mm) and whose frame is
 * 25 × 12.5 mm, i.e. 100 × 50 device pixels.
 */
function emf(records: number[][], frame: [number, number, number, number] = [0, 0, 2500, 1250]): Uint8Array {
  const eof = rec(14, (w) => w.u32(0, 16, 20));
  const body = [...records.flat(), ...eof];
  const header = rec(1, (w) => {
    w.u32(0, 0, 99, 49); // bounds (device pixels)
    w.u32(...frame); // frame (0.01 mm)
    w.u32(0x464d4520, 0x10000, 88 + body.length, records.length + 2);
    w.u16(8, 0); // handles, reserved
    w.u32(0, 0, 0); // description, palette
    w.u32(1000, 1000, 250, 250); // device size in pixels, then in mm
  });
  return Uint8Array.from([...header, ...body]);
}

const bounds = (w: W): W => w.u32(0, 0, 0, 0);
const createPen = (ih: number, style: number, width: number, rgb: [number, number, number]): number[] => rec(38, (w) => w.u32(ih, style, width, 0).u8(...rgb, 0));
const createBrush = (ih: number, style: number, rgb: [number, number, number], hatch = 0): number[] => rec(39, (w) => w.u32(ih, style).u8(...rgb, 0).u32(hatch));
const select = (ih: number): number[] => rec(37, (w) => w.u32(ih));
const rectangle = (l: number, t: number, r: number, b: number): number[] => rec(43, (w) => w.u32(l, t, r, b));
const moveTo = (x: number, y: number): number[] => rec(27, (w) => w.u32(x, y));
const lineTo = (x: number, y: number): number[] => rec(54, (w) => w.u32(x, y));
const xformRec = (type: number, m: number[], mode?: number): number[] =>
  rec(type, (w) => {
    w.f32(...m);
    if (mode !== undefined) w.u32(mode);
  });

function extTextOutW(x: number, y: number, text: string, dx?: number[], options = 0): number[] {
  return rec(84, (w) => {
    bounds(w).u32(1).f32(1, 1);
    const offString = 76;
    const offDx = dx ? offString + Math.ceil((text.length * 2) / 4) * 4 : 0;
    w.u32(x, y, text.length, offString, options, 0, 0, 0, 0, offDx);
    w.utf16(text).pad(4);
    if (dx) w.u32(...dx);
  });
}

function logFontW(ih: number, height: number, weight: number, face: string, italic = 0): number[] {
  return rec(82, (w) => {
    w.u32(ih, height, 0, 0, 0, weight).u8(italic, 0, 0, 0, 0, 0, 0, 0);
    const name = new W().utf16(face);
    w.raw(name.bytes);
    for (let i = name.bytes.length; i < 64; i++) w.u8(0);
  });
}

// ---- WMF writing ----

function wrec(fn: number, build?: (w: W) => void): number[] {
  const body = new W();
  build?.(body);
  body.pad(2);
  return new W().u32(3 + body.bytes.length / 2).u16(fn).raw(body.bytes).bytes;
}

function wmf(records: number[][], placeable?: { box: [number, number, number, number]; inch: number }): Uint8Array {
  const body = [...records.flat(), ...wrec(0)];
  const out = new W();
  if (placeable) {
    const head = new W().u32(0x9ac6cdd7).u16(0, ...placeable.box, placeable.inch).u32(0);
    let sum = 0;
    for (let i = 0; i < head.bytes.length; i += 2) sum ^= head.bytes[i]! | (head.bytes[i + 1]! << 8);
    out.raw(head.bytes).u16(sum);
  }
  out.u16(1, 9, 0x300).u32(9 + body.length / 2).u16(8).u32(64).u16(0);
  return Uint8Array.from([...out.bytes, ...body]);
}

const wPen = (style: number, width: number, rgb: [number, number, number]): number[] => wrec(0x02fa, (w) => w.u16(style, width, 0).u8(...rgb, 0));
const wBrush = (style: number, rgb: [number, number, number]): number[] => wrec(0x02fc, (w) => w.u16(style).u8(...rgb, 0).u16(0));
const wSelect = (i: number): number[] => wrec(0x012d, (w) => w.u16(i));
const wDelete = (i: number): number[] => wrec(0x01f0, (w) => w.u16(i));
const wFont = (height: number, face: string, weight = 400): number[] =>
  wrec(0x02fb, (w) => {
    w.u16(height, 0, 0, 0, weight).u8(0, 0, 0, 0, 0, 0, 0, 0).ascii(face).u8(0);
  });

// ---- assertions ----

const paths = (items: DisplayItem[]): PathItem[] => items.filter((i): i is PathItem => i.type === 'path');
const texts = (items: DisplayItem[]): TextItem[] => items.filter((i): i is TextItem => i.type === 'text');

function points(cmds: PathCommand[]): Array<[number, number]> {
  return cmds.flatMap((c): Array<[number, number]> => (c.op === 'Z' ? [] : [[c.x, c.y]]));
}

function expectPoints(cmds: PathCommand[], expected: Array<[number, number]>): void {
  const actual = points(cmds);
  expect(actual.length).toBe(expected.length);
  actual.forEach(([x, y], i) => {
    expect(x).toBeCloseTo(expected[i]![0], 3);
    expect(y).toBeCloseTo(expected[i]![1], 3);
  });
}

/** The EMF test frame: device pixel (x, y) lands on (10 + 2x, 20 + 2y). */
const FRAME = { x: 10, y: 20, w: 200, h: 100 };

describe('metafile detection and size', () => {
  it('recognises EMF, placeable and standard WMF, and nothing else', () => {
    expect(metafileKind(emf([]))).toBe('emf');
    expect(metafileKind(wmf([], { box: [0, 0, 100, 100], inch: 1440 }))).toBe('wmf');
    expect(metafileKind(wmf([]))).toBe('wmf');
    expect(metafileKind(Uint8Array.of(0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a))).toBeUndefined();
    expect(metafileKind(new Uint8Array(200))).toBeUndefined();
  });

  it('takes the EMF size from the frame in 0.01 mm', () => {
    const size = metafileSize(emf([]))!;
    expect(size.width).toBeCloseTo((25 * 72) / 25.4, 3);
    expect(size.height).toBeCloseTo((12.5 * 72) / 25.4, 3);
  });

  it('takes the WMF size from the placeable box, else from the window extent in twips', () => {
    expect(metafileSize(wmf([], { box: [0, 0, 1000, 500], inch: 1000 }))).toEqual({ width: 72, height: 36 });
    const standard = wmf([wrec(0x020b, (w) => w.u16(0, 0)), wrec(0x020c, (w) => w.u16(720, 2880))]);
    expect(metafileSize(standard)).toEqual({ width: 144, height: 36 });
  });

  it('survives truncated and nonsense data', () => {
    const good = emf([createBrush(1, 0, [255, 0, 0]), select(1), rectangle(0, 0, 10, 10)]);
    for (let n = 0; n < good.length; n += 7) expect(() => renderMetafile(good.subarray(0, n), FRAME, fonts)).not.toThrow();
    const bogus = emf([rec(3, (w) => bounds(w).u32(0xffffffff, 1, 2))]);
    expect(renderMetafile(bogus, FRAME, fonts)).toEqual([]);
    expect(renderMetafile(Uint8Array.of(1, 2, 3), FRAME, fonts)).toEqual([]);
  });
});

describe('EMF', () => {
  it('fills a rectangle with the brush and strokes it with the pen', () => {
    const items = renderMetafile(emf([createPen(1, 0, 3, [255, 0, 0]), createBrush(2, 0, [0, 0, 255]), select(1), select(2), rectangle(10, 5, 60, 45)]), FRAME, fonts);
    expect(items).toHaveLength(1);
    const p = items[0] as PathItem;
    expect(p.type).toBe('path');
    expect(p.fill).toBe('0000FF');
    expect(p.stroke).toBe('FF0000');
    expect(p.strokeWidth).toBeCloseTo(6, 6);
    expect(p.evenOdd).toBe(true);
    expectPoints(p.commands, [
      [30, 30],
      [130, 30],
      [130, 110],
      [30, 110],
    ]);
    expect(p.commands[p.commands.length - 1]!.op).toBe('Z');
  });

  it('draws the default pen as a hairline, and stock NULL_PEN / NULL_BRUSH as nothing', () => {
    const [plain] = paths(renderMetafile(emf([rectangle(0, 0, 10, 10)]), FRAME, fonts));
    expect(plain!.fill).toBe('FFFFFF');
    expect(plain!.stroke).toBe('000000');
    expect(plain!.strokeWidth).toBe(0.25);
    const fillOnly = paths(renderMetafile(emf([select(0x80000008), rectangle(0, 0, 10, 10)]), FRAME, fonts));
    expect(fillOnly[0]!.stroke).toBeUndefined();
    const none = renderMetafile(emf([select(0x80000008), select(0x80000005), rectangle(0, 0, 10, 10)]), FRAME, fonts);
    expect(none).toEqual([]);
  });

  it('reads polygons in their 16-bit and 32-bit forms alike', () => {
    const pts = [
      [0, 0],
      [40, 0],
      [20, 30],
    ];
    const p32 = rec(3, (w) => bounds(w).u32(3, ...pts.flat()));
    const p16 = rec(86, (w) => bounds(w).u32(3).u16(...pts.flat()));
    const [a] = paths(renderMetafile(emf([p32]), FRAME, fonts));
    const [b] = paths(renderMetafile(emf([p16]), FRAME, fonts));
    expect(a!.commands).toEqual(b!.commands);
    expectPoints(a!.commands, [
      [10, 20],
      [90, 20],
      [50, 80],
    ]);
    expect(a!.commands.map((c) => c.op)).toEqual(['M', 'L', 'L', 'Z']);
    // Negative 16-bit coordinates stay negative.
    const neg = rec(87, (w) => bounds(w).u32(2).u16(-5, -5, 5, 5));
    expectPoints(paths(renderMetafile(emf([neg]), FRAME, fonts))[0]!.commands, [
      [0, 10],
      [20, 30],
    ]);
  });

  it('turns an ellipse into four Béziers', () => {
    const [e] = paths(renderMetafile(emf([rec(42, (w) => w.u32(0, 0, 100, 50))]), FRAME, fonts));
    expect(e!.commands.map((c) => c.op)).toEqual(['M', 'C', 'C', 'C', 'C', 'Z']);
    // Centre (110, 70), radii 100 × 50 in points.
    expectPoints(e!.commands, [
      [210, 70],
      [110, 120],
      [10, 70],
      [110, 20],
      [210, 70],
    ]);
    const c1 = e!.commands[1] as Extract<PathCommand, { op: 'C' }>;
    expect(c1.x1).toBeCloseTo(210, 6);
    expect(c1.y1).toBeCloseTo(70 + 50 * 0.5523, 2);
  });

  it('draws arcs counterclockwise by default and clockwise after SetArcDirection', () => {
    // From the right-hand point to the top point of the circle in the box (0,0)-(100,100).
    const arc = rec(45, (w) => w.u32(0, 0, 100, 100, 100, 50, 50, 0));
    const [ccw] = paths(renderMetafile(emf([arc]), { x: 0, y: 0, w: 100, h: 50 }, fonts));
    expect(ccw!.commands.map((c) => c.op)).toEqual(['M', 'C']);
    const mid = (cmds: PathCommand[]): number => (cmds[1] as Extract<PathCommand, { op: 'C' }>).y1;
    expectPoints(ccw!.commands, [
      [100, 50],
      [50, 0],
    ]);
    expect(mid(ccw!.commands)).toBeLessThan(50); // via the upper right
    const [cw] = paths(renderMetafile(emf([rec(57, (w) => w.u32(2)), arc]), { x: 0, y: 0, w: 100, h: 50 }, fonts));
    expect(cw!.commands.filter((c) => c.op === 'C')).toHaveLength(3); // three quarters, the long way round
    const pie = rec(47, (w) => w.u32(0, 0, 100, 100, 100, 50, 50, 0));
    const [p] = paths(renderMetafile(emf([pie]), { x: 0, y: 0, w: 100, h: 50 }, fonts));
    expect(p!.commands.map((c) => c.op)).toEqual(['M', 'L', 'C', 'Z']);
    expectPoints(p!.commands.slice(0, 1), [[50, 50]]);
  });

  it('draws round rectangles, angle arcs and hatched brushes', () => {
    const one = { x: 0, y: 0, w: 100, h: 50 };
    const [rr] = paths(renderMetafile(emf([rec(44, (w) => w.u32(0, 0, 100, 50, 20, 10))]), one, fonts));
    expect(rr!.commands.map((c) => c.op)).toEqual(['M', 'L', 'C', 'L', 'C', 'L', 'C', 'L', 'C', 'Z']);
    expectPoints(rr!.commands.slice(0, 3), [
      [10, 0],
      [90, 0],
      [100, 5],
    ]);
    // AngleArc: a line from the current position to the arc start, a quarter turn counterclockwise, then LineTo carries on.
    const angle = rec(41, (w) => w.u32(50, 50, 50).f32(0, 90));
    const [aa] = paths(renderMetafile(emf([moveTo(0, 50), angle, lineTo(0, 0)]), one, fonts));
    expect(aa!.commands.map((c) => c.op)).toEqual(['M', 'L', 'C', 'L']);
    expectPoints(aa!.commands, [
      [0, 50],
      [100, 50],
      [50, 0],
      [0, 0],
    ]);
    const hatched = createBrush(1, 2, [0, 0, 0], 5);
    const [opaque] = paths(renderMetafile(emf([hatched, select(1), rectangle(0, 0, 10, 10)]), one, fonts));
    expect(opaque!.fill).toBe('A6A6A6'); // 35% black over the white background colour
    const [clear] = paths(renderMetafile(emf([hatched, select(1), rec(18, (w) => w.u32(1)), rectangle(0, 0, 10, 10)]), one, fonts));
    expect([clear!.fill, clear!.fillOpacity]).toEqual(['000000', 0.35]);
  });

  it('records a path with Béziers and fills and strokes it once', () => {
    const items = renderMetafile(
      emf([
        createBrush(1, 0, [0, 128, 0]),
        select(1),
        rec(59),
        moveTo(0, 0),
        rec(88, (w) => bounds(w).u32(3).u16(10, 0, 20, 10, 20, 20)),
        lineTo(0, 20),
        rec(61),
        rec(60),
        rec(63),
      ]),
      FRAME,
      fonts,
    );
    expect(items).toHaveLength(1);
    const p = items[0] as PathItem;
    expect(p.commands.map((c) => c.op)).toEqual(['M', 'C', 'L', 'Z']);
    expect(p.fill).toBe('008000');
    expect(p.stroke).toBe('000000');
    const c = p.commands[1] as Extract<PathCommand, { op: 'C' }>;
    expect([c.x1, c.y1, c.x2, c.y2, c.x, c.y]).toEqual([30, 20, 50, 40, 50, 60]);
  });

  it('keeps successive LineTo calls in one stroked path', () => {
    const items = renderMetafile(emf([moveTo(0, 0), lineTo(10, 0), lineTo(10, 10), lineTo(0, 10)]), FRAME, fonts);
    expect(items).toHaveLength(1);
    const p = items[0] as PathItem;
    expect(p.fill).toBeUndefined();
    expect(p.commands.map((c) => c.op)).toEqual(['M', 'L', 'L', 'L']);
  });

  it('applies geometric dashed pens from ExtCreatePen', () => {
    // PS_GEOMETRIC | PS_DASH | PS_ENDCAP_FLAT, width 2, BS_SOLID.
    const pen = rec(95, (w) => w.u32(1, 0, 0, 0, 0, 0x10000 | 0x200 | 1, 2, 0).u8(0, 0, 255, 0).u32(0, 0));
    const [p] = paths(renderMetafile(emf([pen, select(1), moveTo(0, 0), lineTo(50, 0)]), FRAME, fonts));
    expect(p!.stroke).toBe('0000FF');
    expect(p!.strokeWidth).toBe(4);
    expect(p!.dash).toEqual([12, 4]);
    expect(p!.lineCap).toBe('butt');
  });

  it('places text by alignment, with Dx advances, bold font and colour', () => {
    const items = renderMetafile(
      emf([
        logFontW(1, -10, 700, 'Arial'),
        select(1),
        rec(24, (w) => w.u8(0x11, 0x22, 0x33, 0)),
        rec(22, (w) => w.u32(6 | 24)), // TA_CENTER | TA_BASELINE
        extTextOutW(50, 25, 'Hi', [10, 12]),
      ]),
      FRAME,
      fonts,
    );
    const [t] = texts(items);
    expect(t!.text).toBe('Hi');
    expect(t!.size).toBeCloseTo(20, 6);
    expect(t!.color).toBe('112233');
    expect(t!.font.bold).toBe(true);
    expect(t!.y).toBeCloseTo(70, 6);
    expect(t!.xs[0]).toBeCloseTo(110 - 22, 6);
    expect(t!.xs[1]).toBeCloseTo(110 - 22 + 20, 6);
    expect(t!.widths).toEqual([20, 24]);
  });

  it('aligns text right and top using the font metrics when there is no Dx', () => {
    const items = renderMetafile(emf([logFontW(1, -10, 400, 'Arial'), select(1), rec(22, (w) => w.u32(2)), extTextOutW(50, 25, 'AB')]), FRAME, fonts);
    const [t] = texts(items);
    const total = t!.widths.reduce((a, b) => a + b, 0);
    expect(total).toBeGreaterThan(0);
    expect(t!.xs[0]! + total).toBeCloseTo(110, 6);
    expect(t!.y).toBeCloseTo(70 + t!.font.ascent * 20, 6);
    expect(t!.font.bold).toBe(false);
  });

  it('fills the ETO_OPAQUE rectangle with the background colour', () => {
    const items = renderMetafile(emf([rec(25, (w) => w.u8(255, 255, 0, 0)), extTextOutW(0, 0, '', undefined, 2)]), FRAME, fonts);
    expect(items).toEqual([]);
    const opaque = rec(84, (w) => {
      bounds(w).u32(1).f32(1, 1);
      w.u32(0, 0, 0, 76, 2, 5, 5, 15, 10, 0);
    });
    const [p] = paths(renderMetafile(emf([rec(25, (w) => w.u8(255, 255, 0, 0)), opaque]), FRAME, fonts));
    expect(p!.fill).toBe('FFFF00');
    expectPoints(p!.commands, [
      [20, 30],
      [40, 30],
      [40, 40],
      [20, 40],
    ]);
  });

  it('maps through the world transform, including ModifyWorldTransform', () => {
    const items = paths(
      renderMetafile(
        emf([
          createPen(1, 0, 1, [0, 0, 0]),
          select(1),
          xformRec(35, [2, 0, 0, 2, 10, 5]),
          rectangle(0, 0, 10, 10),
          xformRec(36, [1, 0, 0, 1, 5, 0], 2), // left multiply: translate first
          rectangle(0, 0, 10, 10),
          xformRec(36, [1, 0, 0, 1, 0, 0], 1), // identity
          rectangle(0, 0, 10, 10),
        ]),
        FRAME,
        fonts,
      ),
    );
    expect(items).toHaveLength(3);
    expectPoints(items[0]!.commands.slice(0, 3), [
      [30, 30],
      [70, 30],
      [70, 70],
    ]);
    expect(items[0]!.strokeWidth).toBeCloseTo(4, 6);
    expectPoints(items[1]!.commands.slice(0, 1), [[50, 30]]);
    expectPoints(items[2]!.commands.slice(0, 3), [
      [10, 20],
      [30, 20],
      [30, 40],
    ]);
  });

  it('maps fixed and anisotropic map modes', () => {
    // MM_HIMETRIC: 0.01 mm units, y up; 4 device pixels per mm.
    const [hi] = paths(renderMetafile(emf([rec(17, (w) => w.u32(3)), rectangle(0, 0, 1000, -500)]), FRAME, fonts));
    expectPoints(hi!.commands.slice(0, 3), [
      [10, 20],
      [90, 20],
      [90, 60],
    ]);
    // MM_ANISOTROPIC: window (0,0)+(1000,500) onto viewport (0,0)+(100,50) device pixels.
    const aniso = emf([
      rec(17, (w) => w.u32(8)),
      rec(10, (w) => w.u32(0, 0)),
      rec(9, (w) => w.u32(1000, 500)),
      rec(12, (w) => w.u32(0, 0)),
      rec(11, (w) => w.u32(100, 50)),
      rectangle(0, 0, 1000, 500),
    ]);
    expectPoints(paths(renderMetafile(aniso, FRAME, fonts))[0]!.commands.slice(0, 3), [
      [10, 20],
      [210, 20],
      [210, 120],
    ]);
    // MM_ISOTROPIC keeps units square: the smaller scale wins.
    const iso = emf([rec(17, (w) => w.u32(7)), rec(9, (w) => w.u32(100, 100)), rec(11, (w) => w.u32(100, 50)), rectangle(0, 0, 100, 100)]);
    expectPoints(paths(renderMetafile(iso, FRAME, fonts))[0]!.commands.slice(0, 3), [
      [10, 20],
      [110, 20],
      [110, 120],
    ]);
  });

  it('restores pen, brush and mapping with SaveDC / RestoreDC', () => {
    const items = paths(
      renderMetafile(
        emf([
          createBrush(1, 0, [255, 0, 0]),
          createBrush(2, 0, [0, 0, 255]),
          select(1),
          rec(33), // SaveDC
          select(2),
          rec(10, (w) => w.u32(-10, 0)), // window origin
          rectangle(0, 0, 10, 10),
          rec(34, (w) => w.u32(-1)), // RestoreDC(-1)
          rectangle(0, 0, 10, 10),
          rec(34, (w) => w.u32(-1)), // nothing left to restore: ignored
          rec(40, (w) => w.u32(1)), // deleting the selected brush keeps it in use
          rectangle(0, 0, 10, 10),
        ]),
        FRAME,
        fonts,
      ),
    );
    expect(items.map((p) => p.fill)).toEqual(['0000FF', 'FF0000', 'FF0000']);
    expectPoints(items[0]!.commands.slice(0, 1), [[30, 20]]);
    expectPoints(items[1]!.commands.slice(0, 1), [[10, 20]]);
  });

  it('draws a StretchDIBits bitmap as a raster image', () => {
    const header = new W().u32(40, 2, 2).u16(1, 24).u32(0, 16, 0, 0, 0, 0);
    // Bottom-up rows of 2 BGR pixels padded to 8 bytes: blue, white; then red, green.
    const pixels = new W().u8(255, 0, 0, 255, 255, 255, 0, 0).u8(0, 0, 255, 0, 255, 0, 0, 0);
    const blit = rec(81, (w) => {
      bounds(w).u32(10, 10, 0, 0, 2, 2, 80, 40, 120, 16, 0, 0x00cc0020, 20, 20);
      w.raw(header.bytes).raw(pixels.bytes);
    });
    const items = renderMetafile(emf([blit]), FRAME, fonts);
    expect(items).toHaveLength(1);
    const img = items[0] as ImageItem;
    expect(img.type).toBe('image');
    expect([img.x, img.y, img.w, img.h]).toEqual([30, 40, 40, 40]);
    expect(img.target).toBeUndefined();
    expect(img.external).toBe(false);
    expect(img.raster!.width).toBe(2);
    expect(Array.from(img.raster!.rgba)).toEqual([255, 0, 0, 255, 0, 255, 0, 255, 0, 0, 255, 255, 255, 255, 255, 255]);
  });

  it('decodes a 1-bpp DIB through its colour table and flips a mirrored destination', () => {
    const header = new W().u32(40, 2, -1).u16(1, 1).u32(0, 4, 0, 0, 2, 0);
    const table = new W().u8(0, 0, 0, 0, 255, 255, 255, 0);
    const pixels = new W().u8(0b10000000, 0, 0, 0); // white, black (top-down)
    const blit = rec(81, (w) => {
      bounds(w).u32(10, 10, 0, 0, 2, 1, 80, 48, 128, 4, 0, 0x00cc0020, -10, 10);
      w.raw(header.bytes).raw(table.bytes).raw(pixels.bytes);
    });
    const img = renderMetafile(emf([blit]), FRAME, fonts)[0] as ImageItem;
    expect([img.x, img.w]).toEqual([10, 20]);
    expect(Array.from(img.raster!.rgba)).toEqual([0, 0, 0, 255, 255, 255, 255, 255]);
  });

  it('draws EMF records of a dual EMF+ file and nothing for an EMF+-only file', () => {
    const emfPlus = rec(70, (w) => w.u32(16, 0x2b464d45).u16(0x4001, 1).u32(28, 16));
    expect(renderMetafile(emf([emfPlus]), FRAME, fonts)).toEqual([]);
    expect(renderMetafile(emf([emfPlus, rectangle(0, 0, 10, 10)]), FRAME, fonts)).toHaveLength(1);
  });

  it('scales the picture to whatever frame it is given', () => {
    const file = emf([rectangle(0, 0, 100, 50)]);
    expectPoints(paths(renderMetafile(file, { x: 0, y: 0, w: 100, h: 50 }, fonts))[0]!.commands.slice(0, 3), [
      [0, 0],
      [100, 0],
      [100, 50],
    ]);
    expectPoints(paths(renderMetafile(file, { x: 5, y: 7, w: 50, h: 100 }, fonts))[0]!.commands.slice(0, 3), [
      [5, 7],
      [55, 7],
      [55, 107],
    ]);
  });
});

describe('WMF', () => {
  const placeable = { box: [0, 0, 1000, 500] as [number, number, number, number], inch: 1000 };

  it('allocates objects in the lowest free slot of the object table', () => {
    const file = wmf(
      [
        wPen(0, 0, [0, 255, 0]), // slot 0
        wBrush(0, [255, 0, 0]), // slot 1
        wDelete(0), // frees slot 0
        wBrush(0, [0, 0, 255]), // takes slot 0 again
        wSelect(0),
        wrec(0x041b, (w) => w.u16(200, 300, 100, 100)), // RECTANGLE bottom, right, top, left
        wSelect(1),
        wrec(0x0418, (w) => w.u16(400, 600, 200, 400)), // ELLIPSE
      ],
      placeable,
    );
    const items = paths(renderMetafile(file, { x: 0, y: 0, w: 100, h: 50 }, fonts));
    expect(items.map((p) => p.fill)).toEqual(['0000FF', 'FF0000']);
    expect(items[0]!.stroke).toBe('000000');
    expect(items[0]!.strokeWidth).toBe(0.25);
    expectPoints(items[0]!.commands, [
      [10, 10],
      [30, 10],
      [30, 20],
      [10, 20],
    ]);
    expect(items[1]!.commands.map((c) => c.op)).toEqual(['M', 'C', 'C', 'C', 'C', 'Z']);
    expectPoints(items[1]!.commands.slice(0, 1), [[60, 30]]);
  });

  it('maps the window set by the records onto the frame', () => {
    const file = wmf(
      [
        wrec(0x020b, (w) => w.u16(-100, 50)), // window origin (y, x) = (50, -100)
        wrec(0x020c, (w) => w.u16(-200, 400)), // window extent (y, x) = (-200, 400): y up
        wPen(0, 20, [0, 0, 0]),
        wSelect(0),
        wrec(0x0214, (w) => w.u16(-100, 50)), // MOVETO (y, x)
        wrec(0x0213, (w) => w.u16(-300, 450)), // LINETO
      ],
      placeable,
    );
    const [p] = paths(renderMetafile(file, { x: 0, y: 0, w: 400, h: 200 }, fonts));
    expectPoints(p!.commands, [
      [0, 0],
      [400, 200],
    ]);
    expect(p!.strokeWidth).toBeCloseTo(20, 6);
  });

  it('reads polygons and poly-polygons with the fill mode', () => {
    const file = wmf(
      [
        wrec(0x0106, (w) => w.u16(2)), // WINDING
        wrec(0x0324, (w) => w.u16(3, 0, 0, 100, 0, 50, 100)),
        wrec(0x0538, (w) => w.u16(2, 3, 3, 0, 0, 10, 0, 10, 10, 20, 20, 30, 20, 30, 30)),
      ],
      placeable,
    );
    const [tri, multi] = paths(renderMetafile(file, { x: 0, y: 0, w: 1000, h: 500 }, fonts));
    expect(tri!.evenOdd).toBeUndefined();
    expectPoints(tri!.commands, [
      [0, 0],
      [100, 0],
      [50, 100],
    ]);
    expect(multi!.commands.map((c) => c.op)).toEqual(['M', 'L', 'L', 'Z', 'M', 'L', 'L', 'Z']);
  });

  it('writes text with TEXTOUT and EXTTEXTOUT, mapping the Symbol font to Unicode', () => {
    const file = wmf(
      [
        wFont(-100, 'Times New Roman'),
        wSelect(0),
        wrec(0x012e, (w) => w.u16(24)), // TA_BASELINE
        wrec(0x0209, (w) => w.u8(0, 0, 128, 0)),
        wrec(0x0521, (w) => w.u16(2).ascii('ab').u16(300, 100)), // TEXTOUT: len, string, y, x
        wFont(-100, 'Symbol'),
        wSelect(1),
        wrec(0x0a32, (w) => w.u16(400, 500, 3, 0).ascii('abp').u8(0).u16(60, 60, 60)), // EXTTEXTOUT with dx
      ],
      placeable,
    );
    const [latin, greek] = texts(renderMetafile(file, { x: 0, y: 0, w: 100, h: 50 }, fonts));
    expect(latin!.text).toBe('ab');
    expect(latin!.x).toBeCloseTo(10, 6);
    expect(latin!.y).toBeCloseTo(30, 6);
    expect(latin!.size).toBeCloseTo(10, 6);
    expect(latin!.color).toBe('000080');
    expect(greek!.text).toBe('αβπ');
    expect(greek!.xs.map((x) => Math.round(x * 1000) / 1000)).toEqual([50, 56, 62]);
  });

  it('draws STRETCHDIB bitmaps', () => {
    const dib = new W().u32(40, 1, 1).u16(1, 32).u32(0, 4, 0, 0, 0, 0).u8(10, 20, 30, 0);
    const file = wmf([wrec(0x0f43, (w) => w.u32(0x00cc0020).u16(0, 1, 1, 0, 0, 100, 200, 50, 100).raw(dib.bytes))], placeable);
    const img = renderMetafile(file, { x: 0, y: 0, w: 100, h: 50 }, fonts)[0] as ImageItem;
    expect([img.x, img.y, img.w, img.h]).toEqual([10, 5, 20, 10]);
    expect(Array.from(img.raster!.rgba)).toEqual([30, 20, 10, 255]);
  });

  it('finds the extent of a picture with neither a placeable header nor a window', () => {
    const file = wmf([wrec(0x041b, (w) => w.u16(300, 400, 100, 200))]);
    const [p] = paths(renderMetafile(file, { x: 0, y: 0, w: 50, h: 50 }, fonts));
    const xs = points(p!.commands).map(([x]) => x);
    const ys = points(p!.commands).map(([, y]) => y);
    expect(Math.min(...xs)).toBeCloseTo(0, 3);
    expect(Math.max(...xs)).toBeCloseTo(50, 3);
    expect(Math.min(...ys)).toBeCloseTo(0, 3);
    expect(Math.max(...ys)).toBeCloseTo(50, 3);
  });

  it('scales to the frame', () => {
    const file = wmf([wrec(0x041b, (w) => w.u16(500, 1000, 0, 0))], placeable);
    for (const frame of [
      { x: 0, y: 0, w: 72, h: 36 },
      { x: 100, y: 200, w: 300, h: 30 },
    ]) {
      expectPoints(paths(renderMetafile(file, frame, fonts))[0]!.commands.slice(0, 3), [
        [frame.x, frame.y],
        [frame.x + frame.w, frame.y],
        [frame.x + frame.w, frame.y + frame.h],
      ]);
    }
  });

  it('underlines text set in an underlined font', () => {
    const font = wrec(0x02fb, (w) => w.u16(-100, 0, 0, 0, 400).u8(0, 1, 0, 0, 0, 0, 0, 0).ascii('Arial').u8(0));
    const file = wmf([font, wSelect(0), wrec(0x0521, (w) => w.u16(1).ascii('x').u16(100, 100))], placeable);
    const items = renderMetafile(file, { x: 0, y: 0, w: 100, h: 50 }, fonts);
    const line = items.find((i): i is LineItem => i.type === 'line');
    expect(line).toBeDefined();
    expect(line!.y1).toBeGreaterThan(texts(items)[0]!.y);
  });
});
