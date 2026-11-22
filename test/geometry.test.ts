/** DrawingML shape geometry: guide formulas, arcs, preset shapes, custom geometry and line ends. */
import { describe, expect, it } from 'vitest';
import {
  PRESET_NAMES,
  arcToBeziers,
  builtinGuides,
  customGeometry,
  evaluateFormula,
  evaluateGuides,
  lineEnd,
  presetGeometry,
  readAdjustments,
  readLineEnd,
  shadeFill,
} from '../src/drawing/index.js';
import type { ShapeGeometry } from '../src/drawing/index.js';
import type { PathCommand } from '../src/layout/types.js';
import { firstChild, parseXml } from '../src/xml/index.js';

const A_NS = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"';

/** End points of every command (M, L, C), skipping Z. */
function points(cmds: PathCommand[]): Array<[number, number]> {
  const out: Array<[number, number]> = [];
  for (const c of cmds) if (c.op !== 'Z') out.push([c.x, c.y]);
  return out;
}

function expectPoint(p: [number, number] | undefined, x: number, y: number, digits = 3) {
  expect(p).toBeDefined();
  expect(p?.[0]).toBeCloseTo(x, digits);
  expect(p?.[1]).toBeCloseTo(y, digits);
}

function expectPoints(actual: Array<[number, number]>, expected: Array<[number, number]>, digits = 3) {
  expect(actual).toHaveLength(expected.length);
  expected.forEach(([x, y], i) => expectPoint(actual[i], x, y, digits));
}

function expectRect(g: ShapeGeometry, x: number, y: number, w: number, h: number) {
  expect(g.textRect.x).toBeCloseTo(x, 9);
  expect(g.textRect.y).toBeCloseTo(y, 9);
  expect(g.textRect.w).toBeCloseTo(w, 9);
  expect(g.textRect.h).toBeCloseTo(h, 9);
}

function preset(name: string, w: number, h: number, adjust?: Record<string, number>): ShapeGeometry {
  const g = presetGeometry(name, w, h, adjust);
  if (!g) throw new Error(`no preset ${name}`);
  return g;
}

function path(g: ShapeGeometry, i = 0): PathCommand[] {
  const p = g.paths[i];
  if (!p) throw new Error(`no path ${i}`);
  return p.commands;
}

/** Point on a cubic at t = 0.5. */
function bezierMid(p0: [number, number], c: Extract<PathCommand, { op: 'C' }>): [number, number] {
  return [(p0[0] + 3 * c.x1 + 3 * c.x2 + c.x) / 8, (p0[1] + 3 * c.y1 + 3 * c.y2 + c.y) / 8];
}

describe('guide formulas', () => {
  const g = builtinGuides(100, 60);
  const f = (s: string) => evaluateFormula(s, g);

  it('evaluates every operator', () => {
    expect(f('*/ w 3 4')).toBe(75);
    expect(f('+- w h 10')).toBe(150);
    expect(f('+/ w h 4')).toBe(40);
    expect(f('?: 1 2 3')).toBe(2);
    expect(f('?: 0 2 3')).toBe(3);
    expect(f('?: -1 2 3')).toBe(3);
    expect(f('abs -7')).toBe(7);
    expect(f('at2 1 1')).toBeCloseTo(2700000, 3);
    expect(f('at2 0 1')).toBeCloseTo(5400000, 3);
    expect(f('at2 -1 0')).toBeCloseTo(10800000, 3);
    expect(f('cat2 10 3 4')).toBeCloseTo(6, 9);
    expect(f('sat2 10 3 4')).toBeCloseTo(8, 9);
    expect(f('cos 10 0')).toBeCloseTo(10, 9);
    expect(f('cos 10 cd2')).toBeCloseTo(-10, 9);
    expect(f('cos 10 5400000')).toBeCloseTo(0, 9);
    expect(f('sin 10 5400000')).toBeCloseTo(10, 9);
    expect(f('sin 10 -5400000')).toBeCloseTo(-10, 9);
    expect(f('tan 10 2700000')).toBeCloseTo(10, 9);
    expect(f('max 3 -4')).toBe(3);
    expect(f('min 3 -4')).toBe(-4);
    expect(f('mod 3 4 12')).toBe(13);
    expect(f('pin 0 -5 10')).toBe(0);
    expect(f('pin 0 5 10')).toBe(5);
    expect(f('pin 0 15 10')).toBe(10);
    expect(f('sqrt 16')).toBe(4);
    expect(f('val 42')).toBe(42);
    expect(f('val w')).toBe(100);
  });

  it('degrades instead of throwing on bad input', () => {
    expect(f('*/ 1 2 0')).toBe(0);
    expect(f('+/ 1 2 0')).toBe(0);
    expect(f('sqrt -4')).toBe(0);
    expect(f('nope 1 2')).toBe(0);
    expect(f('+- unknownGuide 5 0')).toBe(5);
  });

  it('defines the built-in guides', () => {
    const expected: Record<string, number> = {
      l: 0, t: 0, r: 100, b: 60, w: 100, h: 60, hc: 50, vc: 30, ss: 60, ls: 100,
      wd2: 50, wd3: 100 / 3, wd4: 25, wd5: 20, wd6: 100 / 6, wd8: 12.5, wd10: 10, wd12: 100 / 12, wd32: 3.125,
      hd2: 30, hd3: 20, hd4: 15, hd5: 12, hd6: 10, hd8: 7.5, hd10: 6, hd12: 5, hd32: 1.875,
      ssd2: 30, ssd4: 15, ssd6: 10, ssd8: 7.5, ssd16: 3.75, ssd32: 1.875,
      cd2: 10800000, cd4: 5400000, cd8: 2700000, '3cd4': 16200000, '3cd8': 8100000, '5cd8': 13500000, '7cd8': 18900000,
    };
    for (const [name, v] of Object.entries(expected)) expect(g.get(name), name).toBeCloseTo(v, 9);
    expect(builtinGuides(40, 90).get('ss')).toBe(40);
    expect(builtinGuides(40, 90).get('ls')).toBe(90);
  });

  it('applies adjust overrides, then evaluates gdLst in order', () => {
    const av = [{ name: 'adj', fmla: 'val 25000' }];
    const gd = [
      { name: 'a', fmla: 'pin 0 adj 50000' },
      { name: '2a', fmla: '*/ a 2 1' },
      { name: 'x1', fmla: '*/ ss 2a 100000' },
    ];
    expect(evaluateGuides(100, 60, av, gd).get('x1')).toBeCloseTo(30, 9);
    expect(evaluateGuides(100, 60, av, gd, { adj: 10000 }).get('x1')).toBeCloseTo(12, 9);
    // The pin keeps an out-of-range handle inside the shape.
    expect(evaluateGuides(100, 60, av, gd, { adj: 90000 }).get('x1')).toBeCloseTo(60, 9);
  });
});

describe('arcTo', () => {
  it('draws a full ellipse from four quarter arcs through its extreme points', () => {
    const g = preset('ellipse', 100, 60);
    const cmds = path(g);
    expect(cmds.map((c) => c.op).join('')).toBe('MCCCCZ');
    expectPoints(points(cmds), [[0, 30], [50, 0], [100, 30], [50, 60], [0, 30]], 9);
    // Each quarter's midpoint lies on the ellipse to within the Bézier approximation.
    const pts = points(cmds);
    cmds.forEach((c, i) => {
      if (c.op !== 'C') return;
      const [x, y] = bezierMid(pts[i - 1] as [number, number], c);
      expect(((x - 50) / 50) ** 2 + ((y - 30) / 30) ** 2).toBeCloseTo(1, 3);
    });
  });

  it('splits sweeps into segments of at most 90 degrees, either direction', () => {
    const segs = (sw: number) => {
      const out: number[][] = [];
      const end = arcToBeziers(100, 30, 50, 30, 0, sw, (...a) => out.push(a));
      return { out, end };
    };
    expect(segs(5400000).out).toHaveLength(1);
    expect(segs(6000000).out).toHaveLength(2);
    expect(segs(10800000).out).toHaveLength(2);
    const full = segs(21600000);
    expect(full.out).toHaveLength(4);
    expect(full.end.x).toBeCloseTo(100, 9);
    expect(full.end.y).toBeCloseTo(30, 9);
    // Clockwise from 0° (y down) passes the bottom; counter-clockwise passes the top.
    expectPoint(segs(5400000).out[0]?.slice(4) as [number, number], 50, 60, 9);
    expectPoint(segs(-5400000).out[0]?.slice(4) as [number, number], 50, 0, 9);
    expect(segs(0).out).toHaveLength(0);
  });

  it('reads angles as directions from the centre, as the preset guides assume', () => {
    // `arc` computes its end points from the angles with cat2/sat2; the arc must land on them.
    const g = preset('arc', 100, 60, { adj1: 1800000, adj2: 9000000 });
    const at = (deg: number): [number, number] => {
      const a = (deg * Math.PI) / 180;
      const r = 1 / Math.sqrt((Math.cos(a) / 50) ** 2 + (Math.sin(a) / 30) ** 2);
      return [50 + r * Math.cos(a), 30 + r * Math.sin(a)];
    };
    const pts = points(path(g, 1));
    const [sx, sy] = at(30);
    const [ex, ey] = at(150);
    expectPoint(pts[0], sx, sy, 6);
    expectPoint(pts[pts.length - 1], ex, ey, 6);
  });
});

describe('preset shapes', () => {
  it('rect', () => {
    const g = preset('rect', 100, 60);
    expectPoints(points(path(g)), [[0, 0], [100, 0], [100, 60], [0, 60]]);
    expect(path(g).at(-1)?.op).toBe('Z');
    expectRect(g, 0, 0, 100, 60);
    expect(g.paths[0]?.fill).toBe('norm');
    expect(g.paths[0]?.stroke).toBe(true);
  });

  it('roundRect corner radius comes from adj (default 16667 of the short side)', () => {
    const g = preset('roundRect', 100, 60);
    const r = (60 * 16667) / 100000;
    const cmds = path(g);
    const pts = points(cmds);
    expectPoint(pts[0], 0, r);
    expectPoint(pts[1], r, 0);
    expectPoint(pts[2], 100 - r, 0);
    const corner = cmds[1];
    if (corner?.op !== 'C') throw new Error('expected a curve');
    const k = 1 - Math.SQRT1_2;
    expectPoint(bezierMid([0, r], corner), r * k, r * k);
    expect(g.textRect.x).toBeCloseTo(r * 0.29289, 4);
    // A 50% handle makes the short sides fully round.
    expectPoint(points(path(preset('roundRect', 100, 60, { adj: 50000 })))[0], 0, 30);
  });

  it('triangle apex follows adj', () => {
    expectPoints(points(path(preset('triangle', 100, 60))), [[0, 60], [50, 0], [100, 60]]);
    expectPoints(points(path(preset('triangle', 100, 60, { adj: 25000 }))), [[0, 60], [25, 0], [100, 60]]);
  });

  it('rightArrow head length from adj2 and shaft thickness from adj1', () => {
    expectPoints(points(path(preset('rightArrow', 100, 60))), [[0, 15], [70, 15], [70, 0], [100, 30], [70, 60], [70, 45], [0, 45]]);
    expectPoints(points(path(preset('rightArrow', 100, 60, { adj1: 20000, adj2: 25000 }))), [
      [0, 24], [85, 24], [85, 0], [100, 30], [85, 60], [85, 36], [0, 36],
    ]);
  });

  it('star5 points', () => {
    const g = preset('star5', 100, 100);
    const pts = points(path(g));
    expect(pts).toHaveLength(10);
    // Outer points: the spec's hf/vf factors stretch the star to touch all four sides.
    const svc = 50 * 1.10557;
    const outer = [-162, -90, -18, 54, 126].map((deg): [number, number] => {
      const a = (deg * Math.PI) / 180;
      return [50 + 50 * 1.05146 * Math.cos(a), svc + svc * Math.sin(a)];
    });
    expectPoint(pts[0], ...(outer[0] as [number, number]), 2);
    expectPoint(pts[2], 50, 0, 2);
    expectPoint(pts[4], ...(outer[2] as [number, number]), 2);
    expectPoint(pts[6], ...(outer[3] as [number, number]), 2);
    expectPoint(pts[8], ...(outer[4] as [number, number]), 2);
    expect(Math.min(...pts.map((p) => p[0]))).toBeCloseTo(0, 2);
    expect(Math.max(...pts.map((p) => p[0]))).toBeCloseTo(100, 2);
    expect(Math.max(...pts.map((p) => p[1]))).toBeCloseTo(100, 2);
    // The inner point below the centre is at the adj radius (19098 → 0.38196 of the outer).
    expectPoint(pts[7], 50, svc + svc * 0.38196, 2);
  });

  it('wedgeRectCallout puts the tail tip at adj1/adj2', () => {
    const g = preset('wedgeRectCallout', 100, 60);
    const pts = points(path(g));
    const tip: [number, number] = [50 - 20.833, 30 + 37.5];
    expect(pts.some(([x, y]) => Math.abs(x - tip[0]) < 1e-3 && Math.abs(y - tip[1]) < 1e-3)).toBe(true);
    // Pointing right: the tail leaves the right side, between 2/12 and 5/12 of the height.
    const right = points(path(preset('wedgeRectCallout', 100, 60, { adj1: 80000, adj2: 0 })));
    const i = right.findIndex(([x, y]) => Math.abs(x - 130) < 1e-6 && Math.abs(y - 30) < 1e-6);
    expect(i).toBeGreaterThan(0);
    expectPoint(right[i - 1], 100, 10);
    expectPoint(right[i + 1], 100, 25);
  });

  it('flowChartDecision is a diamond in a 2×2 path space', () => {
    const g = preset('flowChartDecision', 100, 60);
    expectPoints(points(path(g)), [[0, 30], [50, 0], [100, 30], [50, 60]]);
    expectRect(g, 25, 15, 50, 30);
  });

  it("can's top is a lighter ellipse", () => {
    const g = preset('can', 100, 60);
    expect(g.paths.map((p) => [p.fill, p.stroke])).toEqual([
      ['norm', false],
      ['lighten', false],
      ['none', true],
    ]);
    const y1 = (60 * 25000) / 200000;
    const top = path(g, 1);
    expect(top.map((c) => c.op).join('')).toBe('MCCCCZ');
    expectPoints(points(top), [[0, y1], [50, 0], [100, y1], [50, 2 * y1], [0, y1]]);
    expect(g.textRect.y).toBeCloseTo(2 * y1, 9);
  });

  it('cube shades its faces', () => {
    expect(preset('cube', 100, 60).paths.map((p) => p.fill)).toEqual(['norm', 'darkenLess', 'lightenLess', 'none']);
  });

  it('accepts adj for a shape whose single handle is adj1', () => {
    expectPoint(points(path(preset('bentConnector3', 100, 60, { adj: 25000 })))[1], 25, 0);
  });

  it('returns undefined for an unknown preset', () => {
    expect(presetGeometry('noSuchShape', 10, 10)).toBeUndefined();
  });

  const REQUIRED = [
    'line', 'straightConnector1', 'bentConnector2', 'bentConnector3', 'curvedConnector3', 'rect', 'roundRect', 'round1Rect',
    'round2SameRect', 'round2DiagRect', 'snip1Rect', 'snip2SameRect', 'snip2DiagRect', 'snipRoundRect', 'ellipse', 'triangle',
    'rtTriangle', 'parallelogram', 'trapezoid', 'diamond', 'pentagon', 'hexagon', 'heptagon', 'octagon', 'decagon', 'dodecagon',
    'star4', 'star5', 'star6', 'star7', 'star8', 'star10', 'star12', 'star16', 'star24', 'star32', 'plus', 'mathPlus', 'mathMinus',
    'mathMultiply', 'mathDivide', 'mathEqual', 'mathNotEqual', 'donut', 'noSmoking', 'blockArc', 'can', 'cube', 'bevel', 'frame',
    'halfFrame', 'corner', 'diagStripe', 'plaque', 'foldedCorner', 'smileyFace', 'heart', 'lightningBolt', 'sun', 'moon', 'cloud',
    'arc', 'chord', 'pie', 'pieWedge', 'teardrop', 'bracketPair', 'bracePair', 'leftBracket', 'rightBracket', 'leftBrace',
    'rightBrace', 'rightArrow', 'leftArrow', 'upArrow', 'downArrow', 'leftRightArrow', 'upDownArrow', 'quadArrow',
    'leftRightUpArrow', 'notchedRightArrow', 'stripedRightArrow', 'chevron', 'homePlate', 'bentArrow', 'uturnArrow',
    'curvedRightArrow', 'curvedLeftArrow', 'curvedUpArrow', 'curvedDownArrow', 'circularArrow', 'rightArrowCallout',
    'leftArrowCallout', 'wedgeRectCallout', 'wedgeRoundRectCallout', 'wedgeEllipseCallout', 'cloudCallout', 'borderCallout1',
    'borderCallout2', 'callout1', 'accentCallout1', 'ribbon', 'ribbon2', 'wave', 'doubleWave', 'horizontalScroll', 'verticalScroll',
    'flowChartProcess', 'flowChartAlternateProcess', 'flowChartDecision', 'flowChartInputOutput', 'flowChartPredefinedProcess',
    'flowChartInternalStorage', 'flowChartDocument', 'flowChartMultidocument', 'flowChartTerminator', 'flowChartPreparation',
    'flowChartManualInput', 'flowChartManualOperation', 'flowChartConnector', 'flowChartOffpageConnector', 'flowChartPunchedCard',
    'flowChartPunchedTape', 'flowChartSummingJunction', 'flowChartOr', 'flowChartCollate', 'flowChartSort', 'flowChartExtract',
    'flowChartMerge', 'flowChartOnlineStorage', 'flowChartDelay', 'flowChartMagneticTape', 'flowChartMagneticDisk',
    'flowChartMagneticDrum', 'flowChartDisplay',
  ];

  it('has every required preset', () => {
    expect(REQUIRED.filter((n) => !PRESET_NAMES.includes(n))).toEqual([]);
  });

  it.each([
    [100, 60],
    [60, 100],
    [0, 0],
    [100, 0],
  ])('every preset evaluates to finite paths at %d×%d', (w, h) => {
    for (const name of PRESET_NAMES) {
      const g = preset(name, w, h);
      expect(g.paths.length, name).toBeGreaterThan(0);
      for (const p of g.paths) {
        expect(p.commands[0]?.op, name).toBe('M');
        for (const c of p.commands) {
          for (const v of Object.values(c)) if (typeof v === 'number') expect(Number.isFinite(v), `${name} ${JSON.stringify(c)}`).toBe(true);
        }
      }
      for (const v of Object.values(g.textRect)) expect(Number.isFinite(v), name).toBe(true);
    }
  });
});

describe('custom geometry', () => {
  const xml = `<a:custGeom ${A_NS}>
    <a:avLst/>
    <a:gdLst><a:gd name="half" fmla="*/ w 1 2"/></a:gdLst>
    <a:ahLst/><a:cxnLst/>
    <a:rect l="l" t="t" r="half" b="b"/>
    <a:pathLst>
      <a:path w="200" h="100">
        <a:moveTo><a:pt x="0" y="100"/></a:moveTo>
        <a:lnTo><a:pt x="100" y="0"/></a:lnTo>
        <a:quadBezTo><a:pt x="150" y="0"/><a:pt x="200" y="50"/></a:quadBezTo>
        <a:cubicBezTo><a:pt x="200" y="75"/><a:pt x="150" y="100"/><a:pt x="100" y="100"/></a:cubicBezTo>
        <a:arcTo wR="50" hR="50" stAng="5400000" swAng="5400000"/>
        <a:close/>
      </a:path>
      <a:path fill="none" stroke="1" extrusionOk="0">
        <a:moveTo><a:pt x="0" y="0"/></a:moveTo>
        <a:lnTo><a:pt x="half" y="b"/></a:lnTo>
        <a:lnTo><a:pt x="127000" y="0"/></a:lnTo>
      </a:path>
      <a:path stroke="false" fill="darken"><a:moveTo><a:pt x="0" y="0"/></a:moveTo></a:path>
    </a:pathLst>
  </a:custGeom>`;
  const el = parseXml(xml).root;

  it('scales a path from its own w/h space into the shape', () => {
    const g = customGeometry(el, 50, 25);
    const cmds = path(g);
    expect(cmds.map((c) => c.op).join('')).toBe('MLCCCZ');
    const quad = cmds[2];
    if (quad?.op !== 'C') throw new Error('expected a curve');
    // The quadratic (control 37.5,0) becomes a cubic with controls 2/3 of the way to it.
    expect(quad.x1).toBeCloseTo(100 / 3, 6);
    expect(quad.y1).toBeCloseTo(0, 6);
    expect(quad.x2).toBeCloseTo(125 / 3, 6);
    expect(quad.y2).toBeCloseTo(12.5 / 3, 6);
    expectPoints(points(cmds), [[0, 25], [25, 0], [50, 12.5], [25, 25], [12.5, 12.5]], 6);
  });

  it('reads guides, fill and stroke flags, and EMU coordinates when the path has no w/h', () => {
    const g = customGeometry(el, 50, 25);
    expect(g.paths.map((p) => [p.fill, p.stroke])).toEqual([
      ['norm', true],
      ['none', true],
      ['darken', false],
    ]);
    expectPoints(points(path(g, 1)), [[0, 0], [25, 25], [10, 0]], 6);
    expectRect(g, 0, 0, 25, 25);
  });

  it('reads adjust values from a prstGeom', () => {
    const prst = parseXml(`<a:prstGeom ${A_NS} prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 50000"/></a:avLst></a:prstGeom>`).root;
    const adj = readAdjustments(prst);
    expect(adj).toEqual({ adj: 50000 });
    expectPoint(points(path(preset('roundRect', 100, 60, adj)))[0], 0, 30);
    expect(readAdjustments(parseXml(`<a:prstGeom ${A_NS} prst="rect"/>`).root)).toEqual({});
  });
});

describe('shadeFill', () => {
  it('darkens toward black and lightens toward white', () => {
    expect(shadeFill('808080', 'norm')).toBe('808080');
    expect(shadeFill('808080', 'darken')).toBe('4D4D4D');
    expect(shadeFill('808080', 'darkenLess')).toBe('666666');
    expect(shadeFill('808080', 'lighten')).toBe('B3B3B3');
    expect(shadeFill('808080', 'lightenLess')).toBe('999999');
    expect(shadeFill('ff0000', 'darken')).toBe('990000');
    expect(shadeFill('#000000', 'lighten')).toBe('666666');
  });
});

describe('line ends', () => {
  const tip = { x: 100, y: 50 };
  const from = { x: 0, y: 50 };

  it('triangle sized from the line width, pulling the line back to its base', () => {
    const e = lineEnd('triangle', 'med', 'med', tip, from, 3);
    expect(e.filled).toBe(true);
    expect(e.commands.map((c) => c.op).join('')).toBe('MLLZ');
    expectPoints(points(e.commands), [[100, 50], [91, 54.5], [91, 45.5]], 9);
    expect(e.inset).toBeCloseTo(9 - 1.5, 9);
    // sm/lg pick 2× and 5×.
    const big = lineEnd('triangle', 'lg', 'sm', tip, from, 3);
    expectPoints(points(big.commands), [[100, 50], [94, 57.5], [94, 42.5]], 9);
  });

  it('keeps heads visible on hairlines', () => {
    const e = lineEnd('triangle', 'med', 'med', tip, from, 0.25);
    expect(points(e.commands)[1]?.[0]).toBeCloseTo(94, 9);
  });

  it('follows the line direction', () => {
    const e = lineEnd('triangle', 'med', 'med', { x: 0, y: 100 }, { x: 0, y: 0 }, 2);
    expectPoints(points(e.commands), [[0, 100], [-3, 94], [3, 94]], 9);
  });

  it('stealth has a notch the line runs into', () => {
    const e = lineEnd('stealth', 'med', 'med', tip, from, 3);
    expectPoints(points(e.commands), [[100, 50], [91, 54.5], [94, 50], [91, 45.5]], 9);
    expect(e.inset).toBeCloseTo(6, 9);
  });

  it('diamond and oval sit centred on the end point', () => {
    const d = lineEnd('diamond', 'med', 'med', tip, from, 2);
    expectPoints(points(d.commands), [[103, 50], [100, 53], [97, 50], [100, 47]], 9);
    expect(d.inset).toBe(0);
    const o = lineEnd('oval', 'lg', 'med', tip, from, 2);
    expect(o.commands.map((c) => c.op).join('')).toBe('MCCCCZ');
    expectPoints(points(o.commands), [[103, 50], [100, 55], [97, 50], [100, 45], [103, 50]], 9);
    expect(o.filled).toBe(true);
  });

  it('arrow is an open, stroked chevron', () => {
    const e = lineEnd('arrow', 'med', 'med', tip, from, 2);
    expect(e.filled).toBe(false);
    expect(e.commands.map((c) => c.op).join('')).toBe('MLL');
    expectPoints(points(e.commands), [[94, 53], [100, 50], [94, 47]], 9);
  });

  it('draws nothing for none or unknown types', () => {
    expect(lineEnd('none', 'med', 'med', tip, from, 1)).toEqual({ commands: [], filled: false, inset: 0 });
    expect(lineEnd('bogus', undefined, undefined, tip, from, 1).commands).toEqual([]);
  });

  it('reads headEnd/tailEnd elements', () => {
    const ln = parseXml(`<a:ln ${A_NS}><a:headEnd type="none"/><a:tailEnd type="stealth" w="lg" len="sm"/></a:ln>`).root;
    expect(readLineEnd(firstChild(ln, 'a:headEnd'))).toBeUndefined();
    expect(readLineEnd(firstChild(ln, 'a:tailEnd'))).toEqual({ type: 'stealth', w: 'lg', len: 'sm' });
    expect(readLineEnd(parseXml(`<a:tailEnd ${A_NS} type="oval"/>`).root)).toEqual({ type: 'oval', w: 'med', len: 'med' });
    expect(readLineEnd(undefined)).toBeUndefined();
  });
});
