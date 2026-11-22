/**
 * Preset shape definitions (ECMA-376 Part 1, §20.1.10.56 ST_ShapeType, with the geometry from the
 * spec's presetShapeDefinitions.xml), written out by hand as data.
 *
 * Each shape is kept in a compact text form close to the XML so it can be checked against the spec
 * line by line:
 *   - `av`: adjust handles and their defaults, `name value` pairs;
 *   - `gd`: guides, one `name formula` per line, in spec order;
 *   - `path`: one string per `a:path`. Optional attributes come before a `|` (`none`, `lighten`,
 *     `lightenLess`, `darken`, `darkenLess` for the fill, `nostroke`, `noext` for
 *     extrusionOk="false", `w=N h=N` for the path coordinate space). Commands are `M x y` (moveTo),
 *     `L x y` (lnTo), `A wR hR stAng swAng` (arcTo), `Q x1 y1 x y` (quadBezTo),
 *     `C x1 y1 x2 y2 x y` (cubicBezTo) and `Z` (close);
 *   - `rect`: the text rectangle `l t r b`.
 * The text form is parsed into a ShapeDefinition on first use and cached.
 */
import type { GuideDef } from './formula.js';

export type PathFill = 'norm' | 'none' | 'lighten' | 'lightenLess' | 'darken' | 'darkenLess';

/** A point whose coordinates are guide names or numbers, as in `a:pt`. */
export interface PathPoint {
  x: string;
  y: string;
}

export type PathCommandDef =
  | { op: 'moveTo'; pt: PathPoint }
  | { op: 'lnTo'; pt: PathPoint }
  | { op: 'arcTo'; wR: string; hR: string; stAng: string; swAng: string }
  | { op: 'quadBezTo'; pts: [PathPoint, PathPoint] }
  | { op: 'cubicBezTo'; pts: [PathPoint, PathPoint, PathPoint] }
  | { op: 'close' };

/** One `a:path`. `w`/`h` of 0 mean the path uses the shape's own coordinates. */
export interface PathDef {
  w: number;
  h: number;
  fill: PathFill;
  stroke: boolean;
  extrusionOk: boolean;
  commands: PathCommandDef[];
}

/** `a:rect`: the text box, as guide names or numbers. */
export interface TextRectDef {
  l: string;
  t: string;
  r: string;
  b: string;
}

/** The parsed geometry of a preset or custom shape, before evaluation. */
export interface ShapeDefinition {
  avLst: GuideDef[];
  gdLst: GuideDef[];
  pathLst: PathDef[];
  /** Absent means the whole shape box. */
  textRect?: TextRectDef;
}

interface Src {
  av?: string;
  gd?: string;
  rect?: string;
  path: string | string[];
}

const FILLS = new Set<string>(['none', 'lighten', 'lightenLess', 'darken', 'darkenLess']);

function parseAv(src: string | undefined): GuideDef[] {
  if (!src) return [];
  const t = src.trim().split(/\s+/);
  const out: GuideDef[] = [];
  for (let i = 0; i + 1 < t.length; i += 2) out.push({ name: t[i] as string, fmla: `val ${t[i + 1] as string}` });
  return out;
}

function parseGd(src: string | undefined): GuideDef[] {
  if (!src) return [];
  const out: GuideDef[] = [];
  for (const raw of src.split(/[\n;]/)) {
    const line = raw.trim();
    if (!line) continue;
    const i = line.indexOf(' ');
    if (i < 0) throw new Error(`Bad guide "${line}"`);
    out.push({ name: line.slice(0, i), fmla: line.slice(i + 1).trim() });
  }
  return out;
}

function parsePath(src: string): PathDef {
  const bar = src.indexOf('|');
  const opts = bar < 0 ? [] : src.slice(0, bar).trim().split(/\s+/).filter(Boolean);
  const body = (bar < 0 ? src : src.slice(bar + 1)).trim().split(/\s+/).filter(Boolean);
  const def: PathDef = { w: 0, h: 0, fill: 'norm', stroke: true, extrusionOk: true, commands: [] };
  for (const o of opts) {
    if (FILLS.has(o)) def.fill = o as PathFill;
    else if (o === 'nostroke') def.stroke = false;
    else if (o === 'noext') def.extrusionOk = false;
    else if (o.startsWith('w=')) def.w = Number(o.slice(2));
    else if (o.startsWith('h=')) def.h = Number(o.slice(2));
    else throw new Error(`Bad path option "${o}"`);
  }
  let i = 0;
  const next = (): string => {
    const t = body[i++];
    if (t === undefined) throw new Error(`Path ends early: ${src}`);
    return t;
  };
  const pt = (): PathPoint => ({ x: next(), y: next() });
  while (i < body.length) {
    const op = next();
    switch (op) {
      case 'M':
        def.commands.push({ op: 'moveTo', pt: pt() });
        break;
      case 'L':
        def.commands.push({ op: 'lnTo', pt: pt() });
        break;
      case 'A':
        def.commands.push({ op: 'arcTo', wR: next(), hR: next(), stAng: next(), swAng: next() });
        break;
      case 'Q':
        def.commands.push({ op: 'quadBezTo', pts: [pt(), pt()] });
        break;
      case 'C':
        def.commands.push({ op: 'cubicBezTo', pts: [pt(), pt(), pt()] });
        break;
      case 'Z':
        def.commands.push({ op: 'close' });
        break;
      default:
        throw new Error(`Bad path command "${op}" in ${src}`);
    }
  }
  return def;
}

function parseSrc(src: Src): ShapeDefinition {
  const def: ShapeDefinition = {
    avLst: parseAv(src.av),
    gdLst: parseGd(src.gd),
    pathLst: (Array.isArray(src.path) ? src.path : [src.path]).map(parsePath),
  };
  if (src.rect) {
    const [l = 'l', t = 't', r = 'r', b = 'b'] = src.rect.trim().split(/\s+/);
    def.textRect = { l, t, r, b };
  }
  return def;
}

// ---------------------------------------------------------------------------------------------
// Shared fragments

const RECT_PATH = 'M l t L r t L r b L l b Z';
const ELLIPSE_PATH = 'M l vc A wd2 hd2 cd2 cd4 A wd2 hd2 3cd4 cd4 A wd2 hd2 0 cd4 A wd2 hd2 cd4 cd4 Z';
/** The text box inscribed in an ellipse, used by every round shape. */
const ELLIPSE_RECT_GD = `idx cos wd2 2700000
idy sin hd2 2700000
il +- hc 0 idx
ir +- hc idx 0
it +- vc 0 idy
ib +- vc idy 0`;

/** The wedge callouts: where the tail leaves the body depends on which side the tip is on. */
const WEDGE_GD = `dxPos */ w adj1 100000
dyPos */ h adj2 100000
xPos +- hc dxPos 0
yPos +- vc dyPos 0
dx */ dxPos h 1
dy */ dyPos w 1
adx abs dx
ady abs dy
dq +- adx 0 ady
xg1 ?: dxPos 7 2
xg2 ?: dxPos 10 5
x1 */ w xg1 12
x2 */ w xg2 12
yg1 ?: dyPos 7 2
yg2 ?: dyPos 10 5
y1 */ h yg1 12
y2 */ h yg2 12
t1 ?: dxPos l xPos
xl ?: dq t1 l
t2 ?: dyPos x1 xPos
xt ?: dq x1 t2
t3 ?: dxPos xPos r
xr ?: dq t3 r
t4 ?: dyPos xPos x1
xb ?: dq x1 t4
t5 ?: dxPos y1 yPos
yl ?: dq t5 y1
t6 ?: dyPos t yPos
yt ?: dq t t6
t7 ?: dxPos yPos y1
yr ?: dq t7 y1
t8 ?: dyPos yPos b
yb ?: dq b t8`;

const CLOUD_RECT_GD = `il */ w 2977 21600
it */ h 3262 21600
ir */ w 17087 21600
ib */ h 17337 21600`;
const CLOUD_OUTLINE =
  'w=43200 h=43200 | M 3900 14370 A 6753 9190 -11429249 7426832 A 5333 7267 -8646143 5396714 ' +
  'A 4365 5945 -8748475 5983381 A 4857 6595 -7859164 7034504 A 5333 7273 -4722533 6541615 ' +
  'A 6775 9220 -2776035 7816140 A 5785 7867 37501 6842000 A 6752 9215 1347096 6910353 ' +
  'A 7720 10543 3974558 4542661 A 4360 5918 -16496525 8804134 A 4345 5945 -14809710 9151131 Z';
const CLOUD_DETAIL =
  'none noext w=43200 h=43200 | M 4693 26177 A 4345 5945 5204520 1585770 M 6928 34899 A 4360 5918 4416628 686848 ' +
  'M 16478 39090 A 6752 9215 8257449 844866 M 28827 34751 A 6752 9215 387196 959901 ' +
  'M 34129 22954 A 5785 7867 -4217541 4255042 M 41798 15354 A 5333 7273 1819082 1665090 ' +
  'M 38324 5426 A 4857 6595 -824660 891534 M 29078 3952 A 4857 6595 -8950887 1091722 ' +
  'M 22141 4720 A 4365 5945 -9809656 1061181 M 14000 5192 A 6753 9190 -4002417 739161 ' +
  'M 4127 15789 A 6753 9190 9459261 711490';

/** Callouts keep their leader line ends as fractions of the box in adj1…adj8 (y, x pairs). */
function calloutGd(n: number): string {
  const lines: string[] = [];
  for (let i = 1; i <= n; i++) lines.push(`y${i} */ h adj${2 * i - 1} 100000`, `x${i} */ w adj${2 * i} 100000`);
  return lines.join('\n');
}
const CALLOUT_AV = [
  'adj1 18750 adj2 -8333 adj3 112500 adj4 -38333',
  'adj1 18750 adj2 -8333 adj3 18750 adj4 -16667 adj5 112500 adj6 -46667',
  'adj1 18750 adj2 -8333 adj3 18750 adj4 -16667 adj5 100000 adj6 -16667 adj7 112963 adj8 -8333',
] as const;
function leader(n: number): string {
  let s = 'none noext | M x1 y1';
  for (let i = 2; i <= n + 1; i++) s += ` L x${i} y${i}`;
  return s;
}
/** callout1-3, borderCallout1-3, accentCallout1-3, accentBorderCallout1-3. */
function callout(n: 1 | 2 | 3, border: boolean, accent: boolean): Src {
  const path = [border ? RECT_PATH : `nostroke | ${RECT_PATH}`];
  if (accent) path.push('none noext | M x1 t L x1 b');
  path.push(leader(n));
  return { av: CALLOUT_AV[n - 1], gd: calloutGd(n + 1), path };
}

/**
 * star16, star24 and star32 are regular stars whose points the spec lists as cos/sin multiples;
 * generating them from the same angles gives the same coordinates without hand-copying 64 points.
 * Outer points start at the left (180°) and run clockwise, inner points sit halfway between.
 */
function regularStar(n: number, adj: number): Src {
  const gd: string[] = ['a pin 0 adj 50000', 'iwd2 */ wd2 a 50000', 'ihd2 */ hd2 a 50000'];
  let path = '';
  const step = 360 / n;
  for (let k = 0; k < n; k++) {
    const outer = Math.round((180 + k * step) * 60000);
    const inner = Math.round((180 + (k + 0.5) * step) * 60000);
    gd.push(`odx${k} cos wd2 ${outer}`, `ody${k} sin hd2 ${outer}`, `ox${k} +- hc odx${k} 0`, `oy${k} +- vc ody${k} 0`);
    gd.push(`idx${k} cos iwd2 ${inner}`, `idy${k} sin ihd2 ${inner}`, `ix${k} +- hc idx${k} 0`, `iy${k} +- vc idy${k} 0`);
    path += `${k === 0 ? 'M' : 'L'} ox${k} oy${k} L ix${k} iy${k} `;
  }
  // Text box corners on the inner radius at the last inner angle below 45°, as the spec does.
  let off = 0;
  for (let k = 0; (k + 0.5) * step <= 45; k++) off = (k + 0.5) * step;
  const c = Math.round(off * 60000);
  gd.push(`tdx cos iwd2 ${c}`, `tdy cos ihd2 ${c}`, `tl +- hc 0 tdx`, `tt +- vc 0 tdy`, `tr +- hc tdx 0`, `tb +- vc tdy 0`);
  return { av: `adj ${adj}`, gd: gd.join('\n'), rect: 'tl tt tr tb', path: `${path}Z` };
}

// ---------------------------------------------------------------------------------------------
// The table

const SOURCES: Record<string, Src> = {
  // --- Lines and connectors -------------------------------------------------------------------
  line: { path: 'none | M l t L r b' },
  lineInv: { path: 'none | M l b L r t' },
  straightConnector1: { path: 'none | M l t L r b' },
  bentConnector2: { path: 'none | M l t L r t L r b' },
  bentConnector3: {
    av: 'adj1 50000',
    gd: 'x1 */ w adj1 100000',
    path: 'none | M l t L x1 t L x1 b L r b',
  },
  bentConnector4: {
    av: 'adj1 50000 adj2 50000',
    gd: `x1 */ w adj1 100000
x2 +/ x1 r 2
y2 */ h adj2 100000
y1 +/ t y2 2`,
    path: 'none | M l t L x1 t L x1 y2 L r y2 L r b',
  },
  bentConnector5: {
    av: 'adj1 50000 adj2 50000 adj3 50000',
    gd: `x1 */ w adj1 100000
x3 */ w adj3 100000
x2 +/ x1 x3 2
y2 */ h adj2 100000
y1 +/ t y2 2
y3 +/ b y2 2`,
    path: 'none | M l t L x1 t L x1 y2 L x3 y2 L x3 b L r b',
  },
  curvedConnector2: { path: 'none | M l t C wd2 t r hd2 r b' },
  curvedConnector3: {
    av: 'adj1 50000',
    gd: `x2 */ w adj1 100000
x1 +/ l x2 2
x3 +/ r x2 2
y3 */ h 3 4`,
    path: 'none | M l t C x1 t x2 hd4 x2 vc C x2 y3 x3 b r b',
  },
  curvedConnector4: {
    av: 'adj1 50000 adj2 50000',
    gd: `x2 */ w adj1 100000
x1 +/ l x2 2
x3 +/ r x2 2
x4 +/ x2 x3 2
x5 +/ x3 r 2
y4 */ h adj2 100000
y1 +/ t y4 2
y2 +/ t y1 2
y3 +/ y1 y4 2
y5 +/ b y4 2`,
    path: 'none | M l t C x1 t x2 y2 x2 y1 C x2 y3 x4 y4 x3 y4 C x5 y4 r y5 r b',
  },
  curvedConnector5: {
    av: 'adj1 50000 adj2 50000 adj3 50000',
    gd: `x3 */ w adj1 100000
x6 */ w adj3 100000
x1 +/ x3 x6 2
x2 +/ l x3 2
x4 +/ x3 x1 2
x5 +/ x6 x1 2
x7 +/ x6 r 2
y4 */ h adj2 100000
y1 +/ t y4 2
y2 +/ t y1 2
y3 +/ y1 y4 2
y5 +/ b y4 2
y6 +/ y5 y4 2
y7 +/ y5 b 2`,
    path: 'none | M l t C x2 t x3 y2 x3 y1 C x3 y3 x4 y4 x1 y4 C x5 y4 x6 y6 x6 y5 C x6 y7 x7 b r b',
  },

  // --- Rectangles -----------------------------------------------------------------------------
  rect: { path: RECT_PATH },
  roundRect: {
    av: 'adj 16667',
    gd: `a pin 0 adj 50000
x1 */ ss a 100000
x2 +- r 0 x1
y2 +- b 0 x1
il */ x1 29289 100000
ir +- r 0 il
ib +- b 0 il`,
    rect: 'il il ir ib',
    path: 'M l x1 A x1 x1 cd2 cd4 L x2 t A x1 x1 3cd4 cd4 L r y2 A x1 x1 0 cd4 L x1 b A x1 x1 cd4 cd4 Z',
  },
  round1Rect: {
    av: 'adj 16667',
    gd: `a pin 0 adj 50000
dx1 */ ss a 100000
x1 +- r 0 dx1
idx */ dx1 29289 100000
ir +- r 0 idx`,
    rect: 'l t ir b',
    path: 'M l t L x1 t A dx1 dx1 3cd4 cd4 L r b L l b Z',
  },
  round2SameRect: {
    av: 'adj1 16667 adj2 0',
    gd: `a1 pin 0 adj1 50000
a2 pin 0 adj2 50000
tx1 */ ss a1 100000
tx2 +- r 0 tx1
bx1 */ ss a2 100000
bx2 +- r 0 bx1
by1 +- b 0 bx1
d +- tx1 0 bx1
tdx */ tx1 29289 100000
bdx */ bx1 29289 100000
il ?: d tdx bdx
ir +- r 0 il
ib +- b 0 bdx`,
    rect: 'il tdx ir ib',
    path: 'M tx1 t L tx2 t A tx1 tx1 3cd4 cd4 L r by1 A bx1 bx1 0 cd4 L bx1 b A bx1 bx1 cd4 cd4 L l tx1 A tx1 tx1 cd2 cd4 Z',
  },
  round2DiagRect: {
    av: 'adj1 16667 adj2 0',
    gd: `a1 pin 0 adj1 50000
a2 pin 0 adj2 50000
x1 */ ss a1 100000
y1 +- b 0 x1
a */ ss a2 100000
x2 +- r 0 a
y2 +- b 0 a
dx1 */ x1 29289 100000
dx2 */ a 29289 100000
d +- dx1 0 dx2
dx ?: d dx1 dx2
ir +- r 0 dx
ib +- b 0 dx`,
    rect: 'dx dx ir ib',
    path: 'M x1 t L x2 t A a a 3cd4 cd4 L r y1 A x1 x1 0 cd4 L a b A a a cd4 cd4 L l x1 A x1 x1 cd2 cd4 Z',
  },
  snip1Rect: {
    av: 'adj 16667',
    gd: `a pin 0 adj 50000
dx1 */ ss a 100000
x1 +- r 0 dx1
it */ dx1 1 2
ir +/ x1 r 2`,
    rect: 'l it ir b',
    path: 'M l t L x1 t L r dx1 L r b L l b Z',
  },
  snip2SameRect: {
    av: 'adj1 16667 adj2 0',
    gd: `a1 pin 0 adj1 50000
a2 pin 0 adj2 50000
tx1 */ ss a1 100000
tx2 +- r 0 tx1
bx1 */ ss a2 100000
bx2 +- r 0 bx1
by1 +- b 0 bx1
d +- tx1 0 bx1
dx ?: d tx1 bx1
il */ dx 1 2
ir +- r 0 il
it */ tx1 1 2
ib +/ by1 b 2`,
    rect: 'il it ir ib',
    path: 'M tx1 t L tx2 t L r tx1 L r by1 L bx2 b L bx1 b L l by1 L l tx1 Z',
  },
  snip2DiagRect: {
    av: 'adj1 0 adj2 16667',
    gd: `a1 pin 0 adj1 50000
a2 pin 0 adj2 50000
lx1 */ ss a1 100000
lx2 +- r 0 lx1
ly1 +- b 0 lx1
rx1 */ ss a2 100000
rx2 +- r 0 rx1
ry1 +- b 0 rx1
d +- lx1 0 rx1
dx ?: d lx1 rx1
il */ dx 1 2
ir +- r 0 il
ib +- b 0 il`,
    rect: 'il il ir ib',
    path: 'M lx1 t L rx2 t L r rx1 L r ly1 L lx2 b L rx1 b L l ry1 L l lx1 Z',
  },
  snipRoundRect: {
    av: 'adj1 16667 adj2 16667',
    gd: `a1 pin 0 adj1 50000
a2 pin 0 adj2 50000
x1 */ ss a1 100000
dx2 */ ss a2 100000
x2 +- r 0 dx2
il */ x1 29289 100000
ir +/ x2 r 2`,
    rect: 'il il ir b',
    path: 'M x1 t L x2 t L r dx2 L r b L l b L l x1 A x1 x1 cd2 cd4 Z',
  },

  // --- Basic shapes ---------------------------------------------------------------------------
  ellipse: { gd: ELLIPSE_RECT_GD, rect: 'il it ir ib', path: ELLIPSE_PATH },
  triangle: {
    av: 'adj 50000',
    gd: `a pin 0 adj 100000
x1 */ w a 200000
x2 */ w a 100000
x3 +- x1 wd2 0`,
    rect: 'x1 vc x3 b',
    path: 'M l b L x2 t L r b Z',
  },
  rtTriangle: {
    gd: `it */ h 7 12
ir */ w 7 12
ib */ h 11 12`,
    rect: 'l it ir ib',
    path: 'M l b L l t L r b Z',
  },
  parallelogram: {
    av: 'adj 25000',
    gd: `maxAdj */ 100000 w ss
a pin 0 adj maxAdj
x1 */ ss a 200000
x2 */ ss a 100000
x6 +- r 0 x1
x5 +- r 0 x2
x3 */ x5 1 2
x4 +- r 0 x3
il */ wd2 a maxAdj
q1 */ 5 a maxAdj
q2 +/ 1 q1 12
il */ q2 w 1
it */ q2 h 1
ir +- r 0 il
ib +- b 0 it
q3 */ h hc x2
y1 pin 0 q3 h
y2 +- b 0 y1`,
    rect: 'il it ir ib',
    path: 'M l b L x2 t L r t L x5 b Z',
  },
  trapezoid: {
    av: 'adj 25000',
    gd: `maxAdj */ 50000 w ss
a pin 0 adj maxAdj
x1 */ ss a 200000
x2 */ ss a 100000
x3 +- r 0 x2
x4 +- r 0 x1
il */ wd3 a maxAdj
it */ hd3 a maxAdj
ir +- r 0 il`,
    rect: 'il it ir b',
    path: 'M l b L x2 t L x3 t L r b Z',
  },
  diamond: {
    gd: `ir */ w 3 4
ib */ h 3 4`,
    rect: 'wd4 hd4 ir ib',
    path: 'M l vc L hc t L r vc L hc b Z',
  },
  pentagon: {
    av: 'hf 105146 vf 110557',
    gd: `swd2 */ wd2 hf 100000
shd2 */ hd2 vf 100000
svc */ vc vf 100000
dx1 cos swd2 1080000
dx2 cos swd2 18360000
dy1 sin shd2 1080000
dy2 sin shd2 18360000
x1 +- hc 0 dx1
x2 +- hc 0 dx2
x3 +- hc dx2 0
x4 +- hc dx1 0
y1 +- svc 0 dy1
y2 +- svc 0 dy2
it */ y1 dx2 dx1`,
    rect: 'x2 it x3 y2',
    path: 'M x1 y1 L hc t L x4 y1 L x3 y2 L x2 y2 Z',
  },
  hexagon: {
    av: 'adj 25000 vf 115470',
    gd: `maxAdj */ 50000 w ss
a pin 0 adj maxAdj
shd2 */ hd2 vf 100000
x1 */ ss a 100000
x2 +- r 0 x1
dy1 sin shd2 3600000
y1 +- vc 0 dy1
y2 +- vc dy1 0
q1 */ maxAdj -1 2
q2 +- a q1 0
q3 ?: q2 4 2
q4 ?: q2 3 2
q5 ?: q2 q1 0
q6 +/ a q5 q1
q7 */ q6 q4 -1
q8 +- q3 q7 0
il */ w q8 24
it */ h q8 24
ir +- r 0 il
ib +- b 0 it`,
    rect: 'il it ir ib',
    path: 'M l vc L x1 y1 L x2 y1 L r vc L x2 y2 L x1 y2 Z',
  },
  heptagon: {
    av: 'hf 102572 vf 105210',
    gd: `swd2 */ wd2 hf 100000
shd2 */ hd2 vf 100000
svc */ vc vf 100000
dx1 */ swd2 97493 100000
dx2 */ swd2 78183 100000
dx3 */ swd2 43388 100000
dy1 */ shd2 62349 100000
dy2 */ shd2 22252 100000
dy3 */ shd2 90097 100000
x1 +- hc 0 dx1
x2 +- hc 0 dx2
x3 +- hc 0 dx3
x4 +- hc dx3 0
x5 +- hc dx2 0
x6 +- hc dx1 0
y1 +- svc 0 dy1
y2 +- svc dy2 0
y3 +- svc dy3 0
ib +- b 0 y1`,
    rect: 'x2 y1 x5 ib',
    path: 'M x1 y2 L x2 y1 L hc t L x5 y1 L x6 y2 L x4 y3 L x3 y3 Z',
  },
  octagon: {
    av: 'adj 29289',
    gd: `a pin 0 adj 50000
x1 */ ss a 100000
x2 +- r 0 x1
y2 +- b 0 x1
il */ x1 1 2
ir +- r 0 il
ib +- b 0 il`,
    rect: 'il il ir ib',
    path: 'M l x1 L x1 t L x2 t L r x1 L r y2 L x2 b L x1 b L l y2 Z',
  },
  decagon: {
    av: 'vf 105146',
    gd: `shd2 */ hd2 vf 100000
dx1 cos wd2 2160000
dx2 cos wd2 4320000
x1 +- hc 0 dx1
x2 +- hc 0 dx2
x3 +- hc dx2 0
x4 +- hc dx1 0
dy1 sin shd2 4320000
dy2 sin shd2 2160000
y1 +- vc 0 dy1
y2 +- vc 0 dy2
y3 +- vc dy2 0
y4 +- vc dy1 0`,
    rect: 'x1 y2 x4 y3',
    path: 'M l vc L x1 y2 L x2 y1 L x3 y1 L x4 y2 L r vc L x4 y3 L x3 y4 L x2 y4 L x1 y3 Z',
  },
  dodecagon: {
    gd: `x1 */ w 2894 21600
x2 */ w 7906 21600
x3 */ w 13694 21600
x4 */ w 18706 21600
y1 */ h 2894 21600
y2 */ h 7906 21600
y3 */ h 13694 21600
y4 */ h 18706 21600`,
    rect: 'x1 y1 x4 y4',
    path: 'M l y2 L x1 y1 L x2 t L x3 t L x4 y1 L r y2 L r y3 L x4 y4 L x3 b L x2 b L x1 y4 L l y3 Z',
  },

  // --- Stars ----------------------------------------------------------------------------------
  star4: {
    av: 'adj 12500',
    gd: `a pin 0 adj 50000
iwd2 */ wd2 a 50000
ihd2 */ hd2 a 50000
sdx cos iwd2 2700000
sdy sin ihd2 2700000
sx1 +- hc 0 sdx
sx2 +- hc sdx 0
sy1 +- vc 0 sdy
sy2 +- vc sdy 0
yAdj +- vc 0 ihd2`,
    rect: 'sx1 sy1 sx2 sy2',
    path: 'M l vc L sx1 sy1 L hc t L sx2 sy1 L r vc L sx2 sy2 L hc b L sx1 sy2 Z',
  },
  star5: {
    av: 'adj 19098 hf 105146 vf 110557',
    gd: `a pin 0 adj 50000
swd2 */ wd2 hf 100000
shd2 */ hd2 vf 100000
svc */ vc vf 100000
dx1 cos swd2 1080000
dx2 cos swd2 18360000
dy1 sin shd2 1080000
dy2 sin shd2 18360000
x1 +- hc 0 dx1
x2 +- hc 0 dx2
x3 +- hc dx2 0
x4 +- hc dx1 0
y1 +- svc 0 dy1
y2 +- svc 0 dy2
iwd2 */ swd2 a 50000
ihd2 */ shd2 a 50000
sdx1 cos iwd2 20520000
sdx2 cos iwd2 3240000
sdy1 sin ihd2 3240000
sdy2 sin ihd2 20520000
sx1 +- hc 0 sdx1
sx2 +- hc 0 sdx2
sx3 +- hc sdx2 0
sx4 +- hc sdx1 0
sy1 +- svc 0 sdy1
sy2 +- svc 0 sdy2
sy3 +- svc ihd2 0
yAdj +- svc 0 ihd2`,
    rect: 'sx1 sy1 sx4 sy3',
    path: 'M x1 y1 L sx2 sy1 L hc t L sx3 sy1 L x4 y1 L sx4 sy2 L x3 y2 L hc sy3 L x2 y2 L sx1 sy2 Z',
  },
  star6: {
    av: 'adj 28868 hf 115470',
    gd: `a pin 0 adj 50000
swd2 */ wd2 hf 100000
dx1 cos swd2 1800000
x1 +- hc 0 dx1
x2 +- hc dx1 0
y2 +- vc hd4 0
iwd2 */ swd2 a 50000
ihd2 */ hd2 a 50000
sdx2 */ iwd2 1 2
sx1 +- hc 0 iwd2
sx2 +- hc 0 sdx2
sx3 +- hc sdx2 0
sx4 +- hc iwd2 0
sdy1 sin ihd2 3600000
sy1 +- vc 0 sdy1
sy2 +- vc sdy1 0
yAdj +- vc 0 ihd2`,
    rect: 'x1 hd4 x2 y2',
    path: 'M x1 hd4 L sx2 sy1 L hc t L sx3 sy1 L x2 hd4 L sx4 vc L x2 y2 L sx3 sy2 L hc b L sx2 sy2 L x1 y2 L sx1 vc Z',
  },
  star7: {
    av: 'adj 34601 hf 102572 vf 105210',
    gd: `a pin 0 adj 50000
swd2 */ wd2 hf 100000
shd2 */ hd2 vf 100000
svc */ vc vf 100000
dx1 */ swd2 97493 100000
dx2 */ swd2 78183 100000
dx3 */ swd2 43388 100000
dy1 */ shd2 62349 100000
dy2 */ shd2 22252 100000
dy3 */ shd2 90097 100000
x1 +- hc 0 dx1
x2 +- hc 0 dx2
x3 +- hc 0 dx3
x4 +- hc dx3 0
x5 +- hc dx2 0
x6 +- hc dx1 0
y1 +- svc 0 dy1
y2 +- svc dy2 0
y3 +- svc dy3 0
iwd2 */ swd2 a 50000
ihd2 */ shd2 a 50000
sdx1 */ iwd2 97493 100000
sdx2 */ iwd2 78183 100000
sdx3 */ iwd2 43388 100000
sx1 +- hc 0 sdx1
sx2 +- hc 0 sdx2
sx3 +- hc 0 sdx3
sx4 +- hc sdx3 0
sx5 +- hc sdx2 0
sx6 +- hc sdx1 0
sdy1 */ ihd2 90097 100000
sdy2 */ ihd2 22252 100000
sdy3 */ ihd2 62349 100000
sy1 +- svc 0 sdy1
sy2 +- svc 0 sdy2
sy3 +- svc sdy3 0
sy4 +- svc ihd2 0
yAdj +- svc 0 ihd2`,
    rect: 'sx2 sy1 sx5 sy3',
    path: 'M x1 y2 L sx1 sy2 L x2 y1 L sx3 sy1 L hc t L sx4 sy1 L x5 y1 L sx6 sy2 L x6 y2 L sx5 sy3 L x4 y3 L hc sy4 L x3 y3 L sx2 sy3 Z',
  },
  star8: {
    av: 'adj 38250',
    gd: `a pin 0 adj 50000
dx1 cos wd2 2700000
x1 +- hc 0 dx1
x2 +- hc dx1 0
dy1 sin hd2 2700000
y1 +- vc 0 dy1
y2 +- vc dy1 0
iwd2 */ wd2 a 50000
ihd2 */ hd2 a 50000
sdx1 */ iwd2 92388 100000
sdx2 */ iwd2 38268 100000
sdy1 */ ihd2 92388 100000
sdy2 */ ihd2 38268 100000
sx1 +- hc 0 sdx1
sx2 +- hc 0 sdx2
sx3 +- hc sdx2 0
sx4 +- hc sdx1 0
sy1 +- vc 0 sdy1
sy2 +- vc 0 sdy2
sy3 +- vc sdy2 0
sy4 +- vc sdy1 0
yAdj +- vc 0 ihd2`,
    rect: 'sx1 sy1 sx4 sy4',
    path: 'M l vc L sx1 sy2 L x1 y1 L sx2 sy1 L hc t L sx3 sy1 L x2 y1 L sx4 sy2 L r vc L sx4 sy3 L x2 y2 L sx3 sy4 L hc b L sx2 sy4 L x1 y2 L sx1 sy3 Z',
  },
  star10: {
    av: 'adj 42533 hf 105146',
    gd: `a pin 0 adj 50000
swd2 */ wd2 hf 100000
dx1 */ swd2 95106 100000
dx2 */ swd2 58779 100000
x1 +- hc 0 dx1
x2 +- hc 0 dx2
x3 +- hc dx2 0
x4 +- hc dx1 0
dy1 */ hd2 80902 100000
dy2 */ hd2 30902 100000
y1 +- vc 0 dy1
y2 +- vc 0 dy2
y3 +- vc dy2 0
y4 +- vc dy1 0
iwd2 */ swd2 a 50000
ihd2 */ hd2 a 50000
sdx1 */ iwd2 80902 100000
sdx2 */ iwd2 30902 100000
sdy1 */ ihd2 95106 100000
sdy2 */ ihd2 58779 100000
sx1 +- hc 0 iwd2
sx2 +- hc 0 sdx1
sx3 +- hc 0 sdx2
sx4 +- hc sdx2 0
sx5 +- hc sdx1 0
sx6 +- hc iwd2 0
sy1 +- vc 0 sdy1
sy2 +- vc 0 sdy2
sy3 +- vc sdy2 0
sy4 +- vc sdy1 0
yAdj +- vc 0 ihd2`,
    rect: 'sx2 y2 sx5 y3',
    path: 'M x1 y2 L sx2 sy2 L x2 y1 L sx3 sy1 L hc t L sx4 sy1 L x3 y1 L sx5 sy2 L x4 y2 L sx6 vc L x4 y3 L sx5 sy3 L x3 y4 L sx4 sy4 L hc b L sx3 sy4 L x2 y4 L sx2 sy3 L x1 y3 L sx1 vc Z',
  },
  star12: {
    av: 'adj 37500',
    gd: `a pin 0 adj 50000
dx1 cos wd2 1800000
dy1 sin hd2 3600000
x1 +- hc 0 dx1
x3 */ w 3 4
x4 +- hc dx1 0
y1 +- vc 0 dy1
y3 */ h 3 4
y4 +- vc dy1 0
iwd2 */ wd2 a 50000
ihd2 */ hd2 a 50000
sdx1 cos iwd2 900000
sdx2 cos iwd2 2700000
sdx3 cos iwd2 4500000
sdy1 sin ihd2 4500000
sdy2 sin ihd2 2700000
sdy3 sin ihd2 900000
sx1 +- hc 0 sdx1
sx2 +- hc 0 sdx2
sx3 +- hc 0 sdx3
sx4 +- hc sdx3 0
sx5 +- hc sdx2 0
sx6 +- hc sdx1 0
sy1 +- vc 0 sdy1
sy2 +- vc 0 sdy2
sy3 +- vc 0 sdy3
sy4 +- vc sdy3 0
sy5 +- vc sdy2 0
sy6 +- vc sdy1 0
yAdj +- vc 0 ihd2`,
    rect: 'sx2 sy2 sx5 sy5',
    path:
      'M l vc L sx1 sy3 L x1 hd4 L sx2 sy2 L wd4 y1 L sx3 sy1 L hc t L sx4 sy1 L x3 y1 L sx5 sy2 L x4 hd4 L sx6 sy3 ' +
      'L r vc L sx6 sy4 L x4 y3 L sx5 sy5 L x3 y4 L sx4 sy6 L hc b L sx3 sy6 L wd4 y4 L sx2 sy5 L x1 y3 L sx1 sy4 Z',
  },
  star16: regularStar(16, 37500),
  star24: regularStar(24, 37500),
  star32: regularStar(32, 37500),

  // --- Crosses and math -----------------------------------------------------------------------
  plus: {
    av: 'adj 25000',
    gd: `a pin 0 adj 50000
x1 */ ss a 100000
x2 +- r 0 x1
y2 +- b 0 x1
d +- w 0 h
il ?: d l x1
ir ?: d r x2
it ?: d x1 t
ib ?: d y2 b`,
    rect: 'il it ir ib',
    path: 'M l x1 L x1 x1 L x1 t L x2 t L x2 x1 L r x1 L r y2 L x2 y2 L x2 b L x1 b L x1 y2 L l y2 Z',
  },
  mathPlus: {
    av: 'adj1 23520',
    gd: `a1 pin 0 adj1 73490
dx1 */ w 73490 200000
dy1 */ h 73490 200000
dx2 */ ss a1 200000
x1 +- hc 0 dx1
x2 +- hc 0 dx2
x3 +- hc dx2 0
x4 +- hc dx1 0
y1 +- vc 0 dy1
y2 +- vc 0 dx2
y3 +- vc dx2 0
y4 +- vc dy1 0`,
    rect: 'x1 y2 x4 y3',
    path: 'M x1 y2 L x2 y2 L x2 y1 L x3 y1 L x3 y2 L x4 y2 L x4 y3 L x3 y3 L x3 y4 L x2 y4 L x2 y3 L x1 y3 Z',
  },
  mathMinus: {
    av: 'adj1 23520',
    gd: `a1 pin 0 adj1 100000
dy1 */ h a1 200000
dx1 */ w 73490 200000
y1 +- vc 0 dy1
y2 +- vc dy1 0
x1 +- hc 0 dx1
x2 +- hc dx1 0`,
    rect: 'x1 y1 x2 y2',
    path: 'M x1 y1 L x2 y1 L x2 y2 L x1 y2 Z',
  },
  mathMultiply: {
    av: 'adj1 23520',
    gd: `a1 pin 0 adj1 51965
th */ ss a1 100000
a at2 w h
sa sin 1 a
ca cos 1 a
ta tan 1 a
dl mod w h 0
rw */ dl 51965 100000
lM +- dl 0 rw
xM */ ca lM 2
yM */ sa lM 2
dxAM */ sa th 2
dyAM */ ca th 2
xA +- xM 0 dxAM
yA +- yM dyAM 0
xB +- xM dxAM 0
yB +- yM 0 dyAM
xBC +- hc 0 xB
yBC */ xBC ta 1
yC +- yBC yB 0
xD +- r 0 xB
xE +- r 0 xA
yFE +- vc 0 yA
xFE */ yFE 1 ta
xF +- xE 0 xFE
xL +- xA xFE 0
yG +- b 0 yA
yH +- b 0 yB
yI +- b 0 yC
xC2 +- r 0 xL
yC3 +- b 0 yC`,
    rect: 'xA yB xE yH',
    path: 'M xA yA L xB yB L hc yC L xD yB L xE yA L xF vc L xE yG L xD yH L hc yI L xB yH L xA yG L xL vc Z',
  },
  mathDivide: {
    av: 'adj1 23520 adj2 5880 adj3 11760',
    gd: `a1 pin 1000 adj1 36745
ma1 +- 0 0 a1
ma3h +/ 73490 ma1 4
ma3w */ 36745 w h
maxAdj3 min ma3h ma3w
a3 pin 1000 adj3 maxAdj3
m4a3 */ -4 a3 1
maxAdj2 +- 73490 m4a3 a1
a2 pin 0 adj2 maxAdj2
dy1 */ h a1 200000
yg */ h a2 100000
rad */ h a3 100000
dx1 */ w 73490 200000
y3 +- vc 0 dy1
y4 +- vc dy1 0
a +- yg rad 0
y2 +- y3 0 a
y1 +- y2 0 rad
y5 +- b 0 y1
x1 +- hc 0 dx1
x3 +- hc dx1 0
x2 +- hc 0 rad`,
    rect: 'x1 y3 x3 y4',
    path: 'M hc y1 A rad rad 3cd4 21600000 Z M hc y5 A rad rad cd4 21600000 Z M x1 y3 L x3 y3 L x3 y4 L x1 y4 Z',
  },
  mathEqual: {
    av: 'adj1 23520 adj2 11760',
    gd: `a1 pin 0 adj1 36745
2a1 */ a1 2 1
mAdj2 +- 100000 0 2a1
a2 pin 0 adj2 mAdj2
dy1 */ h a1 100000
dy2 */ h a2 200000
dx1 */ w 73490 200000
y2 +- vc 0 dy2
y3 +- vc dy2 0
y1 +- y2 0 dy1
y4 +- y3 dy1 0
x1 +- hc 0 dx1
x2 +- hc dx1 0
yC1 +/ y1 y2 2
yC2 +/ y3 y4 2`,
    rect: 'x1 y1 x2 y4',
    path: 'M x1 y1 L x2 y1 L x2 y2 L x1 y2 Z M x1 y3 L x2 y3 L x2 y4 L x1 y4 Z',
  },
  mathNotEqual: {
    av: 'adj1 23520 adj2 6600000 adj3 11760',
    gd: `a1 pin 0 adj1 50000
crAng pin 4200000 adj2 6600000
2a1 */ a1 2 1
maxAdj3 +- 100000 0 2a1
a3 pin 0 adj3 maxAdj3
dy1 */ h a1 100000
dy2 */ h a3 200000
dx1 */ w 73490 200000
x1 +- hc 0 dx1
x8 +- hc dx1 0
y2 +- vc 0 dy2
y3 +- vc dy2 0
y1 +- y2 0 dy1
y4 +- y3 dy1 0
cadj2 +- crAng 0 cd4
xadj2 tan hd2 cadj2
len mod xadj2 hd2 0
bhw */ len dy1 hd2
bhw2 */ bhw 1 2
x7 +- hc xadj2 bhw2
dx67 */ xadj2 y1 hd2
x6 +- x7 0 dx67
dx57 */ xadj2 y2 hd2
x5 +- x7 0 dx57
dx47 */ xadj2 y3 hd2
x4 +- x7 0 dx47
dx37 */ xadj2 y4 hd2
x3 +- x7 0 dx37
dx27 */ xadj2 2 1
x2 +- x7 0 dx27
rx7 +- x7 bhw 0
rx6 +- x6 bhw 0
rx5 +- x5 bhw 0
rx4 +- x4 bhw 0
rx3 +- x3 bhw 0
rx2 +- x2 bhw 0
dx7 */ dy1 hd2 len
rxt +- x7 dx7 0
lxt +- rx7 0 dx7
rx ?: cadj2 rxt rx7
lx ?: cadj2 x7 lxt
dy3 */ dy1 xadj2 len
dy4 +- 0 0 dy3
ry ?: cadj2 dy3 t
ly ?: cadj2 t dy4
dlx +- w 0 rx
drx +- w 0 lx
dly +- h 0 ry
dry +- h 0 ly`,
    rect: 'x1 y1 x8 y4',
    path:
      'M x1 y1 L x6 y1 L lx ly L rx ry L rx6 y1 L x8 y1 L x8 y2 L rx5 y2 L rx4 y3 L x8 y3 L x8 y4 L rx3 y4 ' +
      'L drx dry L dlx dly L x3 y4 L x1 y4 L x1 y3 L x4 y3 L x5 y2 L x1 y2 Z',
  },

  // --- Rings, arcs and solids -----------------------------------------------------------------
  donut: {
    av: 'adj 25000',
    gd: `a pin 0 adj 50000
dr */ ss a 100000
iwd2 +- wd2 0 dr
ihd2 +- hd2 0 dr
${ELLIPSE_RECT_GD}`,
    rect: 'il it ir ib',
    path:
      `${ELLIPSE_PATH} M dr vc A iwd2 ihd2 cd2 -5400000 A iwd2 ihd2 cd4 -5400000 ` +
      'A iwd2 ihd2 0 -5400000 A iwd2 ihd2 3cd4 -5400000 Z',
  },
  noSmoking: {
    av: 'adj 18750',
    gd: `a pin 0 adj 50000
dr */ ss a 100000
iwd2 +- wd2 0 dr
ihd2 +- hd2 0 dr
ang at2 w h
ct cos ihd2 ang
st sin iwd2 ang
m mod ct st 0
n */ iwd2 ihd2 m
drd2 */ dr 1 2
dang at2 n drd2
2dang */ dang 2 1
swAng +- -10800000 2dang 0
t3 at2 w h
stAng1 +- t3 0 dang
stAng2 +- stAng1 0 cd2
ct1 cos ihd2 stAng1
st1 sin iwd2 stAng1
m1 mod ct1 st1 0
n1 */ iwd2 ihd2 m1
dx1 cos n1 stAng1
dy1 sin n1 stAng1
x1 +- hc dx1 0
y1 +- vc dy1 0
x2 +- hc 0 dx1
y2 +- vc 0 dy1
${ELLIPSE_RECT_GD}`,
    rect: 'il it ir ib',
    path: `${ELLIPSE_PATH} M x1 y1 A iwd2 ihd2 stAng1 swAng Z M x2 y2 A iwd2 ihd2 stAng2 swAng Z`,
  },
  blockArc: {
    av: 'adj1 10800000 adj2 0 adj3 25000',
    gd: `stAng pin 0 adj1 21599999
istAng pin 0 adj2 21599999
a3 pin 0 adj3 50000
sw11 +- istAng 0 stAng
sw12 +- sw11 21600000 0
swAng ?: sw11 sw11 sw12
iswAng +- 0 0 swAng
wt1 sin wd2 stAng
ht1 cos hd2 stAng
wt3 sin wd2 istAng
ht3 cos hd2 istAng
dx1 cat2 wd2 ht1 wt1
dy1 sat2 hd2 ht1 wt1
dx3 cat2 wd2 ht3 wt3
dy3 sat2 hd2 ht3 wt3
x1 +- hc dx1 0
y1 +- vc dy1 0
x3 +- hc dx3 0
y3 +- vc dy3 0
dr */ ss a3 100000
iwd2 +- wd2 0 dr
ihd2 +- hd2 0 dr
wt2 sin iwd2 istAng
ht2 cos ihd2 istAng
wt4 sin iwd2 stAng
ht4 cos ihd2 stAng
dx2 cat2 iwd2 ht2 wt2
dy2 sat2 ihd2 ht2 wt2
dx4 cat2 iwd2 ht4 wt4
dy4 sat2 ihd2 ht4 wt4
x2 +- hc dx2 0
y2 +- vc dy2 0
x4 +- hc dx4 0
y4 +- vc dy4 0
sw0 +- 21600000 0 stAng
da1 +- swAng 0 sw0
g1 max x1 x2
g2 max x3 x4
g3 max g1 g2
ir ?: da1 r g3
sw1 +- cd4 0 stAng
sw2 +- 27000000 0 stAng
sw3 ?: sw1 sw1 sw2
da2 +- swAng 0 sw3
g5 max y1 y2
g6 max y3 y4
g7 max g5 g6
ib ?: da2 b g7
sw4 +- cd2 0 stAng
sw5 +- 32400000 0 stAng
sw6 ?: sw4 sw4 sw5
da3 +- swAng 0 sw6
g9 min x1 x2
g10 min x3 x4
g11 min g9 g10
il ?: da3 l g11
sw7 +- 3cd4 0 stAng
sw8 +- 37800000 0 stAng
sw9 ?: sw7 sw7 sw8
da4 +- swAng 0 sw9
g13 min y1 y2
g14 min y3 y4
g15 min g13 g14
it ?: da4 t g15`,
    rect: 'il it ir ib',
    path: 'M x1 y1 A wd2 hd2 stAng swAng L x2 y2 A iwd2 ihd2 istAng iswAng Z',
  },
  can: {
    av: 'adj 25000',
    gd: `maxAdj */ 50000 h ss
a pin 0 adj maxAdj
y1 */ ss a 200000
y2 +- y1 y1 0
y3 +- b 0 y1`,
    rect: 'l y2 r y3',
    path: [
      'nostroke noext | M l y1 A wd2 y1 cd2 -10800000 L r y3 A wd2 y1 0 cd4 A wd2 y1 cd4 cd4 Z',
      'lighten nostroke noext | M l y1 A wd2 y1 cd2 cd2 A wd2 y1 0 cd2 Z',
      'none | M r y1 A wd2 y1 0 cd2 A wd2 y1 cd2 cd2 L r y3 A wd2 y1 0 cd4 A wd2 y1 cd4 cd4 L l y1',
    ],
  },
  cube: {
    av: 'adj 25000',
    gd: `a pin 0 adj 100000
y1 */ ss a 100000
y4 +- b 0 y1
y2 */ y4 1 2
y3 +/ y1 b 2
x4 +- r 0 y1
x2 */ x4 1 2
x3 +/ y1 r 2`,
    rect: 'l y1 x4 b',
    path: [
      'nostroke noext | M l y1 L x4 y1 L x4 b L l b Z',
      'darkenLess nostroke noext | M x4 y1 L r t L r y4 L x4 b Z',
      'lightenLess nostroke noext | M l y1 L y1 t L r t L x4 y1 Z',
      'none noext | M l y1 L y1 t L r t L r y4 L x4 b L l b Z M l y1 L x4 y1 L r t M x4 y1 L x4 b',
    ],
  },
  bevel: {
    av: 'adj 12500',
    gd: `a pin 0 adj 50000
x1 */ ss a 100000
x2 +- r 0 x1
y2 +- b 0 x1`,
    rect: 'x1 x1 x2 y2',
    path: [
      'nostroke noext | M x1 x1 L x2 x1 L x2 y2 L x1 y2 Z',
      'lightenLess nostroke noext | M l t L r t L x2 x1 L x1 x1 Z',
      'darkenLess nostroke noext | M l b L x1 y2 L x2 y2 L r b Z',
      'lighten nostroke noext | M l t L x1 x1 L x1 y2 L l b Z',
      'darken nostroke noext | M r t L r b L x2 y2 L x2 x1 Z',
      'none noext | M l t L r t L r b L l b Z M x1 x1 L x2 x1 L x2 y2 L x1 y2 Z M l t L x1 x1 M l b L x1 y2 M r t L x2 x1 M r b L x2 y2',
    ],
  },
  frame: {
    av: 'adj1 12500',
    gd: `a1 pin 0 adj1 50000
x1 */ ss a1 100000
x4 +- r 0 x1
y4 +- b 0 x1`,
    rect: 'x1 x1 x4 y4',
    path: 'M l t L r t L r b L l b Z M x1 x1 L x1 y4 L x4 y4 L x4 x1 Z',
  },
  halfFrame: {
    av: 'adj1 33333 adj2 33333',
    gd: `maxAdj2 */ 100000 w ss
a2 pin 0 adj2 maxAdj2
x1 */ ss a2 100000
g1 */ h x1 w
g2 +- h 0 g1
maxAdj1 */ 100000 g2 ss
a1 pin 0 adj1 maxAdj1
y1 */ ss a1 100000
dx2 */ y1 w h
x2 +- r 0 dx2
dy2 */ x1 h w
y2 +- b 0 dy2`,
    path: 'M l t L r t L x2 y1 L x1 y1 L x1 y2 L l b Z',
  },
  corner: {
    av: 'adj1 50000 adj2 50000',
    gd: `maxAdj1 */ 100000 h ss
maxAdj2 */ 100000 w ss
a1 pin 0 adj1 maxAdj1
a2 pin 0 adj2 maxAdj2
x1 */ ss a2 100000
dy1 */ ss a1 100000
y1 +- b 0 dy1
d +- w 0 h
it ?: d y1 t
ir ?: d r x1`,
    rect: 'l it ir b',
    path: 'M l t L x1 t L x1 y1 L r y1 L r b L l b Z',
  },
  diagStripe: {
    av: 'adj 50000',
    gd: `a pin 0 adj 100000
x2 */ w a 100000
x1 */ x2 1 2
x3 +/ x2 r 2
y2 */ h a 100000
y1 */ y2 1 2
y3 +/ y2 b 2`,
    rect: 'l t x3 y3',
    path: 'M l y2 L x2 t L r t L l b Z',
  },
  plaque: {
    av: 'adj 16667',
    gd: `a pin 0 adj 50000
x1 */ ss a 100000
x2 +- r 0 x1
y2 +- b 0 x1
il */ x1 70711 100000
ir +- r 0 il
ib +- b 0 il`,
    rect: 'il il ir ib',
    path: 'M l x1 A x1 x1 cd4 -5400000 L x2 t A x1 x1 cd2 -5400000 L r y2 A x1 x1 3cd4 -5400000 L x1 b A x1 x1 0 -5400000 Z',
  },
  foldedCorner: {
    av: 'adj 16667',
    gd: `a pin 0 adj 50000
dy2 */ ss a 100000
dy1 */ dy2 1 5
x1 +- r 0 dy2
x2 +- x1 dy1 0
y2 +- b 0 dy2
y1 +- y2 dy1 0`,
    rect: 'l t r y2',
    path: [
      'nostroke noext | M l t L r t L r y2 L x1 b L l b Z',
      'darkenLess nostroke noext | M x1 b L x2 y1 L r y2 Z',
      'none noext | M x1 b L x2 y1 L r y2 L x1 b L l b L l t L r t L r y2',
    ],
  },
  smileyFace: {
    av: 'adj 4653',
    gd: `a pin -4653 adj 4653
x1 */ w 4969 21699
x2 */ w 6215 21600
x3 */ w 13135 21600
x4 */ w 16640 21600
y1 */ h 7570 21600
y3 */ h 16515 21600
dy2 */ h a 100000
y2 +- y3 0 dy2
y4 +- y3 dy2 0
dy3 */ h a 50000
y5 +- y4 dy3 0
${ELLIPSE_RECT_GD}
wR */ w 1125 21600
hR */ h 1125 21600`,
    rect: 'il it ir ib',
    path: [
      'nostroke noext | M l vc A wd2 hd2 cd2 21600000 Z',
      'darkenLess noext | M x2 y1 A wR hR cd2 21600000 M x3 y1 A wR hR cd2 21600000',
      'none noext | M x1 y2 Q hc y5 x4 y2',
      'none | M l vc A wd2 hd2 cd2 21600000 Z',
    ],
  },
  heart: {
    gd: `dx1 */ w 49 48
dx2 */ w 10 48
x1 +- hc 0 dx1
x2 +- hc 0 dx2
x3 +- hc dx2 0
x4 +- hc dx1 0
y1 +- t 0 hd3
il */ w 1 6
ir */ w 5 6
ib */ h 2 3`,
    rect: 'il hd4 ir ib',
    path: 'M hc hd4 C x3 y1 x4 hd4 hc b C x1 hd4 x2 y1 hc hd4 Z',
  },
  lightningBolt: {
    gd: `x1 */ w 5022 21600
x3 */ w 8472 21600
x4 */ w 8757 21600
x5 */ w 10012 21600
x8 */ w 12860 21600
x9 */ w 13917 21600
x11 */ w 16577 21600
y1 */ h 3890 21600
y2 */ h 6080 21600
y4 */ h 7437 21600
y6 */ h 9705 21600
y7 */ h 12007 21600
y10 */ h 14277 21600
y11 */ h 14915 21600`,
    rect: 'x4 y4 x9 y10',
    path:
      'w=21600 h=21600 | M 8472 0 L 12860 6080 L 11050 6797 L 16577 12007 L 14767 12877 L 21600 21600 ' +
      'L 10012 14915 L 12222 13987 L 5022 9705 L 7602 8382 L 0 3890 Z',
  },
  sun: {
    av: 'adj 25000',
    gd: `a pin 12500 adj 46875
g0 +- 50000 0 a
g1 */ g0 30274 32768
g2 */ g0 12540 32768
g3 +- g1 50000 0
g4 +- g2 50000 0
g5 +- 50000 0 g1
g6 +- 50000 0 g2
g7 */ g0 23170 32768
g8 +- 50000 g7 0
g9 +- 50000 0 g7
g10 */ g5 3 4
g11 */ g6 3 4
g12 +- g10 3662 0
g13 +- g11 3662 0
g14 +- g11 12500 0
g15 +- 100000 0 g10
g16 +- 100000 0 g12
g17 +- 100000 0 g13
g18 +- 100000 0 g14
ox1 */ w 18436 21600
oy1 */ h 3163 21600
ox2 */ w 3163 21600
oy2 */ h 18436 21600
x8 */ w g8 100000
x9 */ w g9 100000
x10 */ w g10 100000
x12 */ w g12 100000
x13 */ w g13 100000
x14 */ w g14 100000
x15 */ w g15 100000
x16 */ w g16 100000
x17 */ w g17 100000
x18 */ w g18 100000
x19 */ w a 100000
wR */ w g0 100000
hR */ h g0 100000
y8 */ h g8 100000
y9 */ h g9 100000
y10 */ h g10 100000
y12 */ h g12 100000
y13 */ h g13 100000
y14 */ h g14 100000
y15 */ h g15 100000
y16 */ h g16 100000
y17 */ h g17 100000
y18 */ h g18 100000`,
    rect: 'x9 y9 x8 y8',
    path:
      'M r vc L x15 y18 L x15 y14 Z M ox1 oy1 L x16 y13 L x17 y12 Z M hc t L x18 y10 L x14 y10 Z ' +
      'M ox2 oy1 L x13 y12 L x12 y13 Z M l vc L x10 y14 L x10 y18 Z M ox2 oy2 L x12 y17 L x13 y16 Z ' +
      'M hc b L x14 y15 L x18 y15 Z M ox1 oy2 L x17 y16 L x16 y17 Z M x19 vc A wR hR cd2 21600000 Z',
  },
  moon: {
    av: 'adj 50000',
    gd: `a pin 0 adj 87500
g0 */ ss a 100000
g0w */ g0 w ss
g1 +- ss 0 g0
g2 */ g0 g0 g1
g3 */ ss ss g1
g4 */ g3 2 1
g5 +- g4 0 g2
g6 +- g5 0 g0
g6w */ g6 w ss
g7 */ g5 1 2
g8 +- g7 0 g0
dy1 */ g8 hd2 ss
g10h +- vc 0 dy1
g11h +- vc dy1 0
g12 */ g0 9598 32768
g12w */ g12 w ss
g13 +- ss 0 g12
q1 */ ss ss 1
q2 */ g13 g13 1
q3 +- q1 0 q2
q4 sqrt q3
dy4 */ q4 hd2 ss
g15h +- vc 0 dy4
g16h +- vc dy4 0
g17w +- g6w 0 g0w
g18w */ g17w 1 2
dx2p +- g0w g18w w
dx2 */ dx2p -1 1
dy2 */ hd2 -1 1
stAng1 at2 dx2 dy2
enAngp1 at2 dx2 hd2
enAng1 +- enAngp1 0 21600000
swAng1 +- enAng1 0 stAng1`,
    rect: 'g12w g15h g0w g16h',
    path: 'M r b A w hd2 cd4 cd2 A g18w dy1 stAng1 swAng1 Z',
  },
  cloud: {
    gd: CLOUD_RECT_GD,
    rect: 'il it ir ib',
    path: [CLOUD_OUTLINE, CLOUD_DETAIL],
  },
  arc: {
    av: 'adj1 16200000 adj2 0',
    gd: `stAng pin 0 adj1 21599999
enAng pin 0 adj2 21599999
sw11 +- enAng 0 stAng
sw12 +- sw11 21600000 0
swAng ?: sw11 sw11 sw12
wt1 sin wd2 stAng
ht1 cos hd2 stAng
dx1 cat2 wd2 ht1 wt1
dy1 sat2 hd2 ht1 wt1
wt2 sin wd2 enAng
ht2 cos hd2 enAng
dx2 cat2 wd2 ht2 wt2
dy2 sat2 hd2 ht2 wt2
x1 +- hc dx1 0
y1 +- vc dy1 0
x2 +- hc dx2 0
y2 +- vc dy2 0
sw0 +- 21600000 0 stAng
da1 +- swAng 0 sw0
g1 max x1 x2
ir ?: da1 r g1
sw1 +- cd4 0 stAng
sw2 +- 27000000 0 stAng
sw3 ?: sw1 sw1 sw2
da2 +- swAng 0 sw3
g5 max y1 y2
ib ?: da2 b g5
sw4 +- cd2 0 stAng
sw5 +- 32400000 0 stAng
sw6 ?: sw4 sw4 sw5
da3 +- swAng 0 sw6
g9 min x1 x2
il ?: da3 l g9
sw7 +- 3cd4 0 stAng
sw8 +- 37800000 0 stAng
sw9 ?: sw7 sw7 sw8
da4 +- swAng 0 sw9
g13 min y1 y2
it ?: da4 t g13`,
    rect: 'il it ir ib',
    path: ['nostroke noext | M x1 y1 A wd2 hd2 stAng swAng L hc vc Z', 'none | M x1 y1 A wd2 hd2 stAng swAng'],
  },
  chord: {
    av: 'adj1 2700000 adj2 16200000',
    gd: `stAng pin 0 adj1 21599999
enAng pin 0 adj2 21599999
sw1 +- enAng 0 stAng
sw2 +- sw1 21600000 0
swAng ?: sw1 sw1 sw2
wt1 sin wd2 stAng
ht1 cos hd2 stAng
dx1 cat2 wd2 ht1 wt1
dy1 sat2 hd2 ht1 wt1
wt2 sin wd2 enAng
ht2 cos hd2 enAng
dx2 cat2 wd2 ht2 wt2
dy2 sat2 hd2 ht2 wt2
x1 +- hc dx1 0
y1 +- vc dy1 0
x2 +- hc dx2 0
y2 +- vc dy2 0
x3 +/ x1 x2 2
y3 +/ y1 y2 2
midAng0 */ swAng 1 2
midAng +- stAng midAng0 cd2
${ELLIPSE_RECT_GD}`,
    rect: 'il it ir ib',
    path: 'M x1 y1 A wd2 hd2 stAng swAng Z',
  },
  pie: {
    av: 'adj1 0 adj2 16200000',
    gd: `stAng pin 0 adj1 21599999
enAng pin 0 adj2 21599999
sw1 +- enAng 0 stAng
sw2 +- sw1 21600000 0
swAng ?: sw1 sw1 sw2
wt1 sin wd2 stAng
ht1 cos hd2 stAng
dx1 cat2 wd2 ht1 wt1
dy1 sat2 hd2 ht1 wt1
x1 +- hc dx1 0
y1 +- vc dy1 0
wt2 sin wd2 enAng
ht2 cos hd2 enAng
dx2 cat2 wd2 ht2 wt2
dy2 sat2 hd2 ht2 wt2
x2 +- hc dx2 0
y2 +- vc dy2 0
${ELLIPSE_RECT_GD}`,
    rect: 'il it ir ib',
    path: 'M x1 y1 A wd2 hd2 stAng swAng L hc vc Z',
  },
  pieWedge: {
    gd: `g1 cos w 13500000
g2 sin h 13500000
x1 +- r g1 0
y1 +- b g2 0`,
    rect: 'x1 y1 r b',
    path: 'M l b A w h cd2 cd4 L r b Z',
  },
  teardrop: {
    av: 'adj 100000',
    gd: `a pin 0 adj 200000
r2 sqrt 2
tw */ r2 wd2 1
th */ r2 hd2 1
sw */ tw a 100000
sh */ th a 100000
dx1 cos sw 2700000
dy1 sin sh 2700000
x1 +- hc dx1 0
y1 +- vc 0 dy1
x2 +/ hc x1 2
y2 +/ vc y1 2
${ELLIPSE_RECT_GD}`,
    rect: 'il it ir ib',
    path: 'M l vc A wd2 hd2 cd2 cd4 Q x2 t x1 y1 Q r y2 r vc A wd2 hd2 0 cd4 A wd2 hd2 cd4 cd4 Z',
  },

  // --- Brackets and braces --------------------------------------------------------------------
  bracketPair: {
    av: 'adj 16667',
    gd: `a pin 0 adj 50000
x1 */ ss a 100000
x2 +- r 0 x1
y2 +- b 0 x1
il */ x1 29289 100000
ir +- r 0 il
ib +- b 0 il`,
    rect: 'il il ir ib',
    path: [
      'nostroke noext | M l x1 A x1 x1 cd2 cd4 L x2 t A x1 x1 3cd4 cd4 L r y2 A x1 x1 0 cd4 L x1 b A x1 x1 cd4 cd4 Z',
      'none | M x1 b A x1 x1 cd4 cd4 L l x1 A x1 x1 cd2 cd4 M x2 t A x1 x1 3cd4 cd4 L r y2 A x1 x1 0 cd4',
    ],
  },
  bracePair: {
    av: 'adj 8333',
    gd: `a pin 0 adj 25000
x1 */ ss a 100000
x2 */ ss a 50000
x3 +- r 0 x2
x4 +- r 0 x1
y2 +- vc 0 x1
y3 +- vc x1 0
y4 +- b 0 x1
it */ x1 29289 100000
il +- x1 it 0
ir +- r 0 il
ib +- b 0 it`,
    rect: 'il il ir ib',
    path: [
      'nostroke noext | M x2 b A x1 x1 cd4 cd4 L x1 y3 A x1 x1 0 -5400000 A x1 x1 cd4 -5400000 L x1 x1 A x1 x1 cd2 cd4 ' +
        'L x3 t A x1 x1 3cd4 cd4 L x4 y2 A x1 x1 cd2 -5400000 A x1 x1 3cd4 -5400000 L x4 y4 A x1 x1 0 cd4 Z',
      'none | M x2 b A x1 x1 cd4 cd4 L x1 y3 A x1 x1 0 -5400000 A x1 x1 cd4 -5400000 L x1 x1 A x1 x1 cd2 cd4 ' +
        'M x3 t A x1 x1 3cd4 cd4 L x4 y2 A x1 x1 cd2 -5400000 A x1 x1 3cd4 -5400000 L x4 y4 A x1 x1 0 cd4',
    ],
  },
  leftBracket: {
    av: 'adj 8333',
    gd: `maxAdj */ 50000 h ss
a pin 0 adj maxAdj
y1 */ ss a 100000
y2 +- b 0 y1
dx1 cos w 2700000
dy1 sin y1 2700000
il +- r 0 dx1
it +- y1 0 dy1
ib +- b dy1 y1`,
    rect: 'il it r ib',
    path: [
      'nostroke noext | M r b A w y1 cd4 cd4 L l y1 A w y1 cd2 cd4 Z',
      'none | M r b A w y1 cd4 cd4 L l y1 A w y1 cd2 cd4',
    ],
  },
  rightBracket: {
    av: 'adj 8333',
    gd: `maxAdj */ 50000 h ss
a pin 0 adj maxAdj
y1 */ ss a 100000
y2 +- b 0 y1
dx1 cos w 2700000
dy1 sin y1 2700000
ir +- l dx1 0
it +- y1 0 dy1
ib +- b dy1 y1`,
    rect: 'l it ir ib',
    path: [
      'nostroke noext | M l t A w y1 3cd4 cd4 L r y2 A w y1 0 cd4 Z',
      'none | M l t A w y1 3cd4 cd4 L r y2 A w y1 0 cd4',
    ],
  },
  leftBrace: {
    av: 'adj1 8333 adj2 50000',
    gd: `a2 pin 0 adj2 100000
q1 +- 100000 0 a2
q2 min q1 a2
q3 */ q2 1 2
maxAdj1 */ q3 h ss
a1 pin 0 adj1 maxAdj1
y1 */ ss a1 100000
y3 */ h a2 100000
y4 +- y3 y1 0
dx1 cos wd2 2700000
dy1 sin y1 2700000
il +- r 0 dx1
it +- y1 0 dy1
ib +- b dy1 y1`,
    rect: 'il it r ib',
    path: [
      'nostroke noext | M r b A wd2 y1 cd4 cd4 L hc y4 A wd2 y1 0 -5400000 A wd2 y1 cd4 -5400000 L hc y1 A wd2 y1 cd2 cd4 Z',
      'none | M r b A wd2 y1 cd4 cd4 L hc y4 A wd2 y1 0 -5400000 A wd2 y1 cd4 -5400000 L hc y1 A wd2 y1 cd2 cd4',
    ],
  },
  rightBrace: {
    av: 'adj1 8333 adj2 50000',
    gd: `a2 pin 0 adj2 100000
q1 +- 100000 0 a2
q2 min q1 a2
q3 */ q2 1 2
maxAdj1 */ q3 h ss
a1 pin 0 adj1 maxAdj1
y1 */ ss a1 100000
y3 */ h a2 100000
y2 +- y3 0 y1
y4 +- b 0 y1
dx1 cos wd2 2700000
dy1 sin y1 2700000
ir +- l dx1 0
it +- y1 0 dy1
ib +- b dy1 y1`,
    rect: 'l it ir ib',
    path: [
      'nostroke noext | M l t A wd2 y1 3cd4 cd4 L hc y2 A wd2 y1 cd2 -5400000 A wd2 y1 3cd4 -5400000 L hc y4 A wd2 y1 0 cd4 Z',
      'none | M l t A wd2 y1 3cd4 cd4 L hc y2 A wd2 y1 cd2 -5400000 A wd2 y1 3cd4 -5400000 L hc y4 A wd2 y1 0 cd4',
    ],
  },

  // --- Block arrows ---------------------------------------------------------------------------
  rightArrow: {
    av: 'adj1 50000 adj2 50000',
    gd: `maxAdj2 */ 100000 w ss
a1 pin 0 adj1 100000
a2 pin 0 adj2 maxAdj2
dx1 */ ss a2 100000
x1 +- r 0 dx1
dy1 */ h a1 200000
y1 +- vc 0 dy1
y2 +- vc dy1 0
dx2 */ y1 dx1 hd2
x2 +- x1 dx2 0`,
    rect: 'l y1 x2 y2',
    path: 'M l y1 L x1 y1 L x1 t L r vc L x1 b L x1 y2 L l y2 Z',
  },
  leftArrow: {
    av: 'adj1 50000 adj2 50000',
    gd: `maxAdj2 */ 100000 w ss
a1 pin 0 adj1 100000
a2 pin 0 adj2 maxAdj2
dx2 */ ss a2 100000
x2 +- l dx2 0
dy1 */ h a1 200000
y1 +- vc 0 dy1
y2 +- vc dy1 0
dx1 */ y1 dx2 hd2
x1 +- x2 0 dx1`,
    rect: 'x1 y1 r y2',
    path: 'M l vc L x2 t L x2 y1 L r y1 L r y2 L x2 y2 L x2 b Z',
  },
  upArrow: {
    av: 'adj1 50000 adj2 50000',
    gd: `maxAdj2 */ 100000 h ss
a1 pin 0 adj1 100000
a2 pin 0 adj2 maxAdj2
dy2 */ ss a2 100000
y2 +- t dy2 0
dx1 */ w a1 200000
x1 +- hc 0 dx1
x2 +- hc dx1 0
dy1 */ x1 dy2 wd2
y1 +- y2 0 dy1`,
    rect: 'x1 y1 x2 b',
    path: 'M l y2 L hc t L r y2 L x2 y2 L x2 b L x1 b L x1 y2 Z',
  },
  downArrow: {
    av: 'adj1 50000 adj2 50000',
    gd: `maxAdj2 */ 100000 h ss
a1 pin 0 adj1 100000
a2 pin 0 adj2 maxAdj2
dy1 */ ss a2 100000
y1 +- b 0 dy1
dx1 */ w a1 200000
x1 +- hc 0 dx1
x2 +- hc dx1 0
dy2 */ x1 dy1 wd2
y2 +- y1 dy2 0`,
    rect: 'x1 t x2 y2',
    path: 'M l y1 L x1 y1 L x1 t L x2 t L x2 y1 L r y1 L hc b Z',
  },
  leftRightArrow: {
    av: 'adj1 50000 adj2 50000',
    gd: `maxAdj2 */ 50000 w ss
a1 pin 0 adj1 100000
a2 pin 0 adj2 maxAdj2
x2 */ ss a2 100000
x3 +- r 0 x2
dy */ h a1 200000
y1 +- vc 0 dy
y2 +- vc dy 0
dx1 */ y1 x2 hd2
x1 +- x2 0 dx1
x4 +- x3 dx1 0`,
    rect: 'x1 y1 x4 y2',
    path: 'M l vc L x2 t L x2 y1 L x3 y1 L x3 t L r vc L x3 b L x3 y2 L x2 y2 L x2 b Z',
  },
  upDownArrow: {
    av: 'adj1 50000 adj2 50000',
    gd: `maxAdj2 */ 50000 h ss
a1 pin 0 adj1 100000
a2 pin 0 adj2 maxAdj2
y2 */ ss a2 100000
y3 +- b 0 y2
dx1 */ w a1 200000
x1 +- hc 0 dx1
x2 +- hc dx1 0
dy1 */ x1 y2 wd2
y1 +- y2 0 dy1
y4 +- y3 dy1 0`,
    rect: 'x1 y1 x2 y4',
    path: 'M l y2 L hc t L r y2 L x2 y2 L x2 y3 L r y3 L hc b L l y3 L x1 y3 L x1 y2 Z',
  },
  quadArrow: {
    av: 'adj1 22500 adj2 22500 adj3 22500',
    gd: `a2 pin 0 adj2 50000
maxAdj1 */ a2 2 1
a1 pin 0 adj1 maxAdj1
q1 +- 100000 0 maxAdj1
maxAdj3 */ q1 1 2
a3 pin 0 adj3 maxAdj3
x1 */ ss a3 100000
dx2 */ ss a2 100000
x2 +- hc 0 dx2
x5 +- hc dx2 0
dx3 */ ss a1 200000
x3 +- hc 0 dx3
x4 +- hc dx3 0
x6 +- r 0 x1
y2 +- vc 0 dx2
y5 +- vc dx2 0
y3 +- vc 0 dx3
y4 +- vc dx3 0
y6 +- b 0 x1
il */ dx3 x1 dx2
ir +- r 0 il`,
    rect: 'il y3 ir y4',
    path:
      'M l vc L x1 y2 L x1 y3 L x3 y3 L x3 x1 L x2 x1 L hc t L x5 x1 L x4 x1 L x4 y3 L x6 y3 L x6 y2 L r vc ' +
      'L x6 y5 L x6 y4 L x4 y4 L x4 y6 L x5 y6 L hc b L x2 y6 L x3 y6 L x3 y4 L x1 y4 L x1 y5 Z',
  },
  leftRightUpArrow: {
    av: 'adj1 25000 adj2 25000 adj3 25000',
    gd: `a2 pin 0 adj2 50000
maxAdj1 */ a2 2 1
a1 pin 0 adj1 maxAdj1
q1 +- 100000 0 maxAdj1
maxAdj3 */ q1 1 2
a3 pin 0 adj3 maxAdj3
x1 */ ss a3 100000
dx2 */ ss a2 100000
x2 +- hc 0 dx2
x5 +- hc dx2 0
dx3 */ ss a1 200000
x3 +- hc 0 dx3
x4 +- hc dx3 0
x6 +- r 0 x1
dy2 */ ss a2 50000
y2 +- b 0 dy2
y4 +- b 0 dx2
y3 +- y4 0 dx3
y5 +- y4 dx3 0
il */ dx3 x1 dx2
ir +- r 0 il`,
    rect: 'il y3 ir y5',
    path:
      'M l y4 L x1 y2 L x1 y3 L x3 y3 L x3 x1 L x2 x1 L hc t L x5 x1 L x4 x1 L x4 y3 L x6 y3 L x6 y2 L r y4 ' +
      'L x6 b L x6 y5 L x1 y5 L x1 b Z',
  },
  notchedRightArrow: {
    av: 'adj1 50000 adj2 50000',
    gd: `maxAdj2 */ 100000 w ss
a1 pin 0 adj1 100000
a2 pin 0 adj2 maxAdj2
dx2 */ ss a2 100000
x2 +- r 0 dx2
dy1 */ h a1 200000
y1 +- vc 0 dy1
y2 +- vc dy1 0
x1 */ dy1 dx2 hd2
x3 +- r 0 x1`,
    rect: 'x1 y1 x3 y2',
    path: 'M l y1 L x2 y1 L x2 t L r vc L x2 b L x2 y2 L l y2 L x1 vc Z',
  },
  stripedRightArrow: {
    av: 'adj1 50000 adj2 50000',
    gd: `maxAdj2 */ 84375 w ss
a1 pin 0 adj1 100000
a2 pin 0 adj2 maxAdj2
x4 */ ss 5 32
dx5 */ ss a2 100000
x5 +- r 0 dx5
dy1 */ h a1 200000
y1 +- vc 0 dy1
y2 +- vc dy1 0
dx6 */ dy1 dx5 hd2
x6 +- r 0 dx6`,
    rect: 'x4 y1 x6 y2',
    path:
      'M l y1 L ssd32 y1 L ssd32 y2 L l y2 Z M ssd16 y1 L ssd8 y1 L ssd8 y2 L ssd16 y2 Z ' +
      'M x4 y1 L x5 y1 L x5 t L r vc L x5 b L x5 y2 L x4 y2 Z',
  },
  chevron: {
    av: 'adj 50000',
    gd: `maxAdj */ 100000 w ss
a pin 0 adj maxAdj
x1 */ ss a 100000
x2 +- r 0 x1
x3 */ x2 1 2
dx +- x2 0 x1
il ?: dx x1 l
ir ?: dx x2 r`,
    rect: 'il t ir b',
    path: 'M l t L x2 t L r vc L x2 b L l b L x1 vc Z',
  },
  homePlate: {
    av: 'adj 50000',
    gd: `maxAdj */ 100000 w ss
a pin 0 adj maxAdj
dx1 */ ss a 100000
x1 +- r 0 dx1
ir +/ x1 r 2
x2 */ x1 1 2`,
    rect: 'l t ir b',
    path: 'M l t L x1 t L r vc L x1 b L l b Z',
  },
  bentArrow: {
    av: 'adj1 25000 adj2 25000 adj3 25000 adj4 43750',
    gd: `a2 pin 0 adj2 50000
maxAdj1 */ a2 2 1
a1 pin 0 adj1 maxAdj1
a3 pin 0 adj3 50000
th */ ss a1 100000
aw2 */ ss a2 100000
th2 */ th 1 2
dh2 +- aw2 0 th2
ah */ ss a3 100000
bw +- r 0 ah
bh +- b 0 dh2
bs min bw bh
maxAdj4 */ 100000 bs ss
a4 pin 0 adj4 maxAdj4
bd */ ss a4 100000
bd3 +- bd 0 th
bd2 max bd3 0
x3 +- th bd2 0
x4 +- r 0 ah
y3 +- dh2 th 0
y4 +- y3 dh2 0
y5 +- dh2 bd 0
y6 +- y3 bd2 0`,
    path: 'M l b L l y5 A bd bd cd2 cd4 L x4 dh2 L x4 t L r aw2 L x4 y4 L x4 y3 L x3 y3 A bd2 bd2 3cd4 -5400000 L th b Z',
  },
  uturnArrow: {
    av: 'adj1 25000 adj2 25000 adj3 25000 adj4 43750 adj5 75000',
    gd: `a2 pin 0 adj2 25000
maxAdj1 */ a2 2 1
a1 pin 0 adj1 maxAdj1
q2 */ a1 ss h
q3 +- 100000 0 q2
maxAdj3 */ q3 h ss
a3 pin 0 adj3 maxAdj3
q1 +- a3 a1 0
minAdj5 */ q1 ss h
a5 pin minAdj5 adj5 100000
th */ ss a1 100000
aw2 */ ss a2 100000
th2 */ th 1 2
dh2 +- aw2 0 th2
y5 */ h a5 100000
ah */ ss a3 100000
y4 +- y5 0 ah
x9 +- r 0 dh2
bw */ x9 1 2
bs min bw y4
maxAdj4 */ bs 100000 ss
a4 pin 0 adj4 maxAdj4
bd */ ss a4 100000
bd3 +- bd 0 th
bd2 max bd3 0
x3 +- th bd2 0
x8 +- r 0 aw2
x6 +- x8 0 aw2
x7 +- x6 dh2 0
x4 +- x9 0 bd
x5 +- x7 0 bd2
cx +/ th x7 2`,
    rect: 'l t x9 y4',
    path:
      'M l b L l bd A bd bd cd2 cd4 L x4 t A bd bd 3cd4 cd4 L x9 y4 L r y4 L x8 y5 L x6 y4 L x7 y4 L x7 x3 ' +
      'A bd2 bd2 0 -5400000 L x3 th A bd2 bd2 3cd4 -5400000 L th b Z',
  },
  curvedRightArrow: {
    av: 'adj1 25000 adj2 50000 adj3 25000',
    gd: `maxAdj2 */ 50000 h ss
a2 pin 0 adj2 maxAdj2
a1 pin 0 adj1 a2
th */ ss a1 100000
aw */ ss a2 100000
q1 +/ th aw 4
hR +- hd2 0 q1
q7 */ hR 2 1
q8 */ q7 q7 1
q9 */ th th 1
q10 +- q8 0 q9
q11 sqrt q10
idx */ q11 w q7
maxAdj3 */ 100000 idx ss
a3 pin 0 adj3 maxAdj3
ah */ ss a3 100000
y3 +- hR th 0
q2 */ w w 1
q3 */ ah ah 1
q4 +- q2 0 q3
q5 sqrt q4
dy */ q5 hR w
y5 +- hR dy 0
y7 +- y3 dy 0
q6 +- aw 0 th
dh */ q6 1 2
y4 +- y5 0 dh
y8 +- y7 dh 0
aw2 */ aw 1 2
y6 +- b 0 aw2
x1 +- r 0 ah
swAng at2 ah dy
stAng +- cd2 0 swAng
mswAng +- 0 0 swAng
ix +- r 0 idx
iy +/ hR y3 2
q12 */ th 1 2
dang2 at2 idx q12
swAng2 +- dang2 0 cd4
swAng3 +- cd4 dang2 0
stAng3 +- cd2 0 dang2`,
    path: [
      'nostroke noext | M l hR A w hR cd2 mswAng L x1 y4 L r y6 L x1 y8 L x1 y7 A w hR stAng swAng Z',
      'darkenLess nostroke noext | M r th A w hR 3cd4 swAng2 A w hR stAng3 swAng3 Z',
      'none noext | M l hR A w hR cd2 mswAng L x1 y4 L r y6 L x1 y8 L x1 y7 A w hR stAng swAng L l hR A w hR cd2 cd4 L r th A w hR 3cd4 swAng2',
    ],
  },
  curvedLeftArrow: {
    av: 'adj1 25000 adj2 50000 adj3 25000',
    gd: `maxAdj2 */ 50000 h ss
a2 pin 0 adj2 maxAdj2
a1 pin 0 adj1 a2
th */ ss a1 100000
aw */ ss a2 100000
q1 +/ th aw 4
hR +- hd2 0 q1
q7 */ hR 2 1
q8 */ q7 q7 1
q9 */ th th 1
q10 +- q8 0 q9
q11 sqrt q10
idx */ q11 w q7
maxAdj3 */ 100000 idx ss
a3 pin 0 adj3 maxAdj3
ah */ ss a3 100000
y3 +- hR th 0
q2 */ w w 1
q3 */ ah ah 1
q4 +- q2 0 q3
q5 sqrt q4
dy */ q5 hR w
y5 +- hR dy 0
y7 +- y3 dy 0
q6 +- aw 0 th
dh */ q6 1 2
y4 +- y5 0 dh
y8 +- y7 dh 0
aw2 */ aw 1 2
y6 +- b 0 aw2
x1 +- l ah 0
swAng at2 ah dy
mswAng +- 0 0 swAng
ix +- l idx 0
iy +/ hR y3 2
q12 */ th 1 2
dang2 at2 idx q12
swAng2 +- dang2 0 swAng
swAng3 +- swAng dang2 0
stAng3 +- 0 0 dang2`,
    path: [
      'nostroke noext | M l y6 L x1 y4 L x1 y5 A w hR swAng swAng2 A w hR stAng3 swAng3 L x1 y8 Z',
      'darkenLess nostroke noext | M r y3 A w hR 0 -5400000 L l t A w hR 3cd4 cd4 Z',
      'none noext | M r y3 A w hR 0 -5400000 L l t A w hR 3cd4 cd4 L r y3 A w hR 0 swAng L x1 y8 L l y6 L x1 y4 L x1 y5 A w hR swAng swAng2',
    ],
  },
  curvedUpArrow: {
    av: 'adj1 25000 adj2 50000 adj3 25000',
    gd: `maxAdj2 */ 50000 w ss
a2 pin 0 adj2 maxAdj2
a1 pin 0 adj1 100000
th */ ss a1 100000
aw */ ss a2 100000
q1 +/ th aw 4
wR +- wd2 0 q1
q7 */ wR 2 1
q8 */ q7 q7 1
q9 */ th th 1
q10 +- q8 0 q9
q11 sqrt q10
idy */ q11 h q7
maxAdj3 */ 100000 idy ss
a3 pin 0 adj3 maxAdj3
ah */ ss a3 100000
x3 +- wR th 0
q2 */ h h 1
q3 */ ah ah 1
q4 +- q2 0 q3
q5 sqrt q4
dx */ q5 wR h
x5 +- wR dx 0
x7 +- x3 dx 0
q6 +- aw 0 th
dh */ q6 1 2
x4 +- x5 0 dh
x8 +- x7 dh 0
aw2 */ aw 1 2
x6 +- r 0 aw2
y1 +- t ah 0
swAng at2 ah dx
mswAng +- 0 0 swAng
iy +- t idy 0
ix +/ wR x3 2
q12 */ th 1 2
dang2 at2 idy q12
swAng2 +- dang2 0 swAng
mswAng2 +- 0 0 swAng2
stAng3 +- cd4 0 swAng
swAng3 +- swAng dang2 0
stAng2 +- cd4 0 dang2`,
    path: [
      'nostroke noext | M x6 t L x8 y1 L x7 y1 A wR h stAng3 swAng3 A wR h stAng2 swAng2 L x4 y1 Z',
      'darkenLess nostroke noext | M wR b A wR h cd4 cd4 L th t A wR h cd2 -5400000 Z',
      'none noext | M ix iy A wR h stAng2 swAng2 L x4 y1 L x6 t L x8 y1 L x7 y1 A wR h stAng3 swAng3 L wR b A wR h cd4 cd4 L th t A wR h cd2 -5400000',
    ],
  },
  curvedDownArrow: {
    av: 'adj1 25000 adj2 50000 adj3 25000',
    gd: `maxAdj2 */ 50000 w ss
a2 pin 0 adj2 maxAdj2
a1 pin 0 adj1 100000
th */ ss a1 100000
aw */ ss a2 100000
q1 +/ th aw 4
wR +- wd2 0 q1
q7 */ wR 2 1
q8 */ q7 q7 1
q9 */ th th 1
q10 +- q8 0 q9
q11 sqrt q10
idy */ q11 h q7
maxAdj3 */ 100000 idy ss
a3 pin 0 adj3 maxAdj3
ah */ ss a3 100000
x3 +- wR th 0
q2 */ h h 1
q3 */ ah ah 1
q4 +- q2 0 q3
q5 sqrt q4
dx */ q5 wR h
x5 +- wR dx 0
x7 +- x3 dx 0
q6 +- aw 0 th
dh */ q6 1 2
x4 +- x5 0 dh
x8 +- x7 dh 0
aw2 */ aw 1 2
x6 +- r 0 aw2
y1 +- b 0 ah
swAng at2 ah dx
mswAng +- 0 0 swAng
iy +- b 0 idy
ix +/ wR x3 2
q12 */ th 1 2
dang2 at2 idy q12
stAng +- 3cd4 swAng 0
stAng2 +- 3cd4 0 dang2
swAng2 +- dang2 0 cd4
swAng3 +- cd4 dang2 0`,
    path: [
      'nostroke noext | M x6 b L x4 y1 L x5 y1 A wR h stAng mswAng L x3 t A wR h 3cd4 swAng L x8 y1 Z',
      'darkenLess nostroke noext | M ix iy A wR h stAng2 swAng2 L l b A wR h cd2 swAng3 Z',
      'none noext | M ix iy A wR h stAng2 swAng2 L l b A wR h cd2 cd4 L x3 t A wR h 3cd4 swAng L x8 y1 L x6 b L x4 y1 L x5 y1 A wR h stAng mswAng',
    ],
  },
  circularArrow: {
    av: 'adj1 12500 adj2 1142319 adj3 20457681 adj4 10800000 adj5 12500',
    gd: `a5 pin 0 adj5 25000
maxAdj1 */ a5 2 1
a1 pin 0 adj1 maxAdj1
enAng pin 1 adj3 21599999
stAng pin 0 adj4 21599999
th */ ss a1 100000
thh */ ss a5 100000
th2 */ th 1 2
rw1 +- wd2 th2 thh
rh1 +- hd2 th2 thh
rw2 +- rw1 0 th
rh2 +- rh1 0 th
rw3 +- rw2 th2 0
rh3 +- rh2 th2 0
wtH sin rw3 enAng
htH cos rh3 enAng
dxH cat2 rw3 htH wtH
dyH sat2 rh3 htH wtH
xH +- hc dxH 0
yH +- vc dyH 0
rI min rw2 rh2
u1 */ dxH dxH 1
u2 */ dyH dyH 1
u3 */ rI rI 1
u4 +- u1 0 u3
u5 +- u2 0 u3
u6 */ u4 u5 u1
u7 */ u6 1 u2
u8 +- 1 0 u7
u9 sqrt u8
u10 */ u4 1 dxH
u11 */ u10 1 dyH
u12 +/ 1 u9 u11
u13 at2 1 u12
u14 +- u13 21600000 0
u15 ?: u13 u13 u14
u16 +- u15 0 enAng
u17 +- u16 21600000 0
u18 ?: u16 u16 u17
u19 +- u18 0 cd2
u20 +- u18 0 21600000
u21 ?: u19 u20 u18
maxAng abs u21
aAng pin 0 adj2 maxAng
ptAng +- enAng aAng 0
wtA sin rw3 ptAng
htA cos rh3 ptAng
dxA cat2 rw3 htA wtA
dyA sat2 rh3 htA wtA
xA +- hc dxA 0
yA +- vc dyA 0
wtE sin rw1 stAng
htE cos rh1 stAng
dxE cat2 rw1 htE wtE
dyE sat2 rh1 htE wtE
xE +- hc dxE 0
yE +- vc dyE 0
dxG cos thh ptAng
dyG sin thh ptAng
xG +- xH dxG 0
yG +- yH dyG 0
dxB cos thh ptAng
dyB sin thh ptAng
xB +- xH 0 dxB
yB +- yH 0 dyB
sx1 +- xB 0 hc
sy1 +- yB 0 vc
sx2 +- xG 0 hc
sy2 +- yG 0 vc
rO min rw1 rh1
x1O */ sx1 rO rw1
y1O */ sy1 rO rh1
x2O */ sx2 rO rw1
y2O */ sy2 rO rh1
dxO +- x2O 0 x1O
dyO +- y2O 0 y1O
dO mod dxO dyO 0
q1 */ x1O y2O 1
q2 */ x2O y1O 1
DO +- q1 0 q2
q3 */ rO rO 1
q4 */ dO dO 1
q5 */ q3 q4 1
q6 */ DO DO 1
q7 +- q5 0 q6
q8 max q7 0
sdelO sqrt q8
ndyO */ dyO -1 1
sdyO ?: ndyO -1 1
q9 */ sdyO dxO 1
q10 */ q9 sdelO 1
q11 */ DO dyO 1
dxF1 +/ q11 q10 q4
q12 +- q11 0 q10
dxF2 */ q12 1 q4
adyO abs dyO
q13 */ adyO sdelO 1
q14 */ DO dxO -1
dyF1 +/ q14 q13 q4
q15 +- q14 0 q13
dyF2 */ q15 1 q4
q16 +- x2O 0 dxF1
q17 +- x2O 0 dxF2
q18 +- y2O 0 dyF1
q19 +- y2O 0 dyF2
q20 mod q16 q18 0
q21 mod q17 q19 0
q22 +- q21 0 q20
dxF ?: q22 dxF1 dxF2
dyF ?: q22 dyF1 dyF2
sdxF */ dxF rw1 rO
sdyF */ dyF rh1 rO
xF +- hc sdxF 0
yF +- vc sdyF 0
x1I */ sx1 rI rw2
y1I */ sy1 rI rh2
x2I */ sx2 rI rw2
y2I */ sy2 rI rh2
dxI +- x2I 0 x1I
dyI +- y2I 0 y1I
dI mod dxI dyI 0
v1 */ x1I y2I 1
v2 */ x2I y1I 1
DI +- v1 0 v2
v3 */ rI rI 1
v4 */ dI dI 1
v5 */ v3 v4 1
v6 */ DI DI 1
v7 +- v5 0 v6
v8 max v7 0
sdelI sqrt v8
v9 */ sdyO dxI 1
v10 */ v9 sdelI 1
v11 */ DI dyI 1
dxC1 +/ v11 v10 v4
v12 +- v11 0 v10
dxC2 */ v12 1 v4
adyI abs dyI
v13 */ adyI sdelI 1
v14 */ DI dxI -1
dyC1 +/ v14 v13 v4
v15 +- v14 0 v13
dyC2 */ v15 1 v4
v16 +- x1I 0 dxC1
v17 +- x1I 0 dxC2
v18 +- y1I 0 dyC1
v19 +- y1I 0 dyC2
v20 mod v16 v18 0
v21 mod v17 v19 0
v22 +- v21 0 v20
dxC ?: v22 dxC1 dxC2
dyC ?: v22 dyC1 dyC2
sdxC */ dxC rw2 rI
sdyC */ dyC rh2 rI
xC +- hc sdxC 0
yC +- vc sdyC 0
ist0 at2 sdxC sdyC
ist1 +- ist0 21600000 0
istAng ?: ist0 ist0 ist1
isw1 +- stAng 0 istAng
isw2 +- isw1 0 21600000
iswAng ?: isw1 isw2 isw1
p1 +- xF 0 xC
p2 +- yF 0 yC
p3 mod p1 p2 0
p4 */ p3 1 2
p5 +- p4 0 thh
xGp ?: p5 xF xG
yGp ?: p5 yF yG
xBp ?: p5 xC xB
yBp ?: p5 yC yB
en0 at2 sdxF sdyF
en1 +- en0 21600000 0
en2 ?: en0 en0 en1
sw0 +- en2 0 stAng
sw1 +- sw0 21600000 0
swAng ?: sw0 sw0 sw1
idx cos rw1 2700000
idy sin rh1 2700000
il +- hc 0 idx
ir +- hc idx 0
it +- vc 0 idy
ib +- vc idy 0`,
    rect: 'il it ir ib',
    path: 'M xE yE A rw1 rh1 stAng swAng L xGp yGp L xA yA L xBp yBp L xC yC A rw2 rh2 istAng iswAng Z',
  },

  // --- Callouts -------------------------------------------------------------------------------
  rightArrowCallout: {
    av: 'adj1 25000 adj2 25000 adj3 25000 adj4 64977',
    gd: `maxAdj2 */ 50000 h ss
a2 pin 0 adj2 maxAdj2
maxAdj1 */ a2 2 1
a1 pin 0 adj1 maxAdj1
maxAdj3 */ 100000 w ss
a3 pin 0 adj3 maxAdj3
q2 */ a3 ss w
maxAdj4 +- 100000 0 q2
a4 pin 0 adj4 maxAdj4
dy1 */ ss a2 100000
dy2 */ ss a1 200000
y1 +- vc 0 dy1
y2 +- vc 0 dy2
y3 +- vc dy2 0
y4 +- vc dy1 0
dx3 */ ss a3 100000
x3 +- r 0 dx3
x2 */ w a4 100000
x1 */ x2 1 2`,
    rect: 'l t x2 b',
    path: 'M l t L x2 t L x2 y2 L x3 y2 L x3 y1 L r vc L x3 y4 L x3 y3 L x2 y3 L x2 b L l b Z',
  },
  leftArrowCallout: {
    av: 'adj1 25000 adj2 25000 adj3 25000 adj4 64977',
    gd: `maxAdj2 */ 50000 h ss
a2 pin 0 adj2 maxAdj2
maxAdj1 */ a2 2 1
a1 pin 0 adj1 maxAdj1
maxAdj3 */ 100000 w ss
a3 pin 0 adj3 maxAdj3
q2 */ a3 ss w
maxAdj4 +- 100000 0 q2
a4 pin 0 adj4 maxAdj4
dy1 */ ss a2 100000
dy2 */ ss a1 200000
y1 +- vc 0 dy1
y2 +- vc 0 dy2
y3 +- vc dy2 0
y4 +- vc dy1 0
x1 */ ss a3 100000
dx2 */ w a4 100000
x2 +- r 0 dx2
x3 +/ x2 r 2`,
    rect: 'x2 t r b',
    path: 'M l vc L x1 y1 L x1 y2 L x2 y2 L x2 t L r t L r b L x2 b L x2 y3 L x1 y3 L x1 y4 Z',
  },
  upArrowCallout: {
    av: 'adj1 25000 adj2 25000 adj3 25000 adj4 64977',
    gd: `maxAdj2 */ 50000 w ss
a2 pin 0 adj2 maxAdj2
maxAdj1 */ a2 2 1
a1 pin 0 adj1 maxAdj1
maxAdj3 */ 100000 h ss
a3 pin 0 adj3 maxAdj3
q2 */ a3 ss h
maxAdj4 +- 100000 0 q2
a4 pin 0 adj4 maxAdj4
dx1 */ ss a2 100000
dx2 */ ss a1 200000
x1 +- hc 0 dx1
x2 +- hc 0 dx2
x3 +- hc dx2 0
x4 +- hc dx1 0
y1 */ ss a3 100000
dy2 */ h a4 100000
y2 +- b 0 dy2
y3 +/ y2 b 2`,
    rect: 'l y2 r b',
    path: 'M l y2 L x2 y2 L x2 y1 L x1 y1 L hc t L x4 y1 L x3 y1 L x3 y2 L r y2 L r b L l b Z',
  },
  downArrowCallout: {
    av: 'adj1 25000 adj2 25000 adj3 25000 adj4 64977',
    gd: `maxAdj2 */ 50000 w ss
a2 pin 0 adj2 maxAdj2
maxAdj1 */ a2 2 1
a1 pin 0 adj1 maxAdj1
maxAdj3 */ 100000 h ss
a3 pin 0 adj3 maxAdj3
q2 */ a3 ss h
maxAdj4 +- 100000 0 q2
a4 pin 0 adj4 maxAdj4
dx1 */ ss a2 100000
dx2 */ ss a1 200000
x1 +- hc 0 dx1
x2 +- hc 0 dx2
x3 +- hc dx2 0
x4 +- hc dx1 0
dy3 */ ss a3 100000
y3 +- b 0 dy3
y2 */ h a4 100000
y1 */ y2 1 2`,
    rect: 'l t r y2',
    path: 'M l t L r t L r y2 L x3 y2 L x3 y3 L x4 y3 L hc b L x1 y3 L x2 y3 L x2 y2 L l y2 Z',
  },
  leftRightArrowCallout: {
    av: 'adj1 25000 adj2 25000 adj3 25000 adj4 48123',
    gd: `maxAdj2 */ 50000 h ss
a2 pin 0 adj2 maxAdj2
maxAdj1 */ a2 2 1
a1 pin 0 adj1 maxAdj1
maxAdj3 */ 50000 w ss
a3 pin 0 adj3 maxAdj3
q2 */ a3 ss wd2
maxAdj4 +- 100000 0 q2
a4 pin 0 adj4 maxAdj4
dy1 */ ss a2 100000
dy2 */ ss a1 200000
y1 +- vc 0 dy1
y2 +- vc 0 dy2
y3 +- vc dy2 0
y4 +- vc dy1 0
x1 */ ss a3 100000
x4 +- r 0 x1
dx2 */ w a4 200000
x2 +- hc 0 dx2
x3 +- hc dx2 0`,
    rect: 'x2 t x3 b',
    path:
      'M l vc L x1 y1 L x1 y2 L x2 y2 L x2 t L x3 t L x3 y2 L x4 y2 L x4 y1 L r vc L x4 y4 L x4 y3 ' +
      'L x3 y3 L x3 b L x2 b L x2 y3 L x1 y3 L x1 y4 Z',
  },
  wedgeRectCallout: {
    av: 'adj1 -20833 adj2 62500',
    gd: WEDGE_GD,
    path: 'M l t L x1 t L xt yt L x2 t L r t L r y1 L xr yr L r y2 L r b L x2 b L xb yb L x1 b L l b L l y2 L xl yl L l y1 Z',
  },
  wedgeRoundRectCallout: {
    av: 'adj1 -20833 adj2 62500 adj3 16667',
    gd: `${WEDGE_GD}
u1 */ ss adj3 100000
u2 +- r 0 u1
v2 +- b 0 u1
il */ u1 29289 100000
ir +- r 0 il
ib +- b 0 il`,
    rect: 'il il ir ib',
    path:
      'M l u1 A u1 u1 cd2 cd4 L x1 t L xt yt L x2 t L u2 t A u1 u1 3cd4 cd4 L r y1 L xr yr L r y2 L r v2 ' +
      'A u1 u1 0 cd4 L x2 b L xb yb L x1 b L u1 b A u1 u1 cd4 cd4 L l y2 L xl yl L l y1 Z',
  },
  wedgeEllipseCallout: {
    av: 'adj1 -20833 adj2 62500',
    gd: `dxPos */ w adj1 100000
dyPos */ h adj2 100000
xPos +- hc dxPos 0
yPos +- vc dyPos 0
sdx */ dxPos h 1
sdy */ dyPos w 1
pang at2 sdx sdy
stAng +- pang 660000 0
enAng +- pang 0 660000
dx1 cos wd2 stAng
dy1 sin hd2 stAng
x1 +- hc dx1 0
y1 +- vc dy1 0
dx2 cos wd2 enAng
dy2 sin hd2 enAng
x2 +- hc dx2 0
y2 +- vc dy2 0
stAng1 at2 dx1 dy1
enAng1 at2 dx2 dy2
swAng1 +- enAng1 0 stAng1
swAng2 +- swAng1 21600000 0
swAng ?: swAng1 swAng1 swAng2
${ELLIPSE_RECT_GD}`,
    rect: 'il it ir ib',
    path: 'M xPos yPos L x1 y1 A wd2 hd2 stAng1 swAng Z',
  },
  cloudCallout: {
    av: 'adj1 -20833 adj2 62500',
    gd: `dxPos */ w adj1 100000
dyPos */ h adj2 100000
xPos +- hc dxPos 0
yPos +- vc dyPos 0
ht cat2 hd2 dxPos dyPos
wt sat2 wd2 dxPos dyPos
g2 cat2 wd2 ht wt
g3 sat2 hd2 ht wt
g4 +- hc g2 0
g5 +- vc g3 0
g6 +- g4 0 xPos
g7 +- g5 0 yPos
g8 mod g6 g7 0
g9 */ ss 6600 21600
g10 +- g8 0 g9
g11 */ g10 1 3
g12 */ ss 1800 21600
g13 +- g11 g12 0
g14 */ g13 g6 g8
g15 */ g13 g7 g8
g16 +- g14 xPos 0
g17 +- g15 yPos 0
g18 */ ss 4800 21600
g19 */ g11 2 1
g20 +- g18 g19 0
g21 */ g20 g6 g8
g22 */ g20 g7 g8
g23 +- g21 xPos 0
g24 +- g22 yPos 0
g25 */ ss 1200 21600
g26 */ ss 600 21600
x23 +- xPos g26 0
x24 +- g16 g25 0
x25 +- g23 g12 0
${CLOUD_RECT_GD}`,
    rect: 'il it ir ib',
    path: [
      CLOUD_OUTLINE,
      'M x23 yPos A g26 g26 0 21600000 Z',
      'M x24 g17 A g25 g25 0 21600000 Z',
      'M x25 g24 A g12 g12 0 21600000 Z',
      CLOUD_DETAIL,
    ],
  },
  callout1: callout(1, false, false),
  callout2: callout(2, false, false),
  callout3: callout(3, false, false),
  borderCallout1: callout(1, true, false),
  borderCallout2: callout(2, true, false),
  borderCallout3: callout(3, true, false),
  accentCallout1: callout(1, false, true),
  accentCallout2: callout(2, false, true),
  accentCallout3: callout(3, false, true),
  accentBorderCallout1: callout(1, true, true),
  accentBorderCallout2: callout(2, true, true),
  accentBorderCallout3: callout(3, true, true),

  // --- Ribbons, waves and scrolls -------------------------------------------------------------
  ribbon: {
    av: 'adj1 16667 adj2 50000',
    gd: `a1 pin 0 adj1 33333
a2 pin 25000 adj2 75000
x10 +- r 0 wd8
dx2 */ w a2 200000
x2 +- hc 0 dx2
x9 +- hc dx2 0
x3 +- x2 wd32 0
x8 +- x9 0 wd32
x5 +- x2 wd8 0
x6 +- x9 0 wd8
x4 +- x5 0 wd32
x7 +- x6 wd32 0
y1 */ h a1 200000
y2 */ h a1 100000
y4 +- b 0 y2
y3 */ y4 1 2
hR */ h a1 400000
y5 +- b 0 hR
y6 +- y2 0 hR`,
    rect: 'x2 y2 x9 b',
    path: [
      'nostroke noext | M l t L x4 t A wd32 hR 3cd4 cd2 L x3 y1 A wd32 hR 3cd4 -10800000 L x8 y2 A wd32 hR cd4 -10800000 ' +
        'L x7 y1 A wd32 hR cd4 cd2 L r t L x10 y3 L r y4 L x9 y4 L x9 y5 A wd32 hR 0 cd4 L x3 b A wd32 hR cd4 cd4 ' +
        'L x2 y4 L l y4 L wd8 y3 Z',
      'darkenLess nostroke noext | M x5 hR A wd32 hR 0 cd4 L x3 y1 A wd32 hR 3cd4 -10800000 L x5 y2 Z ' +
        'M x6 hR A wd32 hR cd2 -5400000 L x8 y1 A wd32 hR 3cd4 cd2 L x6 y2 Z',
      'none noext | M l t L x4 t A wd32 hR 3cd4 cd2 L x3 y1 A wd32 hR 3cd4 -10800000 L x8 y2 A wd32 hR cd4 -10800000 ' +
        'L x7 y1 A wd32 hR cd4 cd2 L r t L x10 y3 L r y4 L x9 y4 L x9 y5 A wd32 hR 0 cd4 L x3 b A wd32 hR cd4 cd4 ' +
        'L x2 y4 L l y4 L wd8 y3 Z M x5 hR L x5 y2 M x6 y2 L x6 hR M x2 y4 L x2 y6 M x9 y6 L x9 y4',
    ],
  },
  ribbon2: {
    av: 'adj1 16667 adj2 50000',
    gd: `a1 pin 0 adj1 33333
a2 pin 25000 adj2 75000
x10 +- r 0 wd8
dx2 */ w a2 200000
x2 +- hc 0 dx2
x9 +- hc dx2 0
x3 +- x2 wd32 0
x8 +- x9 0 wd32
x5 +- x2 wd8 0
x6 +- x9 0 wd8
x4 +- x5 0 wd32
x7 +- x6 wd32 0
dy1 */ h a1 200000
y1 +- b 0 dy1
dy2 */ h a1 100000
y2 +- b 0 dy2
y4 +- t dy2 0
y3 +/ y4 b 2
hR */ h a1 400000
y6 +- b 0 hR
y7 +- y1 0 hR`,
    rect: 'x2 t x9 y2',
    path: [
      'nostroke noext | M l b L x4 b A wd32 hR cd4 -10800000 L x3 y1 A wd32 hR cd4 cd2 L x8 y2 A wd32 hR 3cd4 cd2 ' +
        'L x7 y1 A wd32 hR 3cd4 -10800000 L r b L x10 y3 L r y4 L x9 y4 L x9 hR A wd32 hR 0 -5400000 L x3 t ' +
        'A wd32 hR 3cd4 -5400000 L x2 y4 L l y4 L wd8 y3 Z',
      'darkenLess nostroke noext | M x5 y6 A wd32 hR 0 -5400000 L x3 y1 A wd32 hR cd4 cd2 L x5 y2 Z ' +
        'M x6 y6 A wd32 hR cd2 cd4 L x8 y1 A wd32 hR cd4 -10800000 L x6 y2 Z',
      'none noext | M l b L x4 b A wd32 hR cd4 -10800000 L x3 y1 A wd32 hR cd4 cd2 L x8 y2 A wd32 hR 3cd4 cd2 ' +
        'L x7 y1 A wd32 hR 3cd4 -10800000 L r b L x10 y3 L r y4 L x9 y4 L x9 hR A wd32 hR 0 -5400000 L x3 t ' +
        'A wd32 hR 3cd4 -5400000 L x2 y4 L l y4 L wd8 y3 Z M x5 y2 L x5 y6 M x6 y6 L x6 y2 M x2 y7 L x2 y4 M x9 y4 L x9 y7',
    ],
  },
  wave: {
    av: 'adj1 12500 adj2 0',
    gd: `a1 pin 0 adj1 20000
a2 pin -10000 adj2 10000
y1 */ h a1 100000
dy2 */ y1 10 3
y2 +- y1 0 dy2
y3 +- y1 dy2 0
y4 +- b 0 y1
y5 +- y4 0 dy2
y6 +- y4 dy2 0
dx1 */ w a2 100000
of2 */ w a2 50000
x1 abs dx1
dx2 ?: of2 0 of2
x2 +- l 0 dx2
dx5 ?: of2 of2 0
x5 +- r 0 dx5
dx3 +/ dx2 x5 3
x3 +- x2 dx3 0
x4 +/ x3 x5 2
x6 +- l dx5 0
x10 +- r dx2 0
x7 +- x6 dx3 0
x8 +/ x7 x10 2
x9 +- r 0 x1
xAdj +- hc dx1 0
xAdj2 +- hc 0 dx1
il max x2 x6
ir min x5 x10
it */ h a1 50000
ib +- b 0 it`,
    rect: 'il it ir ib',
    path: 'M x2 y1 C x3 y2 x4 y3 x5 y1 L x10 y4 C x8 y6 x7 y5 x6 y4 Z',
  },
  doubleWave: {
    av: 'adj1 6250 adj2 0',
    gd: `a1 pin 0 adj1 12500
a2 pin -10000 adj2 10000
y1 */ h a1 100000
dy2 */ y1 10 3
y2 +- y1 0 dy2
y3 +- y1 dy2 0
y4 +- b 0 y1
y5 +- y4 0 dy2
y6 +- y4 dy2 0
dx1 */ w a2 100000
of2 */ w a2 50000
x1 abs dx1
dx2 ?: of2 0 of2
x2 +- l 0 dx2
dx8 ?: of2 of2 0
x8 +- r 0 dx8
dx3 +/ dx2 x8 6
x3 +- x2 dx3 0
dx4 +/ dx2 x8 3
x4 +- x2 dx4 0
x5 +/ x2 x8 2
x6 +- x5 dx3 0
x7 +/ x6 x8 2
x9 +- l dx8 0
x15 +- r dx2 0
x10 +- x9 dx3 0
x11 +- x9 dx4 0
x12 +/ x9 x15 2
x13 +- x12 dx3 0
x14 +/ x13 x15 2
x16 +- r 0 x1
xAdj +- hc dx1 0
il max x2 x9
ir min x8 x15
it */ h a1 50000
ib +- b 0 it`,
    rect: 'il it ir ib',
    path: 'M x2 y1 C x3 y2 x4 y3 x5 y1 C x6 y2 x7 y3 x8 y1 L x15 y4 C x14 y6 x13 y5 x12 y4 C x11 y6 x10 y5 x9 y4 Z',
  },
  horizontalScroll: {
    av: 'adj 12500',
    gd: `a pin 0 adj 25000
ch */ ss a 100000
ch2 */ ch 1 2
ch4 */ ch 1 4
y3 +- ch ch2 0
y4 +- ch ch 0
y6 +- b 0 ch
y7 +- b 0 ch2
y5 +- y6 0 ch2
x3 +- r 0 ch
x4 +- r 0 ch2`,
    rect: 'ch y4 x4 y6',
    path: [
      'nostroke noext | M r ch2 A ch2 ch2 0 cd4 L x4 ch2 A ch4 ch4 0 cd2 L x3 ch L ch2 ch A ch2 ch2 3cd4 -5400000 ' +
        'L l y7 A ch2 ch2 cd2 -10800000 L ch y6 L x4 y6 A ch2 ch2 cd4 -5400000 Z ' +
        'M ch2 y4 A ch2 ch2 cd4 -5400000 A ch4 ch4 0 -10800000 Z',
      'darkenLess nostroke noext | M ch2 y4 A ch2 ch2 cd4 -5400000 A ch4 ch4 0 -10800000 Z ' +
        'M x4 ch A ch2 ch2 cd4 -16200000 A ch4 ch4 cd2 -10800000 Z',
      'none noext | M l y3 A ch2 ch2 cd2 cd4 L x3 ch L x3 ch2 A ch2 ch2 cd2 cd2 L r y5 A ch2 ch2 0 cd4 L ch y6 ' +
        'L ch y7 A ch2 ch2 0 cd2 Z M x3 ch L x4 ch A ch2 ch2 cd4 -5400000 M x4 ch L x4 ch2 A ch4 ch4 0 cd2 ' +
        'M ch2 y4 L ch2 y3 A ch4 ch4 cd2 cd2 A ch2 ch2 0 cd2 M ch y3 L ch y6',
    ],
  },
  verticalScroll: {
    av: 'adj 12500',
    gd: `a pin 0 adj 25000
ch */ ss a 100000
ch2 */ ch 1 2
ch4 */ ch 1 4
x3 +- ch ch2 0
x4 +- ch ch 0
x6 +- r 0 ch
x7 +- r 0 ch2
x5 +- x6 0 ch2
y3 +- b 0 ch
y4 +- b 0 ch2`,
    rect: 'x4 ch x6 y4',
    path: [
      'nostroke noext | M ch2 b A ch2 ch2 cd4 -5400000 L ch2 y4 A ch4 ch4 cd4 cd2 L ch y3 L ch ch2 A ch2 ch2 cd2 cd4 ' +
        'L x7 t A ch2 ch2 3cd4 cd2 L x6 ch L x6 y4 A ch2 ch2 0 cd4 Z M x4 ch2 A ch2 ch2 0 cd4 A ch4 ch4 cd4 cd2 Z',
      'darkenLess nostroke noext | M x4 ch2 A ch2 ch2 0 cd4 A ch4 ch4 cd4 cd2 Z M ch y4 A ch2 ch2 0 3cd4 A ch4 ch4 3cd4 cd2 Z',
      'none noext | M ch y3 L ch ch2 A ch2 ch2 cd2 cd4 L x7 t A ch2 ch2 3cd4 cd2 L x6 ch L x6 y4 A ch2 ch2 0 cd4 ' +
        'L ch2 b A ch2 ch2 cd4 cd2 Z M x3 t A ch2 ch2 3cd4 cd2 A ch4 ch4 cd4 cd2 L x4 ch2 M x6 ch L x3 ch ' +
        'M ch2 y3 A ch4 ch4 3cd4 cd2 L ch y4 M ch2 b A ch2 ch2 cd4 -5400000 L ch y3',
    ],
  },

  // --- Flowchart ------------------------------------------------------------------------------
  flowChartProcess: { path: 'w=1 h=1 | M 0 0 L 1 0 L 1 1 L 0 1 Z' },
  flowChartAlternateProcess: {
    gd: `x2 +- r 0 ssd6
y2 +- b 0 ssd6
il */ ssd6 29289 100000
ir +- r 0 il
ib +- b 0 il`,
    rect: 'il il ir ib',
    path: 'M l ssd6 A ssd6 ssd6 cd2 cd4 L x2 t A ssd6 ssd6 3cd4 cd4 L r y2 A ssd6 ssd6 0 cd4 L ssd6 b A ssd6 ssd6 cd4 cd4 Z',
  },
  flowChartDecision: {
    gd: `ir */ w 3 4
ib */ h 3 4`,
    rect: 'wd4 hd4 ir ib',
    path: 'w=2 h=2 | M 0 1 L 1 0 L 2 1 L 1 2 Z',
  },
  flowChartInputOutput: {
    gd: `x3 */ w 2 5
x4 */ w 3 5
x5 */ w 4 5
x6 */ w 9 10`,
    rect: 'wd5 t x5 b',
    path: 'w=5 h=5 | M 0 5 L 1 0 L 5 0 L 4 5 Z',
  },
  flowChartPredefinedProcess: {
    gd: 'x2 */ w 7 8',
    rect: 'wd8 t x2 b',
    path: [
      'nostroke noext w=1 h=1 | M 0 0 L 1 0 L 1 1 L 0 1 Z',
      'none noext w=8 h=8 | M 1 0 L 1 8 M 7 0 L 7 8',
      'none w=1 h=1 | M 0 0 L 1 0 L 1 1 L 0 1 Z',
    ],
  },
  flowChartInternalStorage: {
    rect: 'wd8 hd8 r b',
    path: [
      'nostroke noext w=1 h=1 | M 0 0 L 1 0 L 1 1 L 0 1 Z',
      'none noext w=8 h=8 | M 1 0 L 1 8 M 0 1 L 8 1',
      'none w=1 h=1 | M 0 0 L 1 0 L 1 1 L 0 1 Z',
    ],
  },
  flowChartDocument: {
    gd: `y1 */ h 17322 21600
y2 */ h 20172 21600`,
    rect: 'l t r y1',
    path: 'w=21600 h=21600 | M 0 0 L 21600 0 L 21600 17322 C 10800 17322 10800 23922 0 20172 Z',
  },
  flowChartMultidocument: {
    gd: `y2 */ h 3675 21600
y8 */ h 20782 21600
x3 */ w 9298 21600
x4 */ w 12286 21600
x5 */ w 18595 21600`,
    rect: 'l y2 x5 y8',
    path: [
      'nostroke noext w=21600 h=21600 | M 0 20782 C 9298 23542 9298 18022 18595 18022 L 18595 3675 L 0 3675 Z ' +
        'M 1532 3675 L 1532 1815 L 20000 1815 L 20000 16252 C 19298 16252 18595 16352 18595 16352 L 18595 3675 Z ' +
        'M 2972 1815 L 2972 0 L 21600 0 L 21600 14392 C 20800 14392 20000 14467 20000 14467 L 20000 1815 Z',
      'none w=21600 h=21600 | M 0 3675 L 18595 3675 L 18595 18022 C 9298 18022 9298 23542 0 20782 Z ' +
        'M 1532 3675 L 1532 1815 L 20000 1815 L 20000 16252 C 19298 16252 18595 16352 18595 16352 ' +
        'M 2972 1815 L 2972 0 L 21600 0 L 21600 14392 C 20800 14392 20000 14467 20000 14467',
    ],
  },
  flowChartTerminator: {
    gd: `il */ w 1018 21600
ir */ w 20582 21600
it */ h 3163 21600
ib */ h 18437 21600`,
    rect: 'il it ir ib',
    path: 'w=21600 h=21600 | M 3475 0 L 18125 0 A 3475 10800 3cd4 cd2 L 3475 21600 A 3475 10800 cd4 cd2 Z',
  },
  flowChartPreparation: {
    gd: 'x2 */ w 4 5',
    rect: 'wd5 t x2 b',
    path: 'w=10 h=10 | M 0 5 L 2 0 L 8 0 L 10 5 L 8 10 L 2 10 Z',
  },
  flowChartManualInput: {
    rect: 'l hd5 r b',
    path: 'w=5 h=5 | M 0 1 L 5 0 L 5 5 L 0 5 Z',
  },
  flowChartManualOperation: {
    gd: `x3 */ w 4 5
x4 */ w 9 10`,
    rect: 'wd5 t x3 b',
    path: 'w=5 h=5 | M 0 0 L 5 0 L 4 5 L 1 5 Z',
  },
  flowChartConnector: { gd: ELLIPSE_RECT_GD, rect: 'il it ir ib', path: ELLIPSE_PATH },
  flowChartOffpageConnector: {
    gd: 'y1 */ h 4 5',
    rect: 'l t r y1',
    path: 'w=10 h=10 | M 0 0 L 10 0 L 10 8 L 5 10 L 0 8 Z',
  },
  flowChartPunchedCard: {
    rect: 'l hd5 r b',
    path: 'w=5 h=5 | M 0 1 L 1 0 L 5 0 L 5 5 L 0 5 Z',
  },
  flowChartPunchedTape: {
    gd: `y2 */ h 9 10
ib */ h 4 5`,
    rect: 'l hd5 r ib',
    path: 'w=20 h=20 | M 0 2 A 5 2 cd2 -10800000 A 5 2 cd2 cd2 L 20 18 A 5 2 0 -10800000 A 5 2 0 cd2 Z',
  },
  flowChartSummingJunction: {
    gd: ELLIPSE_RECT_GD,
    rect: 'il it ir ib',
    path: [`nostroke noext | ${ELLIPSE_PATH}`, 'none noext | M il it L ir ib M ir it L il ib', `none | ${ELLIPSE_PATH}`],
  },
  flowChartOr: {
    gd: ELLIPSE_RECT_GD,
    rect: 'il it ir ib',
    path: [`nostroke noext | ${ELLIPSE_PATH}`, 'none noext | M hc t L hc b M l vc L r vc', `none | ${ELLIPSE_PATH}`],
  },
  flowChartCollate: {
    gd: `ir */ w 3 4
ib */ h 3 4`,
    rect: 'wd4 hd4 ir ib',
    path: 'w=2 h=2 | M 0 0 L 2 0 L 1 1 L 2 2 L 0 2 L 1 1 Z',
  },
  flowChartSort: {
    gd: `ir */ w 3 4
ib */ h 3 4`,
    rect: 'wd4 hd4 ir ib',
    path: [
      'nostroke noext w=2 h=2 | M 0 1 L 1 0 L 2 1 L 1 2 Z',
      'none noext w=2 h=2 | M 0 1 L 2 1',
      'none w=2 h=2 | M 0 1 L 1 0 L 2 1 L 1 2 Z',
    ],
  },
  flowChartExtract: {
    gd: 'x2 */ w 3 4',
    rect: 'wd4 vc x2 b',
    path: 'w=2 h=2 | M 0 2 L 1 0 L 2 2 Z',
  },
  flowChartMerge: {
    gd: 'x2 */ w 3 4',
    rect: 'wd4 t x2 vc',
    path: 'w=2 h=2 | M 0 0 L 2 0 L 1 2 Z',
  },
  flowChartOfflineStorage: {
    gd: 'x4 */ w 3 4',
    rect: 'wd4 t x4 vc',
    path: [
      'nostroke noext w=2 h=2 | M 0 0 L 2 0 L 1 2 Z',
      'none noext w=5 h=5 | M 2 4 L 3 4',
      'none w=2 h=2 | M 0 0 L 2 0 L 1 2 Z',
    ],
  },
  flowChartOnlineStorage: {
    gd: 'x2 */ w 5 6',
    rect: 'wd6 t x2 b',
    path: 'w=6 h=6 | M 1 0 L 6 0 A 1 3 3cd4 -10800000 L 1 6 A 1 3 cd4 cd2 Z',
  },
  flowChartDelay: {
    gd: `idx cos wd2 2700000
idy sin hd2 2700000
ir +- hc idx 0
it +- vc 0 idy
ib +- vc idy 0`,
    rect: 'l it ir ib',
    path: 'M l t L hc t A wd2 hd2 3cd4 cd2 L l b Z',
  },
  flowChartMagneticTape: {
    gd: `${ELLIPSE_RECT_GD}
ang1 at2 w h`,
    rect: 'il it ir ib',
    path: 'M hc b A wd2 hd2 cd4 cd4 A wd2 hd2 cd2 cd4 A wd2 hd2 3cd4 cd4 A wd2 hd2 0 ang1 L r ib L r b Z',
  },
  flowChartMagneticDisk: {
    gd: 'y3 */ h 3 6',
    rect: 'l y3 r b',
    path: [
      'nostroke noext w=6 h=6 | M 0 1 A 3 1 cd2 cd2 L 6 5 A 3 1 0 cd2 Z',
      'none noext w=6 h=6 | M 6 1 A 3 1 0 cd2',
      'none w=6 h=6 | M 0 1 A 3 1 cd2 cd2 L 6 5 A 3 1 0 cd2 Z',
    ],
  },
  flowChartMagneticDrum: {
    gd: 'x2 */ w 2 3',
    rect: 'wd6 t x2 b',
    path: [
      'nostroke noext w=6 h=6 | M 1 0 L 5 0 A 1 3 3cd4 cd2 L 1 6 A 1 3 cd4 cd2 Z',
      'none noext w=6 h=6 | M 5 6 A 1 3 cd4 cd2',
      'none w=6 h=6 | M 1 0 L 5 0 A 1 3 3cd4 cd2 L 1 6 A 1 3 cd4 cd2 Z',
    ],
  },
  flowChartDisplay: {
    gd: 'x2 */ w 5 6',
    rect: 'wd6 t x2 b',
    path: 'w=6 h=6 | M 0 3 L 1 0 L 5 0 A 1 3 3cd4 cd2 L 1 6 Z',
  },
};

const cache = new Map<string, ShapeDefinition>();

/** Names of every preset this module can draw. */
export const PRESET_NAMES: readonly string[] = Object.keys(SOURCES);

/** The parsed definition of a preset shape, or undefined for one this module does not have. */
export function presetDefinition(name: string): ShapeDefinition | undefined {
  const hit = cache.get(name);
  if (hit) return hit;
  if (!Object.prototype.hasOwnProperty.call(SOURCES, name)) return undefined;
  const src = SOURCES[name];
  if (!src) return undefined;
  const def = parseSrc(src);
  cache.set(name, def);
  return def;
}
