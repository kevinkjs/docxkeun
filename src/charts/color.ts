/**
 * DrawingML colours as charts use them: a:srgbClr, a:schemeClr, a:sysClr, a:prstClr, a:scrgbClr and
 * a:hslClr, with the transforms Office applies (lumMod/lumOff/sat/hue in HSL, tint/shade in linear
 * RGB, alpha), plus the fill and line properties of a c:spPr.
 */
import { attr, elements, firstChild, type XmlElement } from '../xml/dom.js';

/** Theme colour lookup by scheme name (accent1, tx1, bg1, dk1, lt1, …) → RRGGBB. */
export type SchemeColorFn = (name: string) => string | undefined;

/** A colour with its opacity (0..1). */
export interface Paint {
  color: string;
  opacity: number;
}

/** Office's default theme (2013–2022), used when the caller's theme does not know a name. */
const DEFAULT_SCHEME: Record<string, string> = {
  dk1: '000000',
  lt1: 'FFFFFF',
  dk2: '44546A',
  lt2: 'E7E6E6',
  accent1: '4472C4',
  accent2: 'ED7D31',
  accent3: 'A5A5A5',
  accent4: 'FFC000',
  accent5: '5B9BD5',
  accent6: '70AD47',
  hlink: '0563C1',
  folHlink: '954F72',
};

/** The text/background names map onto dark/light slots (Word's default colour map). */
const SCHEME_ALIASES: Record<string, string> = { tx1: 'dk1', tx2: 'dk2', bg1: 'lt1', bg2: 'lt2', dk1: 'tx1', lt1: 'bg1', dk2: 'tx2', lt2: 'bg2' };

const PRESET: Record<string, string> = {
  black: '000000',
  white: 'FFFFFF',
  red: 'FF0000',
  green: '008000',
  lime: '00FF00',
  blue: '0000FF',
  yellow: 'FFFF00',
  cyan: '00FFFF',
  aqua: '00FFFF',
  magenta: 'FF00FF',
  fuchsia: 'FF00FF',
  gray: '808080',
  grey: '808080',
  silver: 'C0C0C0',
  maroon: '800000',
  navy: '000080',
  olive: '808000',
  orange: 'FFA500',
  purple: '800080',
  teal: '008080',
  ltGray: 'D3D3D3',
  dkGray: 'A9A9A9',
};

/** Look a scheme name up in the caller's theme, then through the tx/bg aliases, then Office's default theme. */
export function lookupScheme(name: string, scheme: SchemeColorFn | undefined): string {
  const alias = SCHEME_ALIASES[name];
  const found = scheme?.(name) ?? (alias ? scheme?.(alias) : undefined) ?? DEFAULT_SCHEME[name] ?? (alias ? DEFAULT_SCHEME[alias] : undefined);
  return (found ?? '000000').toUpperCase();
}

type Rgb = [number, number, number];

function hexToRgb(hex: string): Rgb {
  const h = /^[0-9a-fA-F]{6}$/.test(hex) ? hex : '000000';
  return [0, 2, 4].map((i) => Number.parseInt(h.slice(i, i + 2), 16) / 255) as Rgb;
}

function rgbToHex([r, g, b]: Rgb): string {
  return [r, g, b]
    .map((v) => Math.round(Math.max(0, Math.min(1, v)) * 255).toString(16).padStart(2, '0'))
    .join('')
    .toUpperCase();
}

function rgbToHsl([r, g, b]: Rgb): Rgb {
  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  const l = (max + min) / 2;
  if (max === min) return [0, 0, l];
  const d = max - min;
  const s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
  let h: number;
  if (max === r) h = (g - b) / d + (g < b ? 6 : 0);
  else if (max === g) h = (b - r) / d + 2;
  else h = (r - g) / d + 4;
  return [h / 6, s, l];
}

function hslToRgb([h, s, l]: Rgb): Rgb {
  if (s === 0) return [l, l, l];
  const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
  const p = 2 * l - q;
  const channel = (t: number): number => {
    if (t < 0) t += 1;
    if (t > 1) t -= 1;
    if (t < 1 / 6) return p + (q - p) * 6 * t;
    if (t < 1 / 2) return q;
    if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6;
    return p;
  };
  return [channel(h + 1 / 3), channel(h), channel(h - 1 / 3)];
}

// tint and shade work on linear light (scRGB), which is why a 50% shade is darker than halving sRGB.
const toLinear = (c: number): number => (c <= 0.04045 ? c / 12.92 : ((c + 0.055) / 1.055) ** 2.4);
const toSrgb = (c: number): number => (c <= 0.0031308 ? c * 12.92 : 1.055 * c ** (1 / 2.4) - 0.055);

const clamp01 = (v: number): number => Math.max(0, Math.min(1, v));

/** Apply a list of DrawingML colour transforms (children of the colour element) in document order. */
function applyTransforms(rgbIn: Rgb, alphaIn: number, transforms: XmlElement[]): { rgb: Rgb; alpha: number } {
  let rgb = rgbIn;
  let alpha = alphaIn;
  const hsl = (f: (v: Rgb) => Rgb): void => {
    rgb = hslToRgb(f(rgbToHsl(rgb)));
  };
  for (const t of transforms) {
    const raw = Number(attr(t, 'val') ?? '0');
    const v = Number.isFinite(raw) ? raw / 100000 : 0;
    switch (t.name.slice(t.name.indexOf(':') + 1)) {
      case 'lumMod':
        hsl(([h, s, l]) => [h, s, clamp01(l * v)]);
        break;
      case 'lumOff':
        hsl(([h, s, l]) => [h, s, clamp01(l + v)]);
        break;
      case 'lum':
        hsl(([h, s]) => [h, s, clamp01(v)]);
        break;
      case 'satMod':
        hsl(([h, s, l]) => [h, clamp01(s * v), l]);
        break;
      case 'satOff':
        hsl(([h, s, l]) => [h, clamp01(s + v), l]);
        break;
      case 'sat':
        hsl(([h, , l]) => [h, clamp01(v), l]);
        break;
      case 'hueMod':
        hsl(([h, s, l]) => [(((h * v) % 1) + 1) % 1, s, l]);
        break;
      case 'hueOff':
        // Angles are in 60000ths of a degree, so the 100000 scaling above is undone here.
        hsl(([h, s, l]) => [(((h + (raw / 60000 / 360)) % 1) + 1) % 1, s, l]);
        break;
      case 'hue':
        hsl(([, s, l]) => [((raw / 60000 / 360) % 1 + 1) % 1, s, l]);
        break;
      case 'tint':
        // Towards white in linear space: c' = 1 - (1 - c) * tint.
        rgb = rgb.map((c) => toSrgb(1 - (1 - toLinear(c)) * clamp01(v))) as Rgb;
        break;
      case 'shade':
        rgb = rgb.map((c) => toSrgb(toLinear(c) * clamp01(v))) as Rgb;
        break;
      case 'inv':
        rgb = rgb.map((c) => 1 - c) as Rgb;
        break;
      case 'comp':
        hsl(([h, s, l]) => [(h + 0.5) % 1, s, l]);
        break;
      case 'gray': {
        const y = 0.299 * rgb[0] + 0.587 * rgb[1] + 0.114 * rgb[2];
        rgb = [y, y, y];
        break;
      }
      case 'alpha':
        alpha = clamp01(v);
        break;
      case 'alphaMod':
        alpha = clamp01(alpha * v);
        break;
      case 'alphaOff':
        alpha = clamp01(alpha + v);
        break;
      default:
        break;
    }
  }
  return { rgb, alpha };
}

const COLOR_ELEMENTS = new Set(['a:srgbClr', 'a:schemeClr', 'a:sysClr', 'a:prstClr', 'a:scrgbClr', 'a:hslClr']);

/** Resolve a colour element itself (a:srgbClr, a:schemeClr, …). `phClr` stands for `placeholder` when given. */
export function resolveColorElement(node: XmlElement, scheme: SchemeColorFn | undefined, placeholder?: string): Paint | undefined {
  let base: Rgb;
  switch (node.name) {
    case 'a:srgbClr':
      base = hexToRgb(attr(node, 'val') ?? '000000');
      break;
    case 'a:schemeClr': {
      const name = attr(node, 'val') ?? 'tx1';
      base = hexToRgb(name === 'phClr' ? (placeholder ?? '000000') : lookupScheme(name, scheme));
      break;
    }
    case 'a:sysClr':
      base = hexToRgb(attr(node, 'lastClr') ?? (attr(node, 'val') === 'window' ? 'FFFFFF' : '000000'));
      break;
    case 'a:prstClr':
      base = hexToRgb(PRESET[attr(node, 'val') ?? ''] ?? '000000');
      break;
    case 'a:scrgbClr': {
      const c = (n: string): number => toSrgb(clamp01(Number(attr(node, n) ?? '0') / 100000));
      base = [c('r'), c('g'), c('b')];
      break;
    }
    case 'a:hslClr': {
      const h = Number(attr(node, 'hue') ?? '0') / 60000 / 360;
      base = hslToRgb([h, clamp01(Number(attr(node, 'sat') ?? '0') / 100000), clamp01(Number(attr(node, 'lum') ?? '0') / 100000)]);
      break;
    }
    default:
      return undefined;
  }
  const { rgb, alpha } = applyTransforms(base, 1, elements(node));
  return { color: rgbToHex(rgb), opacity: alpha };
}

/** The colour inside a container such as a:solidFill, a:fgClr or a:gs. */
export function colorIn(container: XmlElement | undefined, scheme: SchemeColorFn | undefined, placeholder?: string): Paint | undefined {
  if (!container) return undefined;
  for (const c of elements(container)) if (COLOR_ELEMENTS.has(c.name)) return resolveColorElement(c, scheme, placeholder);
  return undefined;
}

/** A scheme colour with lumMod/lumOff (the series palette variations), computed without XML. */
export function schemeVariant(name: string, scheme: SchemeColorFn | undefined, lumMod = 1, lumOff = 0): string {
  const [h, s, l] = rgbToHsl(hexToRgb(lookupScheme(name, scheme)));
  return rgbToHex(hslToRgb([h, s, clamp01(clamp01(l * lumMod) + lumOff)]));
}

/** Shift a colour's HSL lightness (used for derived shades such as monochrome chart styles). */
export function adjustLightness(hex: string, lumMod: number, lumOff = 0): string {
  const [h, s, l] = rgbToHsl(hexToRgb(hex));
  return rgbToHex(hslToRgb([h, s, clamp01(clamp01(l * lumMod) + lumOff)]));
}

/** Relative brightness 0..1, to pick a readable text colour over a fill. */
export function brightness(hex: string): number {
  const [r, g, b] = hexToRgb(hex);
  return 0.299 * r + 0.587 * g + 0.114 * b;
}

/**
 * The fill an element group (c:spPr, a:ln) declares: `null` for a:noFill, undefined when nothing is
 * said (the chart's automatic formatting applies). Gradients use their first stop and patterns their
 * foreground, since the display list only has flat fills.
 */
export function fillOf(spPr: XmlElement | undefined, scheme: SchemeColorFn | undefined, placeholder?: string): Paint | null | undefined {
  if (!spPr) return undefined;
  for (const c of elements(spPr)) {
    switch (c.name) {
      case 'a:noFill':
        return null;
      case 'a:solidFill':
        return colorIn(c, scheme, placeholder) ?? undefined;
      case 'a:gradFill': {
        const stops = elements(firstChild(c, 'a:gsLst') ?? c, 'a:gs');
        const first = stops[0];
        return first ? (colorIn(first, scheme, placeholder) ?? undefined) : undefined;
      }
      case 'a:pattFill':
        return colorIn(firstChild(c, 'a:fgClr'), scheme, placeholder) ?? undefined;
      default:
        break;
    }
  }
  return undefined;
}

/** Line properties of a c:spPr's a:ln (or of an a:ln directly). Unset parts stay undefined. */
export interface LineProps {
  /** `null` for a:noFill (no line). */
  paint?: Paint | null;
  /** Width in points. */
  width?: number;
  /** a:prstDash value (solid, dash, sysDot, …). */
  dash?: string;
  cap?: 'butt' | 'round' | 'square';
  join?: 'miter' | 'round' | 'bevel';
}

export function lineOf(spPr: XmlElement | undefined, scheme: SchemeColorFn | undefined, placeholder?: string): LineProps {
  const ln = spPr?.name === 'a:ln' ? spPr : firstChild(spPr, 'a:ln');
  if (!ln) return {};
  const out: LineProps = {};
  const paint = fillOf(ln, scheme, placeholder);
  if (paint !== undefined) out.paint = paint;
  const w = attr(ln, 'w');
  if (w !== undefined && Number.isFinite(Number(w))) out.width = Number(w) / 12700;
  const dash = attr(firstChild(ln, 'a:prstDash'), 'val');
  if (dash) out.dash = dash;
  const cap = attr(ln, 'cap');
  if (cap === 'rnd') out.cap = 'round';
  else if (cap === 'sq') out.cap = 'square';
  else if (cap === 'flat') out.cap = 'butt';
  if (firstChild(ln, 'a:round')) out.join = 'round';
  else if (firstChild(ln, 'a:bevel')) out.join = 'bevel';
  else if (firstChild(ln, 'a:miter')) out.join = 'miter';
  return out;
}

/** Dash pattern in points for an a:prstDash preset, scaled by the line width as Office does. */
export function dashPattern(dash: string | undefined, width: number): number[] | undefined {
  const w = Math.max(width, 0.5);
  switch (dash) {
    case undefined:
    case 'solid':
      return undefined;
    case 'dot':
      return [w, 3 * w];
    case 'sysDot':
      return [w, w];
    case 'sysDash':
      return [3 * w, w];
    case 'dash':
      return [4 * w, 3 * w];
    case 'lgDash':
      return [8 * w, 3 * w];
    case 'dashDot':
      return [4 * w, 3 * w, w, 3 * w];
    case 'sysDashDot':
      return [3 * w, w, w, w];
    case 'lgDashDot':
      return [8 * w, 3 * w, w, 3 * w];
    case 'lgDashDotDot':
      return [8 * w, 3 * w, w, 3 * w, w, 3 * w];
    case 'sysDashDotDot':
      return [3 * w, w, w, w, w, w];
    default:
      return [4 * w, 3 * w];
  }
}
