/**
 * Device-context state shared by the EMF and WMF players: the GDI objects a picture selects
 * (pens, brushes, fonts), the attributes SaveDC/RestoreDC preserve, GDI's constants and its
 * stock objects. Points in the state are replaced, never mutated, so a saved copy stays intact.
 */
import { IDENTITY } from './geometry.js';
import type { Matrix, Pt } from './geometry.js';

export const MM_TEXT = 1;
export const MM_LOMETRIC = 2;
export const MM_HIMETRIC = 3;
export const MM_LOENGLISH = 4;
export const MM_HIENGLISH = 5;
export const MM_TWIPS = 6;
export const MM_ISOTROPIC = 7;
export const MM_ANISOTROPIC = 8;

export const PS_SOLID = 0;
export const PS_DASH = 1;
export const PS_DOT = 2;
export const PS_DASHDOT = 3;
export const PS_DASHDOTDOT = 4;
export const PS_NULL = 5;
export const PS_INSIDEFRAME = 6;
export const PS_USERSTYLE = 7;
export const PS_ALTERNATE = 8;
export const PS_GEOMETRIC = 0x10000;

export const BS_SOLID = 0;
export const BS_NULL = 1;
export const BS_HATCHED = 2;
export const BS_PATTERN = 3;
export const BS_DIBPATTERN = 5;
export const BS_DIBPATTERNPT = 6;

export const TRANSPARENT = 1;
export const OPAQUE = 2;

export const ALTERNATE = 1;
export const WINDING = 2;

export const AD_COUNTERCLOCKWISE = 1;
export const AD_CLOCKWISE = 2;

export const TA_UPDATECP = 1;
export const TA_RIGHT = 2;
export const TA_CENTER = 6;
export const TA_BOTTOM = 8;
export const TA_BASELINE = 24;

export const ETO_OPAQUE = 0x2;
export const ETO_CLIPPED = 0x4;
export const ETO_GLYPH_INDEX = 0x10;
export const ETO_NO_RECT = 0x100;
export const ETO_SMALL_CHARS = 0x200;
export const ETO_PDY = 0x2000;

/** Raster operations the players understand; anything else that needs the destination is skipped. */
export const SRCCOPY = 0x00cc0020;
export const PATCOPY = 0x00f00021;
export const BLACKNESS = 0x00000042;
export const WHITENESS = 0x00ff0062;

export interface Pen {
  kind: 'pen';
  /** PS_* line style (the low nibble of the pen style). */
  style: number;
  /** Width in logical units; 0 means one device pixel. */
  width: number;
  /** Cosmetic pens are one device pixel wide whatever the transform. */
  cosmetic: boolean;
  color: string;
  cap: 'round' | 'square' | 'butt';
  join: 'round' | 'bevel' | 'miter';
  /** PS_USERSTYLE dash lengths (logical units for geometric pens, pixels for cosmetic ones). */
  dashes?: number[];
}

export interface Brush {
  kind: 'brush';
  style: 'solid' | 'null' | 'hatched' | 'pattern';
  color: string;
}

export interface Font {
  kind: 'font';
  /** Logical units: negative is the em height, positive the cell height, 0 the default size. */
  height: number;
  weight: number;
  italic: boolean;
  underline: boolean;
  strikeOut: boolean;
  charset: number;
  faceName: string;
}

/** Palettes, regions and colour spaces: they occupy object slots but draw nothing here. */
export interface OtherObject {
  kind: 'other';
}

export type GdiObject = Pen | Brush | Font | OtherObject;

export interface DCState {
  mapMode: number;
  windowOrg: Pt;
  windowExt: Pt;
  viewportOrg: Pt;
  viewportExt: Pt;
  /** World → page transform (EMF only; identity for WMF). */
  world: Matrix;
  /** Set once a world transform is used: GDI then is in GM_ADVANCED, which changes how arcs turn. */
  advanced: boolean;
  pen: Pen;
  brush: Brush;
  font: Font;
  textColor: string;
  bkColor: string;
  bkMode: number;
  textAlign: number;
  polyFillMode: number;
  arcDirection: number;
  /** Current position in logical units. */
  position: Pt;
  /** Extra space after each character (SetTextCharacterExtra), logical units. */
  charExtra: number;
}

/** Pen style bits (PS_ENDCAP_*, PS_JOIN_*, PS_GEOMETRIC) to a pen. */
export function makePen(styleBits: number, width: number, color: string, dashes?: number[]): Pen {
  const capBits = styleBits & 0xf00;
  const joinBits = styleBits & 0xf000;
  const style = styleBits & 0xf;
  const pen: Pen = {
    kind: 'pen',
    style,
    width: Math.abs(width),
    // ExtCreatePen marks geometric pens; a LOGPEN's width is always in logical units.
    cosmetic: false,
    color,
    cap: capBits === 0x100 ? 'square' : capBits === 0x200 ? 'butt' : 'round',
    join: joinBits === 0x1000 ? 'bevel' : joinBits === 0x2000 ? 'miter' : 'round',
  };
  if (dashes && dashes.length) pen.dashes = dashes;
  return pen;
}

export function makeBrush(style: number, color: string): Brush {
  switch (style) {
    case BS_SOLID:
      return { kind: 'brush', style: 'solid', color };
    case BS_HATCHED:
      return { kind: 'brush', style: 'hatched', color };
    case BS_PATTERN:
    case BS_DIBPATTERN:
    case BS_DIBPATTERNPT:
      // Without the bitmap at hand a pattern reads best as mid grey.
      return { kind: 'brush', style: 'pattern', color: '808080' };
    default:
      return { kind: 'brush', style: 'null', color };
  }
}

const SYSTEM_FONT: Font = { kind: 'font', height: 0, weight: 400, italic: false, underline: false, strikeOut: false, charset: 0, faceName: 'Arial' };

/** GDI stock objects, addressed in metafiles with the high bit set (EMF) or not at all (WMF). */
export function stockObject(index: number): GdiObject | undefined {
  const brush = (color: string): Brush => ({ kind: 'brush', style: 'solid', color });
  const pen = (color: string): Pen => ({ kind: 'pen', style: PS_SOLID, width: 0, cosmetic: true, color, cap: 'round', join: 'round' });
  switch (index) {
    case 0:
      return brush('FFFFFF');
    case 1:
      return brush('C0C0C0');
    case 2:
      return brush('808080');
    case 3:
      return brush('404040');
    case 4:
      return brush('000000');
    case 5:
      return { kind: 'brush', style: 'null', color: '000000' };
    case 6:
      return pen('FFFFFF');
    case 7:
      return pen('000000');
    case 8:
      return { ...pen('000000'), style: PS_NULL };
    case 10:
    case 11:
    case 16:
      return { ...SYSTEM_FONT, faceName: 'Courier New' };
    case 12:
    case 13:
    case 14:
    case 17:
      return SYSTEM_FONT;
    case 18:
      return brush('FFFFFF');
    case 19:
      return pen('000000');
    default:
      return undefined;
  }
}

export function defaultState(): DCState {
  return {
    mapMode: MM_TEXT,
    windowOrg: { x: 0, y: 0 },
    windowExt: { x: 1, y: 1 },
    viewportOrg: { x: 0, y: 0 },
    viewportExt: { x: 1, y: 1 },
    world: IDENTITY,
    advanced: false,
    pen: stockObject(7) as Pen,
    brush: stockObject(0) as Brush,
    font: SYSTEM_FONT,
    textColor: '000000',
    bkColor: 'FFFFFF',
    bkMode: OPAQUE,
    textAlign: 0,
    polyFillMode: ALTERNATE,
    arcDirection: AD_COUNTERCLOCKWISE,
    position: { x: 0, y: 0 },
    charExtra: 0,
  };
}

/** Make a selected object current in the state; other kinds are ignored. */
export function selectInto(state: DCState, obj: GdiObject | undefined): void {
  if (!obj) return;
  if (obj.kind === 'pen') state.pen = obj;
  else if (obj.kind === 'brush') state.brush = obj;
  else if (obj.kind === 'font') state.font = obj;
}

/** A COLORREF (R, G, B, flags) as RRGGBB. Palette-index colours keep their bytes as an approximation. */
export function colorRef(r: number, g: number, b: number): string {
  return hex2(r) + hex2(g) + hex2(b);
}

function hex2(v: number): string {
  return (Math.max(0, Math.min(255, Math.round(v))) | 0).toString(16).padStart(2, '0').toUpperCase();
}

/** Mix two RRGGBB colours: `t` of `a` and the rest of `b`. */
export function mixColors(a: string, b: string, t: number): string {
  const ch = (s: string, i: number): number => Number.parseInt(s.slice(i, i + 2), 16) || 0;
  return colorRef(ch(a, 0) * t + ch(b, 0) * (1 - t), ch(a, 2) * t + ch(b, 2) * (1 - t), ch(a, 4) * t + ch(b, 4) * (1 - t));
}
