/** Unit conversions. Everything inside the model and layout is in points (1/72 inch). */

export const TWIPS_PER_PT = 20;
export const EMU_PER_PT = 12700;
export const EMU_PER_INCH = 914400;

export function twipsToPt(v: number): number {
  return v / TWIPS_PER_PT;
}
export function ptToTwips(v: number): number {
  return Math.round(v * TWIPS_PER_PT);
}
export function halfPtToPt(v: number): number {
  return v / 2;
}
export function eighthPtToPt(v: number): number {
  return v / 8;
}
export function emuToPt(v: number): number {
  return v / EMU_PER_PT;
}
export function ptToEmu(v: number): number {
  return Math.round(v * EMU_PER_PT);
}

/** ST_OnOff: absent `w:val` means on. */
export function parseOnOff(value: string | undefined): boolean {
  if (value === undefined) return true;
  return !(value === '0' || value === 'false' || value === 'off');
}

export function parseIntSafe(value: string | undefined): number | undefined {
  if (value === undefined || value === '') return undefined;
  const n = Number.parseInt(value, 10);
  return Number.isFinite(n) ? n : undefined;
}

export function parseFloatSafe(value: string | undefined): number | undefined {
  if (value === undefined || value === '') return undefined;
  const n = Number.parseFloat(value);
  return Number.isFinite(n) ? n : undefined;
}

/**
 * ST_TwipsMeasure / ST_SignedTwipsMeasure: a number of twips, or (Strict and some writers)
 * a number with a unit suffix such as `12pt`, `1in`, `2.5cm`, `10mm`, `3pc`, `1pi`.
 */
export function parseTwipsMeasure(value: string | undefined): number | undefined {
  if (value === undefined || value === '') return undefined;
  const m = /^(-?\d+(?:\.\d+)?)(mm|cm|in|pt|pc|pi)?$/.exec(value.trim());
  if (!m) return undefined;
  const n = Number.parseFloat(m[1]!);
  switch (m[2]) {
    case undefined:
      return twipsToPt(n);
    case 'pt':
      return n;
    case 'in':
      return n * 72;
    case 'cm':
      return (n * 72) / 2.54;
    case 'mm':
      return (n * 72) / 25.4;
    default:
      return n * 12; // pc / pi
  }
}

export type WidthType = 'dxa' | 'pct' | 'auto' | 'nil';

export interface TableWidth {
  type: WidthType;
  /** Points for dxa; fraction (0..1) for pct. */
  value: number;
}

/** ST_TblWidth: dxa twips, pct in fiftieths of a percent (or `50%`), auto, nil. */
export function parseTableWidth(w: string | undefined, type: string | undefined): TableWidth | undefined {
  const t = (type ?? 'dxa') as WidthType;
  if (t === 'auto' || t === 'nil') return { type: t, value: 0 };
  if (w === undefined) return undefined;
  if (t === 'pct') {
    if (w.endsWith('%')) return { type: 'pct', value: Number.parseFloat(w) / 100 };
    return { type: 'pct', value: (Number.parseFloat(w) || 0) / 5000 };
  }
  const pt = parseTwipsMeasure(w);
  return pt === undefined ? undefined : { type: 'dxa', value: pt };
}
