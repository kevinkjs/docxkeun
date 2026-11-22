/**
 * The shaping buffer: one record per glyph, holding the glyph id, the source character and
 * cluster, feature mask, GDEF-derived properties, ligature-component bookkeeping (so marks
 * can attach to the right component of a ligature) and, once positioning starts, the
 * advance/offset/attachment fields. Substitutions mutate the array in place.
 */
import { GlyphProps } from './gdef.js';

export interface GlyphInfo {
  glyph: number;
  /** Unicode code point the glyph came from (the first one, for ligatures). */
  codepoint: number;
  /** UTF-16 index of the first character of the cluster. */
  cluster: number;
  /** Feature bits enabled on this glyph. */
  mask: number;
  /** GlyphProps bits (class, substitution history, mark attachment class). */
  props: number;
  /** UFlag bits of the source character. */
  uflags: number;
  /** Ligature id shared by a ligature and the marks attached to its components (0 = none). */
  ligId: number;
  /** Component index for marks/components, or the component count on a ligature base. */
  ligComp: number;
  isLigBase: boolean;
  /** Syllable serial (Indic); 0 means no per-syllable restriction. */
  syllable: number;
  /** Shaper-private: Indic category, Arabic joining action. */
  cat: number;
  /** Shaper-private: Indic position. */
  pos: number;
  xAdvance: number;
  yAdvance: number;
  xOffset: number;
  yOffset: number;
  /** Attachment kind (ATTACH_MARK / ATTACH_CURSIVE) and the relative index of the parent. */
  attachType: number;
  attachChain: number;
}

export const ATTACH_MARK = 1;
export const ATTACH_CURSIVE = 2;

export function newGlyphInfo(codepoint: number, cluster: number, uflags: number): GlyphInfo {
  return {
    glyph: 0,
    codepoint,
    cluster,
    mask: 0,
    props: 0,
    uflags,
    ligId: 0,
    ligComp: 0,
    isLigBase: false,
    syllable: 0,
    cat: 0,
    pos: 0,
    xAdvance: 0,
    yAdvance: 0,
    xOffset: 0,
    yOffset: 0,
    attachType: 0,
    attachChain: 0,
  };
}

export function cloneInfo(info: GlyphInfo): GlyphInfo {
  return { ...info };
}

export const isMark = (info: GlyphInfo): boolean => (info.props & GlyphProps.Mark) !== 0;
export const isBaseGlyph = (info: GlyphInfo): boolean => (info.props & GlyphProps.BaseGlyph) !== 0;
export const isLigature = (info: GlyphInfo): boolean => (info.props & GlyphProps.Ligature) !== 0;
export const isSubstituted = (info: GlyphInfo): boolean => (info.props & GlyphProps.Substituted) !== 0;
export const isLigated = (info: GlyphInfo): boolean => (info.props & GlyphProps.Ligated) !== 0;
export const isMultiplied = (info: GlyphInfo): boolean => (info.props & GlyphProps.Multiplied) !== 0;
export const ligatedAndDidntMultiply = (info: GlyphInfo): boolean => isLigated(info) && !isMultiplied(info);

export function ligComp(info: GlyphInfo): number {
  return info.isLigBase ? 0 : info.ligComp;
}

export function ligNumComps(info: GlyphInfo): number {
  return info.isLigBase && isLigature(info) ? info.ligComp : 1;
}

export function setLigPropsForLigature(info: GlyphInfo, ligId: number, numComps: number): void {
  info.ligId = ligId;
  info.isLigBase = true;
  info.ligComp = numComps & 0x0f;
}

export function setLigPropsForMark(info: GlyphInfo, ligId: number, comp: number): void {
  info.ligId = ligId;
  info.isLigBase = false;
  info.ligComp = comp & 0x0f;
}

/**
 * Gives every glyph in [start, end) the smallest cluster value among them, widening the range
 * over neighbours that already share an edge cluster so clusters stay contiguous.
 */
export function mergeClusters(buf: GlyphInfo[], start: number, end: number): void {
  if (end - start < 2) return;
  let cluster = buf[start]!.cluster;
  for (let i = start + 1; i < end; i++) cluster = Math.min(cluster, buf[i]!.cluster);
  if (cluster !== buf[end - 1]!.cluster) while (end < buf.length && buf[end - 1]!.cluster === buf[end]!.cluster) end++;
  if (cluster !== buf[start]!.cluster) while (start > 0 && buf[start - 1]!.cluster === buf[start]!.cluster) start--;
  for (let i = start; i < end; i++) buf[i]!.cluster = cluster;
}

/** Removes glyph `i`, folding its cluster into a neighbour when it was the last of it. */
export function deleteGlyph(buf: GlyphInfo[], i: number): void {
  const cluster = buf[i]!.cluster;
  const next = buf[i + 1];
  const prev = buf[i - 1];
  if (!((next && next.cluster === cluster) || (prev && prev.cluster === cluster))) {
    if (prev) {
      if (cluster < prev.cluster) {
        const old = prev.cluster;
        for (let j = i - 1; j >= 0 && buf[j]!.cluster === old; j--) buf[j]!.cluster = cluster;
      }
    } else if (next) {
      mergeClusters(buf, i, i + 2);
    }
  }
  buf.splice(i, 1);
}
