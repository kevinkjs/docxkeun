/**
 * GDEF: glyph classes (base / ligature / mark / component), mark attachment classes and
 * mark glyph sets. Lookup flags filter glyphs through these, so they decide which glyphs a
 * ligature or kerning lookup "sees" past (e.g. a kern pair that skips an accent).
 */
import { BinaryData, parseClassDef, parseCoverage, type ClassDef, type Coverage } from './layout-common.js';

/** Glyph property bits, aligned with the LookupFlag Ignore* bits so one AND tests both. */
export const GlyphProps = {
  BaseGlyph: 0x02,
  Ligature: 0x04,
  Mark: 0x08,
  /** History bits kept across substitutions (Indic final reordering inspects them). */
  Substituted: 0x10,
  Ligated: 0x20,
  Multiplied: 0x40,
  Preserve: 0x70,
  /** The mark attachment class lives in the high byte, like LookupFlag.MarkAttachmentType. */
  MarkAttachClass: 0xff00,
} as const;

export class Gdef {
  readonly glyphClasses: ClassDef | undefined;
  private readonly markAttachClasses: ClassDef | undefined;
  private readonly markSets: Coverage[] = [];

  constructor(bytes: Uint8Array | undefined) {
    if (!bytes || bytes.length < 12) return;
    const d = new BinaryData(bytes);
    const minor = d.u16(2);
    const classDef = d.u16(4);
    const markAttach = d.u16(10);
    if (classDef) this.glyphClasses = parseClassDef(d, classDef);
    if (markAttach) this.markAttachClasses = parseClassDef(d, markAttach);
    if (minor >= 2 && bytes.length >= 14) {
      const sets = d.u16(12);
      if (sets && d.u16(sets) === 1) {
        const count = d.u16(sets + 2);
        for (let i = 0; i < count; i++) {
          const rel = d.u32(sets + 4 + i * 4);
          this.markSets.push(rel ? parseCoverage(d, sets + rel) : new Map());
        }
      }
    }
  }

  /** True when the font classifies glyphs itself (otherwise classes come from Unicode). */
  get hasGlyphClasses(): boolean {
    return this.glyphClasses !== undefined && this.glyphClasses.size > 0;
  }

  /** Glyph property bits for a glyph (0 for unclassified / component glyphs). */
  glyphProps(glyph: number): number {
    switch (this.glyphClasses?.get(glyph) ?? 0) {
      case 1:
        return GlyphProps.BaseGlyph;
      case 2:
        return GlyphProps.Ligature;
      case 3:
        return GlyphProps.Mark | ((this.markAttachClasses?.get(glyph) ?? 0) << 8);
      default:
        return 0;
    }
  }

  markSetCovers(set: number, glyph: number): boolean {
    return this.markSets[set]?.has(glyph) ?? false;
  }
}
