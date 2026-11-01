/** Font provider backed by registered font files (exact metrics; works in Node and browsers). */
import type { FontRegistry, ResolvedFace } from '../fonts/registry.js';
import type { FontProvider, LayoutFont } from './types.js';

class SfntLayoutFont implements LayoutFont {
  readonly key: string;
  readonly family: string;
  readonly ascent: number;
  readonly descent: number;
  readonly underlinePosition: number;
  readonly underlineThickness: number;
  readonly face;
  readonly syntheticBold: boolean;
  readonly syntheticItalic: boolean;
  readonly bold: boolean;
  readonly italic: boolean;
  private readonly cache = new Map<number, number>();
  private readonly upem: number;

  constructor(resolved: ResolvedFace) {
    const { face } = resolved;
    const f = face.font;
    this.face = face;
    this.upem = f.unitsPerEm;
    this.key = face.id + (resolved.syntheticBold ? 'b' : '') + (resolved.syntheticItalic ? 'i' : '');
    this.family = face.family;
    this.ascent = f.lineAscent / this.upem;
    this.descent = f.lineDescent / this.upem;
    this.underlinePosition = -f.underlinePosition / this.upem;
    this.underlineThickness = Math.max(f.underlineThickness / this.upem, 0.04);
    this.syntheticBold = resolved.syntheticBold;
    this.syntheticItalic = resolved.syntheticItalic;
    this.bold = face.bold || resolved.syntheticBold;
    this.italic = face.italic || resolved.syntheticItalic;
  }

  advance(cp: number): number {
    let a = this.cache.get(cp);
    if (a === undefined) {
      const font = this.face.font;
      a = font.advance(font.glyphId(cp)) / this.upem;
      this.cache.set(cp, a);
    }
    return a;
  }

  glyphId(cp: number): number {
    return this.face.font.glyphId(cp);
  }

  has(cp: number): boolean {
    return this.face.font.hasGlyph(cp);
  }
}

export class RegistryFontProvider implements FontProvider {
  private readonly fonts = new Map<string, SfntLayoutFont>();

  constructor(private readonly registry: FontRegistry) {
    if (registry.isEmpty) throw new Error('docxkeun layout needs at least one registered font (FontRegistry.add).');
  }

  private wrap(resolved: ResolvedFace): SfntLayoutFont {
    const key = resolved.face.id + (resolved.syntheticBold ? 'b' : '') + (resolved.syntheticItalic ? 'i' : '');
    let font = this.fonts.get(key);
    if (!font) {
      font = new SfntLayoutFont(resolved);
      this.fonts.set(key, font);
    }
    return font;
  }

  get(family: string | undefined, bold: boolean, italic: boolean): LayoutFont {
    return this.wrap(this.registry.resolve(family, bold, italic)!);
  }

  forChar(font: LayoutFont, cp: number): LayoutFont {
    if (font.has(cp) || !font.face) return font;
    const resolved = this.registry.faceForChar({ face: font.face, syntheticBold: font.syntheticBold, syntheticItalic: font.syntheticItalic }, cp);
    return this.wrap(resolved);
  }
}
