/**
 * Font registry: maps the font names a document asks for to font files the host registered,
 * with metric-compatible substitutes for the usual Office fonts.
 */
import { Font } from './sfnt.js';

export interface FontFace {
  /** Stable id used for PDF resources and CSS @font-face names. */
  id: string;
  family: string;
  bold: boolean;
  italic: boolean;
  font: Font;
}

export interface ResolvedFace {
  face: FontFace;
  /** The face is regular but bold was requested: embolden when drawing. */
  syntheticBold: boolean;
  /** The face is upright but italic was requested: slant when drawing. */
  syntheticItalic: boolean;
}

/** Open fonts designed with the same metrics as common proprietary ones. */
export const METRIC_COMPATIBLE: Record<string, string> = {
  calibri: 'Carlito',
  'calibri light': 'Carlito',
  cambria: 'Caladea',
  'cambria math': 'Caladea',
  arial: 'Liberation Sans',
  helvetica: 'Liberation Sans',
  'arial narrow': 'Liberation Sans Narrow',
  'times new roman': 'Liberation Serif',
  times: 'Liberation Serif',
  'courier new': 'Liberation Mono',
  courier: 'Liberation Mono',
};

export class FontRegistry {
  private readonly families = new Map<string, FontFace[]>();
  private readonly aliases = new Map<string, string>(Object.entries(METRIC_COMPATIBLE));
  private nextId = 1;
  /** Families tried, in order, when a requested family is unavailable. */
  fallbacks: string[] = [];

  /** Register a font file. Family, weight and style default to what the font declares. */
  add(data: Uint8Array, options: { family?: string; bold?: boolean; italic?: boolean } = {}): FontFace {
    const font = Font.parse(data);
    return this.addFont(font, options);
  }

  addFont(font: Font, options: { family?: string; bold?: boolean; italic?: boolean } = {}): FontFace {
    const family = options.family ?? font.familyName;
    const face: FontFace = {
      id: `F${this.nextId++}`,
      family,
      bold: options.bold ?? font.isBold,
      italic: options.italic ?? font.isItalic,
      font,
    };
    const key = family.toLowerCase();
    const list = this.families.get(key) ?? [];
    // Replace an existing face with the same style.
    const i = list.findIndex((f) => f.bold === face.bold && f.italic === face.italic);
    if (i >= 0) list[i] = face;
    else list.push(face);
    this.families.set(key, list);
    if (this.fallbacks.length === 0) this.fallbacks.push(family);
    return face;
  }

  /** Fetch and register a font file. */
  async addFromUrl(url: string, options: { family?: string; bold?: boolean; italic?: boolean } = {}): Promise<FontFace> {
    const response = await fetch(url);
    if (!response.ok) throw new Error(`Could not load font ${url}: ${response.status}`);
    return this.add(new Uint8Array(await response.arrayBuffer()), options);
  }

  /** Make `name` render with `family` (e.g. a corporate font with a substitute). */
  alias(name: string, family: string): void {
    this.aliases.set(name.toLowerCase(), family);
  }

  hasFamily(name: string): boolean {
    return this.families.has(name.toLowerCase());
  }

  faces(): FontFace[] {
    return [...this.families.values()].flat();
  }

  get isEmpty(): boolean {
    return this.families.size === 0;
  }

  private familyFaces(name: string | undefined): FontFace[] | undefined {
    if (name) {
      const direct = this.families.get(name.toLowerCase());
      if (direct) return direct;
      const alias = this.aliases.get(name.toLowerCase());
      if (alias) {
        const viaAlias = this.families.get(alias.toLowerCase());
        if (viaAlias) return viaAlias;
      }
    }
    for (const f of this.fallbacks) {
      const list = this.families.get(f.toLowerCase());
      if (list) return list;
    }
    return undefined;
  }

  /** Whether a family resolves without falling back. */
  isAvailable(name: string): boolean {
    const key = name.toLowerCase();
    const alias = this.aliases.get(key);
    return this.families.has(key) || (!!alias && this.families.has(alias.toLowerCase()));
  }

  resolve(name: string | undefined, bold: boolean, italic: boolean): ResolvedFace | undefined {
    const list = this.familyFaces(name);
    if (!list || list.length === 0) return undefined;
    const exact = list.find((f) => f.bold === bold && f.italic === italic);
    if (exact) return { face: exact, syntheticBold: false, syntheticItalic: false };
    const sameItalic = list.find((f) => f.italic === italic);
    const sameBold = list.find((f) => f.bold === bold);
    const regular = list.find((f) => !f.bold && !f.italic);
    const face = sameItalic ?? sameBold ?? regular ?? list[0]!;
    return { face, syntheticBold: bold && !face.bold, syntheticItalic: italic && !face.italic };
  }

  /** A face that has a glyph for the code point: the preferred one, then any registered face. */
  faceForChar(preferred: ResolvedFace, codePoint: number): ResolvedFace {
    if (preferred.face.font.hasGlyph(codePoint)) return preferred;
    for (const f of this.fallbacks) {
      const r = this.resolve(f, preferred.face.bold || preferred.syntheticBold, preferred.face.italic || preferred.syntheticItalic);
      if (r && r.face.font.hasGlyph(codePoint)) return r;
    }
    for (const face of this.faces()) {
      if (face.font.hasGlyph(codePoint)) return { face, syntheticBold: false, syntheticItalic: false };
    }
    return preferred;
  }
}
