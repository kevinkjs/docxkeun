/**
 * Embeds fonts as Type0 / CIDFontType2 (TrueType) or CIDFontType0 via OpenType (CFF) with
 * Identity-H encoding: content streams use glyph ids, a ToUnicode map keeps text searchable.
 */
import type { FontFace } from '../fonts/registry.js';
import { subsetTrueType } from '../fonts/subset.js';
import { name, str, type PdfRef, type PdfValue } from './objects.js';
import type { PdfWriter } from './writer.js';

export class PdfFont {
  readonly glyphs = new Map<number, string>();

  constructor(readonly face: FontFace, readonly resource: string) {}

  /** Record glyphs used by a text item (glyph id → the text it represents). */
  use(glyph: number, text: string): void {
    if (!this.glyphs.has(glyph)) this.glyphs.set(glyph, text);
  }

  /** Advance of a glyph in PDF text space units (1/1000 em). */
  width(glyph: number): number {
    const f = this.face.font;
    return (f.advance(glyph) * 1000) / f.unitsPerEm;
  }

  async write(writer: PdfWriter, ref: PdfRef): Promise<void> {
    const font = this.face.font;
    if (!font.embeddingAllowed) {
      throw new Error(`The font "${this.face.family}" does not allow embedding (OS/2 fsType restricted license)`);
    }
    const scale = 1000 / font.unitsPerEm;
    const tag = subsetTag([...this.glyphs.keys()]);
    const baseName = `${tag}+${(font.postscriptName || this.face.family).replace(/[^\x21-\x7e]|[[\](){}<>/%#\s]/g, '')}`;
    const used = new Set(this.glyphs.keys());
    used.add(0);
    const program = subsetTrueType(font, used);
    const fontFile = font.isCFF
      ? await writer.stream({ Subtype: name('OpenType') }, program)
      : await writer.stream({ Length1: program.length }, program);

    let flags = font.isSymbolic ? 4 : 32;
    if (font.isFixedPitch) flags |= 1;
    if (font.isItalic) flags |= 64;
    const descriptor = writer.add({
      Type: name('FontDescriptor'),
      FontName: name(baseName),
      Flags: flags,
      FontBBox: font.bbox.map((v) => Math.round(v * scale)),
      ItalicAngle: font.italicAngle,
      Ascent: Math.round(font.lineAscent * scale),
      Descent: -Math.round(font.lineDescent * scale),
      CapHeight: Math.round(font.capHeight * scale),
      StemV: font.isBold ? 120 : 80,
      [font.isCFF ? 'FontFile3' : 'FontFile2']: fontFile,
    });

    // /W: runs of consecutive glyph ids share one array.
    const ids = [...used].sort((a, b) => a - b);
    const w: PdfValue[] = [];
    for (let i = 0; i < ids.length; ) {
      const start = ids[i]!;
      const run: number[] = [];
      while (i < ids.length && ids[i] === start + run.length) run.push(Math.round(this.width(ids[i++]!)));
      w.push(start, run);
    }

    const cidFont = writer.add({
      Type: name('Font'),
      Subtype: name(font.isCFF ? 'CIDFontType0' : 'CIDFontType2'),
      BaseFont: name(baseName),
      CIDSystemInfo: { Registry: str('Adobe'), Ordering: str('Identity'), Supplement: 0 },
      FontDescriptor: descriptor,
      W: w,
      DW: Math.round(this.width(0)),
      CIDToGIDMap: font.isCFF ? undefined : name('Identity'),
    });
    const toUnicode = await writer.stream({}, toUnicodeCMap(this.glyphs));
    writer.set(ref, {
      Type: name('Font'),
      Subtype: name('Type0'),
      BaseFont: name(baseName),
      Encoding: name('Identity-H'),
      DescendantFonts: [cidFont],
      ToUnicode: toUnicode,
    });
  }
}

function subsetTag(glyphs: number[]): string {
  let h = 2166136261;
  for (const g of glyphs) h = Math.imul(h ^ g, 16777619) >>> 0;
  let tag = '';
  for (let i = 0; i < 6; i++) {
    tag += String.fromCharCode(65 + (h % 26));
    h = Math.floor(h / 26) + i * 7919;
  }
  return tag;
}

function utf16Hex(text: string): string {
  let s = '';
  for (let i = 0; i < text.length; i++) s += text.charCodeAt(i).toString(16).padStart(4, '0');
  return s.toUpperCase();
}

export function toUnicodeCMap(glyphs: Map<number, string>): Uint8Array {
  const entries = [...glyphs].filter(([, t]) => t.length > 0).sort((a, b) => a[0] - b[0]);
  const lines = [
    '/CIDInit /ProcSet findresource begin',
    '12 dict begin',
    'begincmap',
    '/CIDSystemInfo << /Registry (Adobe) /Ordering (UCS) /Supplement 0 >> def',
    '/CMapName /Adobe-Identity-UCS def',
    '/CMapType 2 def',
    '1 begincodespacerange',
    '<0000> <FFFF>',
    'endcodespacerange',
  ];
  for (let i = 0; i < entries.length; i += 100) {
    const chunk = entries.slice(i, i + 100);
    lines.push(`${chunk.length} beginbfchar`);
    for (const [gid, text] of chunk) lines.push(`<${gid.toString(16).padStart(4, '0').toUpperCase()}> <${utf16Hex(text)}>`);
    lines.push('endbfchar');
  }
  lines.push('endcmap', 'CMapName currentdict /CMap defineresource pop', 'end', 'end');
  return new TextEncoder().encode(lines.join('\n'));
}
