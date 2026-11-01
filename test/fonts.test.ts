import { describe, expect, it } from 'vitest';
import { existsSync } from 'node:fs';
import { BUNDLED_FONTS, deobfuscateFont, Font, FontRegistry, loadBundledFonts, obfuscateFont, subsetTrueType } from '../src/fonts/index.js';
import { bytesEqual } from '../src/util/bytes.js';
import { FONTS_DIR, SYSTEM_FONTS, haveSystemFonts, readFont } from './helpers/fonts.js';

function sfntChecksum(bytes: Uint8Array): number {
  let sum = 0;
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength - (bytes.byteLength % 4));
  for (let i = 0; i + 4 <= bytes.length; i += 4) sum = (sum + view.getUint32(i)) >>> 0;
  return sum;
}

describe.skipIf(!haveSystemFonts)('TrueType parser', () => {
  it('reads Arial metrics', () => {
    const font = Font.parse(readFont(SYSTEM_FONTS.arial));
    expect(font.familyName).toBe('Arial');
    expect(font.unitsPerEm).toBe(2048);
    expect(font.winAscent).toBe(1854);
    expect(font.winDescent).toBe(434);
    expect(font.advance(font.glyphId(0x41))).toBe(1366); // 'A'
    expect(font.advance(font.glyphId(0x20))).toBe(569); // space
    expect(font.hasGlyph(0x20ac)).toBe(true); // €
    expect(font.isBold).toBe(false);
    expect(font.embeddingAllowed).toBe(true);
    expect(font.postscriptName).toBe('ArialMT');
  });

  it('reads bold and italic flags', () => {
    expect(Font.parse(readFont(SYSTEM_FONTS.arialBold)).isBold).toBe(true);
    expect(Font.parse(readFont(SYSTEM_FONTS.arialItalic)).isItalic).toBe(true);
  });

  it('reads faces of a .ttc collection', () => {
    const faces = Font.fromCollection(readFont(SYSTEM_FONTS.helveticaTtc));
    expect(faces.length).toBeGreaterThan(1);
    expect(faces[0]!.familyName).toMatch(/Helvetica/);
  });

  it('subsets while keeping glyph ids and a valid checksum', () => {
    const font = Font.parse(readFont(SYSTEM_FONTS.arial));
    const used = new Set([...'Hello Äö'].map((c) => font.glyphId(c.codePointAt(0)!)));
    const subset = subsetTrueType(font, used);
    expect(subset.length).toBeLessThan(font.data.length / 4);
    expect(sfntChecksum(subset)).toBe(0xb1b0afba);
    const parsed = new Font(subset);
    expect(parsed.numGlyphs).toBe(font.numGlyphs);
    for (const gid of used) {
      const a = font.glyphRange(gid)!;
      const b = parsed.glyphRange(gid)!;
      // Glyphs are padded to 4 bytes in the subset.
      expect(b.length - a.length).toBeLessThan(4);
      expect(bytesEqual(font.data.subarray(a.offset, a.offset + a.length), parsed.data.subarray(b.offset, b.offset + a.length))).toBe(true);
      expect(parsed.advance(gid)).toBe(font.advance(gid));
    }
    expect(parsed.glyphRange(font.glyphId(0x5a))!.length).toBe(0); // 'Z' not kept
  });
});

describe('bundled fonts', () => {
  it('load through loadBundledFonts and stand in for Office fonts', async () => {
    const registry = new FontRegistry();
    const faces = await loadBundledFonts(registry, { load: async (file) => readFont(new URL(file, FONTS_DIR)) });
    expect(faces).toHaveLength(20);
    expect(faces.every((f) => f.font.embeddingAllowed)).toBe(true);
    expect(registry.resolve('Calibri', false, false)!.face.family).toBe('Carlito');
    expect(registry.resolve('Calibri Light', true, true)!.face).toMatchObject({ family: 'Carlito', bold: true, italic: true });
    expect(registry.resolve('Cambria', false, true)!.face.family).toBe('Caladea');
    expect(registry.resolve('Times New Roman', true, false)!.face.family).toBe('Liberation Serif');
    expect(registry.resolve('Courier New', false, false)!.face.family).toBe('Liberation Mono');
    expect(registry.resolve('Some Corporate Font', false, false)!.face.family).toBe('Carlito');
  });

  it('are all listed by the docxkeun-fonts package, as files that exist', async () => {
    const { fontUrls } = (await import('../packages/fonts/index.js')) as { fontUrls: Record<string, URL> };
    expect(Object.keys(fontUrls).sort()).toEqual(BUNDLED_FONTS.map((f) => f.file).sort());
    for (const url of Object.values(fontUrls)) expect(existsSync(url), url.href).toBe(true);
    // In Node the URLs are file: URLs, read with fs (fetch cannot read file: URLs).
    const loaded = await loadBundledFonts(new FontRegistry(), { load: async (file) => readFont(fontUrls[file]!), families: ['Caladea'] });
    expect(loaded.map((f) => f.family)).toEqual(['Caladea', 'Caladea', 'Caladea', 'Caladea']);
  });

  it.skipIf(!haveSystemFonts)('Liberation Sans and Serif match Arial and Times New Roman widths', () => {
    const pairs: Array<[string, string]> = [
      ['LiberationSans-Regular.ttf', SYSTEM_FONTS.arial],
      ['LiberationSerif-Regular.ttf', SYSTEM_FONTS.times],
    ];
    for (const [bundled, system] of pairs) {
      const a = Font.parse(readFont(new URL(bundled, FONTS_DIR)));
      const b = Font.parse(readFont(system));
      expect(a.unitsPerEm).toBe(b.unitsPerEm);
      expect([a.winAscent, a.winDescent]).toEqual([b.winAscent, b.winDescent]);
      for (let cp = 0x20; cp < 0x7f; cp++) expect(a.advance(a.glyphId(cp)), String.fromCharCode(cp)).toBe(b.advance(b.glyphId(cp)));
    }
  });
});

describe('embedded font obfuscation', () => {
  it('round-trips and only touches the first 32 bytes', () => {
    const data = Uint8Array.from({ length: 100 }, (_, i) => i);
    const key = '{6C8E3A1F-0B2C-4D5E-8F90-A1B2C3D4E5F6}';
    const hidden = obfuscateFont(data, key);
    expect(bytesEqual(hidden.subarray(32), data.subarray(32))).toBe(true);
    expect(bytesEqual(hidden.subarray(0, 32), data.subarray(0, 32))).toBe(false);
    // Key bytes are the GUID read backwards: first byte XORs with 0xF6.
    expect(hidden[0]).toBe(0 ^ 0xf6);
    expect(hidden[15]).toBe(15 ^ 0x6c);
    expect(bytesEqual(deobfuscateFont(hidden, key), data)).toBe(true);
  });
});
