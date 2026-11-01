/**
 * Test fonts: the open fonts of packages/fonts (docxkeun-fonts), plus (for a few parser
 * tests) fonts installed on this Mac, which are never committed and are skipped when missing.
 */
import { existsSync, readFileSync } from 'node:fs';
import { BUNDLED_FAMILIES, BUNDLED_FONTS, FontRegistry } from '../../src/fonts/index.js';

const SUPPLEMENTAL = '/System/Library/Fonts/Supplemental';

export const SYSTEM_FONTS = {
  arial: `${SUPPLEMENTAL}/Arial.ttf`,
  arialBold: `${SUPPLEMENTAL}/Arial Bold.ttf`,
  arialItalic: `${SUPPLEMENTAL}/Arial Italic.ttf`,
  times: `${SUPPLEMENTAL}/Times New Roman.ttf`,
  helveticaTtc: '/System/Library/Fonts/Helvetica.ttc',
};

export const haveSystemFonts = existsSync(SYSTEM_FONTS.arial) && existsSync(SYSTEM_FONTS.times);

export const FONTS_DIR = new URL('../../packages/fonts/', import.meta.url);

export function readFont(path: string | URL): Uint8Array {
  return new Uint8Array(readFileSync(path));
}

/** A registry with every bundled font (Carlito, Caladea, Liberation Sans/Serif/Mono). */
export function testRegistry(): FontRegistry {
  const registry = new FontRegistry();
  for (const f of BUNDLED_FONTS) registry.add(readFont(new URL(f.file, FONTS_DIR)), { family: f.family, bold: f.bold, italic: f.italic });
  registry.fallbacks = [...BUNDLED_FAMILIES];
  return registry;
}
