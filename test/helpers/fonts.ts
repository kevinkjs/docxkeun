/**
 * Test fonts. Tests use fonts installed on the machine (never committed to the repo) and are
 * skipped where they are not available.
 */
import { existsSync, readFileSync } from 'node:fs';
import { FontRegistry } from '../../src/fonts/index.js';

const SUPPLEMENTAL = '/System/Library/Fonts/Supplemental';

export const SYSTEM_FONTS = {
  arial: `${SUPPLEMENTAL}/Arial.ttf`,
  arialBold: `${SUPPLEMENTAL}/Arial Bold.ttf`,
  arialItalic: `${SUPPLEMENTAL}/Arial Italic.ttf`,
  arialBoldItalic: `${SUPPLEMENTAL}/Arial Bold Italic.ttf`,
  times: `${SUPPLEMENTAL}/Times New Roman.ttf`,
  timesBold: `${SUPPLEMENTAL}/Times New Roman Bold.ttf`,
  courier: `${SUPPLEMENTAL}/Courier New.ttf`,
  helveticaTtc: '/System/Library/Fonts/Helvetica.ttc',
};

export const haveSystemFonts = existsSync(SYSTEM_FONTS.arial) && existsSync(SYSTEM_FONTS.times);

export function readFont(path: string): Uint8Array {
  return new Uint8Array(readFileSync(path));
}

/** A registry with Arial and Times New Roman standing in for every font. */
export function testRegistry(): FontRegistry {
  const registry = new FontRegistry();
  registry.add(readFont(SYSTEM_FONTS.arial), { family: 'Arial' });
  for (const [key, bold, italic] of [['arialBold', true, false], ['arialItalic', false, true], ['arialBoldItalic', true, true]] as const) {
    if (existsSync(SYSTEM_FONTS[key])) registry.add(readFont(SYSTEM_FONTS[key]), { family: 'Arial', bold, italic });
  }
  registry.add(readFont(SYSTEM_FONTS.times), { family: 'Times New Roman' });
  if (existsSync(SYSTEM_FONTS.timesBold)) registry.add(readFont(SYSTEM_FONTS.timesBold), { family: 'Times New Roman', bold: true });
  if (existsSync(SYSTEM_FONTS.courier)) registry.add(readFont(SYSTEM_FONTS.courier), { family: 'Courier New' });
  // Office's default fonts fall back to Arial in tests.
  registry.alias('Calibri', 'Arial');
  registry.alias('Calibri Light', 'Arial');
  registry.alias('Cambria', 'Times New Roman');
  registry.fallbacks = ['Arial'];
  return registry;
}
