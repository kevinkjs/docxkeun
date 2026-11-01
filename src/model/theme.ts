/** DrawingML theme: font scheme and colour scheme (word/theme/theme1.xml). */
import { attr, elements, firstChild, type XmlDocument, type XmlElement } from '../xml/index.js';

export interface ThemeFonts {
  latin?: string;
  eastAsia?: string;
  cs?: string;
}

export class Theme {
  major: ThemeFonts = { latin: 'Calibri Light' };
  minor: ThemeFonts = { latin: 'Calibri' };
  readonly colors = new Map<string, string>([
    ['dk1', '000000'],
    ['lt1', 'FFFFFF'],
    ['dk2', '44546A'],
    ['lt2', 'E7E6E6'],
    ['accent1', '4472C4'],
    ['accent2', 'ED7D31'],
    ['accent3', 'A5A5A5'],
    ['accent4', 'FFC000'],
    ['accent5', '5B9BD5'],
    ['accent6', '70AD47'],
    ['hlink', '0563C1'],
    ['folHlink', '954F72'],
  ]);

  static parse(doc: XmlDocument | undefined): Theme {
    const theme = new Theme();
    if (!doc) return theme;
    const elementsNode = firstChild(doc.root, 'a:themeElements');
    const fontScheme = firstChild(elementsNode, 'a:fontScheme');
    const readFonts = (node: XmlElement | undefined): ThemeFonts | undefined => {
      if (!node) return undefined;
      return {
        latin: attr(firstChild(node, 'a:latin'), 'typeface') || undefined,
        eastAsia: attr(firstChild(node, 'a:ea'), 'typeface') || undefined,
        cs: attr(firstChild(node, 'a:cs'), 'typeface') || undefined,
      };
    };
    theme.major = readFonts(firstChild(fontScheme, 'a:majorFont')) ?? theme.major;
    theme.minor = readFonts(firstChild(fontScheme, 'a:minorFont')) ?? theme.minor;
    const clrScheme = firstChild(elementsNode, 'a:clrScheme');
    if (clrScheme) {
      for (const c of elements(clrScheme)) {
        const key = c.name.slice(2);
        const srgb = attr(firstChild(c, 'a:srgbClr'), 'val');
        const sys = firstChild(c, 'a:sysClr');
        const value = srgb ?? attr(sys, 'lastClr');
        if (value) theme.colors.set(key, value.toUpperCase());
      }
    }
    return theme;
  }

  /** Resolve `w:asciiTheme` style values such as `minorHAnsi` to a typeface. */
  font(themeRef: string | undefined): string | undefined {
    if (!themeRef) return undefined;
    const scheme = themeRef.startsWith('major') ? this.major : this.minor;
    if (themeRef.endsWith('EastAsia')) return scheme.eastAsia || scheme.latin;
    if (themeRef.endsWith('Bidi')) return scheme.cs || scheme.latin;
    return scheme.latin;
  }

  /** Resolve a `w:themeColor` value (dark1, accent1, text1, hyperlink...) to RRGGBB. */
  color(name: string | undefined, tint?: string, shade?: string): string | undefined {
    if (!name) return undefined;
    const key =
      {
        dark1: 'dk1',
        light1: 'lt1',
        dark2: 'dk2',
        light2: 'lt2',
        text1: 'dk1',
        background1: 'lt1',
        text2: 'dk2',
        background2: 'lt2',
        hyperlink: 'hlink',
        followedHyperlink: 'folHlink',
      }[name] ?? name;
    const base = this.colors.get(key);
    if (!base) return undefined;
    return applyTintShade(base, tint, shade);
  }
}

/** w:themeTint / w:themeShade are hex bytes (0..FF) applied in RGB space, as Word does. */
export function applyTintShade(hex: string, tint?: string, shade?: string): string {
  let [r, g, b] = [0, 2, 4].map((i) => Number.parseInt(hex.slice(i, i + 2), 16)) as [number, number, number];
  if (shade) {
    const f = Number.parseInt(shade, 16) / 255;
    [r, g, b] = [r * f, g * f, b * f];
  }
  if (tint) {
    const f = Number.parseInt(tint, 16) / 255;
    [r, g, b] = [r * f + 255 * (1 - f), g * f + 255 * (1 - f), b * f + 255 * (1 - f)];
  }
  return [r, g, b].map((v) => Math.round(Math.max(0, Math.min(255, v))).toString(16).padStart(2, '0')).join('').toUpperCase();
}

/** Theme-aware colour of a w:color element (w:val / w:themeColor / w:themeTint / w:themeShade). */
export function resolveColorElement(theme: Theme, node: XmlElement | undefined): string | undefined {
  if (!node) return undefined;
  const themed = theme.color(attr(node, 'w:themeColor'), attr(node, 'w:themeTint'), attr(node, 'w:themeShade'));
  return themed ?? attr(node, 'w:val');
}

