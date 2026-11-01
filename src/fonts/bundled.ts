/**
 * The open, metric-compatible fonts shipped in docxkit's `fonts/` folder
 * (Carlito, Caladea, Liberation Sans/Serif/Mono; SIL OFL 1.1).
 */
import type { FontFace, FontRegistry } from './registry.js';

export type BundledFamily = 'Carlito' | 'Caladea' | 'Liberation Sans' | 'Liberation Serif' | 'Liberation Mono';

const FILE_PREFIX: Record<BundledFamily, string> = {
  Carlito: 'Carlito',
  Caladea: 'Caladea',
  'Liberation Sans': 'LiberationSans',
  'Liberation Serif': 'LiberationSerif',
  'Liberation Mono': 'LiberationMono',
};

const STYLES: Array<[string, boolean, boolean]> = [
  ['Regular', false, false],
  ['Bold', true, false],
  ['Italic', false, true],
  ['BoldItalic', true, true],
];

export const BUNDLED_FAMILIES = Object.keys(FILE_PREFIX) as BundledFamily[];

/** Every bundled font file with the family and style it provides. */
export const BUNDLED_FONTS: Array<{ file: string; family: BundledFamily; bold: boolean; italic: boolean }> = BUNDLED_FAMILIES.flatMap((family) =>
  STYLES.map(([style, bold, italic]) => ({ file: `${FILE_PREFIX[family]}-${style}.ttf`, family, bold, italic })),
);

export interface LoadBundledFontsOptions {
  /**
   * URL of the `fonts/` folder, ending in "/". In a bundler, for example
   * `new URL('docxkit/fonts/', import.meta.url)`; or wherever you host the files.
   */
  baseUrl?: string | URL;
  /** Custom loader instead of fetch, e.g. `(file) => readFile(join(fontsDir, file))` in Node. */
  load?: (file: string) => Promise<Uint8Array | ArrayBuffer>;
  /** Families to load (default: all five). */
  families?: BundledFamily[];
}

/** Load the bundled fonts into a registry. Carlito (Calibri's stand-in) becomes the fallback. */
export async function loadBundledFonts(registry: FontRegistry, options: LoadBundledFontsOptions = {}): Promise<FontFace[]> {
  const wanted = new Set(options.families ?? BUNDLED_FAMILIES);
  const load =
    options.load ??
    (async (file: string) => {
      if (options.baseUrl === undefined) throw new Error('loadBundledFonts needs a baseUrl or a load function');
      const url = new URL(file, typeof options.baseUrl === 'string' ? new URL(options.baseUrl, globalThis.location?.href ?? 'file:///') : options.baseUrl);
      const response = await fetch(url);
      if (!response.ok) throw new Error(`Could not load ${url}: ${response.status}`);
      return response.arrayBuffer();
    });
  const files = BUNDLED_FONTS.filter((f) => wanted.has(f.family));
  const data = await Promise.all(files.map(async (f) => load(f.file)));
  const faces = files.map((f, i) => {
    const bytes = data[i]!;
    return registry.add(bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes), { family: f.family, bold: f.bold, italic: f.italic });
  });
  const preferred = BUNDLED_FAMILIES.filter((f) => wanted.has(f));
  registry.fallbacks = [...new Set([...preferred, ...registry.fallbacks])];
  return faces;
}
