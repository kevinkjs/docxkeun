/**
 * Fonts for page layout and PDF export: the bundled open fonts (docxkeun-fonts), the document's own
 * embedded fonts, and any font folders given on the command line (for example Microsoft Office's).
 */
import { readdir, readFile } from 'node:fs/promises';
import { extname, join } from 'node:path';
import { Font, FontRegistry, loadBundledFonts, registerEmbeddedFonts, type DocxDocument } from 'docxkeun';
import { fontUrls } from 'docxkeun-fonts';

const bundled = new Map<string, Promise<Uint8Array>>();

function bundledFile(file: string): Promise<Uint8Array> {
  let bytes = bundled.get(file);
  if (!bytes) {
    const url = fontUrls[file];
    if (!url) throw new Error(`docxkeun-fonts has no ${file}`);
    bytes = readFile(url).then((b) => new Uint8Array(b));
    bundled.set(file, bytes);
  }
  return bytes;
}

const extraFonts = new Map<string, Promise<Font[]>>();

async function readFontFolders(folders: string[]): Promise<Font[]> {
  const out: Font[] = [];
  for (const folder of folders) {
    let names: string[];
    try {
      names = await readdir(folder);
    } catch {
      process.stderr.write(`docxkeun-mcp: cannot read font folder ${folder}\n`);
      continue;
    }
    for (const name of names.sort()) {
      if (!['.ttf', '.otf'].includes(extname(name).toLowerCase())) continue;
      const file = join(folder, name);
      try {
        out.push(Font.parse(new Uint8Array(await readFile(file))));
      } catch {
        // Fonts docxkeun cannot parse (e.g. font collections or some CFF-based .otf files) are skipped.
        process.stderr.write(`docxkeun-mcp: skipped font ${file}\n`);
      }
    }
  }
  return out;
}

/** A font registry for one document. */
export async function fontsFor(doc: DocxDocument, fontFolders: string[] = []): Promise<FontRegistry> {
  const registry = new FontRegistry();
  await loadBundledFonts(registry, { load: bundledFile });
  if (fontFolders.length) {
    const key = fontFolders.join('\0');
    let fonts = extraFonts.get(key);
    if (!fonts) {
      fonts = readFontFolders(fontFolders);
      extraFonts.set(key, fonts);
    }
    for (const font of await fonts) registry.addFont(font);
  }
  await registerEmbeddedFonts(doc, registry);
  return registry;
}
