/** Fonts embedded in a document (word/fonts/*.odttf, listed in fontTable.xml). */
import type { DocxDocument } from '../model/document.js';
import { attr, elements } from '../xml/index.js';
import { deobfuscateFont } from './odttf.js';
import type { FontFace, FontRegistry } from './registry.js';

const VARIANTS: Array<[string, boolean, boolean]> = [
  ['w:embedRegular', false, false],
  ['w:embedBold', true, false],
  ['w:embedItalic', false, true],
  ['w:embedBoldItalic', true, true],
];

/** Register every embedded font of a document; returns the faces that were added. */
export async function registerEmbeddedFonts(doc: DocxDocument, registry: FontRegistry): Promise<FontFace[]> {
  const part = doc.parts.fontTable;
  if (!part) return [];
  const xml = doc.xmlSync(part) ?? (await doc.pkg.getXml(part));
  const rels = await doc.pkg.relationships(part);
  const added: FontFace[] = [];
  for (const fontEl of elements(xml.root, 'w:font')) {
    const family = attr(fontEl, 'w:name');
    if (!family) continue;
    for (const [tag, bold, italic] of VARIANTS) {
      const embed = elements(fontEl, tag)[0];
      const rId = attr(embed, 'r:id');
      if (!embed || !rId) continue;
      const rel = rels.find((r) => r.id === rId);
      if (!rel || rel.external) continue;
      try {
        let data = await doc.pkg.getBytes(doc.pkg.resolve(part, rel));
        const key = attr(embed, 'w:fontKey');
        if (key) data = deobfuscateFont(data, key);
        added.push(registry.add(data, { family, bold, italic }));
      } catch {
        // A damaged embedded font should not stop the document from opening.
      }
    }
  }
  return added;
}
