export { Font, FontParseError, type TableRecord } from './sfnt.js';
export { subsetTrueType, closeOverComposites, standaloneFont, buildSfnt } from './subset.js';
export { deobfuscateFont, obfuscateFont } from './odttf.js';
export { FontRegistry, METRIC_COMPATIBLE, type FontFace, type ResolvedFace } from './registry.js';
export { registerEmbeddedFonts } from './embedded.js';
