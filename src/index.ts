/**
 * docxkeun: read, render, edit and convert .docx files in the browser (and Node),
 * with no runtime dependencies.
 */
import { DocxDocument, type OpenOptions } from './model/document.js';

export { DocxDocument, DocxError, type OpenOptions, type CoreProperties, type StoryPart } from './model/document.js';
export type { Block, DocParagraph, DocTable, Inline, TextInline, ImageInline, MathInline, ShapeInfo, DrawingContent, DrawingChild, Story } from './model/content.js';
export { blocksToText } from './model/content.js';
export { FontRegistry, Font, METRIC_COMPATIBLE, registerEmbeddedFonts, loadBundledFonts, BUNDLED_FONTS, type FontFace, type BundledFamily } from './fonts/index.js';
export { layoutDocument, type LayoutOptions, type LayoutResult, type Page, type DisplayItem } from './layout/index.js';
export { readComments, type CommentInfo } from './comments/read.js';
export { readChart, describeChart, renderChart, type ChartModel } from './charts/index.js';
export { mathToText, layoutMath, type MathLayout, type MathLayoutOptions } from './math/index.js';
export { renderMetafile, metafileKind, metafileSize } from './metafile/index.js';
export { sniffImage, decodeRaster, type ImageKind } from './images/index.js';
export { imageInfo, type ImageInfo } from './images/size.js';

/** Open a .docx (or a password-protected one, when docxkeun/security is loaded). */
export function openDocx(input: Uint8Array | ArrayBuffer | Blob, options?: OpenOptions): Promise<DocxDocument> {
  return DocxDocument.open(input, options);
}

/** A new, empty document with Word's default styles. */
export function createDocx(): Promise<DocxDocument> {
  return DocxDocument.create();
}
