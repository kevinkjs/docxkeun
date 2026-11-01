/**
 * URLs of the bundled font files. Each URL is written as a literal `new URL('./file', import.meta.url)`,
 * so bundlers (Vite, webpack 5, Rollup, esbuild with asset loaders) copy the files into your build,
 * and in Node the URLs are file: URLs you can pass to fs.readFile.
 *
 * Use with docxkeun:
 *   import { FontRegistry, loadBundledFonts } from 'docxkeun';
 *   import { fontUrls } from 'docxkeun-fonts';
 *   await loadBundledFonts(registry, { urls: fontUrls });
 */
export const fontUrls = Object.freeze({
  'Carlito-Regular.ttf': new URL('./Carlito-Regular.ttf', import.meta.url),
  'Carlito-Bold.ttf': new URL('./Carlito-Bold.ttf', import.meta.url),
  'Carlito-Italic.ttf': new URL('./Carlito-Italic.ttf', import.meta.url),
  'Carlito-BoldItalic.ttf': new URL('./Carlito-BoldItalic.ttf', import.meta.url),
  'Caladea-Regular.ttf': new URL('./Caladea-Regular.ttf', import.meta.url),
  'Caladea-Bold.ttf': new URL('./Caladea-Bold.ttf', import.meta.url),
  'Caladea-Italic.ttf': new URL('./Caladea-Italic.ttf', import.meta.url),
  'Caladea-BoldItalic.ttf': new URL('./Caladea-BoldItalic.ttf', import.meta.url),
  'LiberationSans-Regular.ttf': new URL('./LiberationSans-Regular.ttf', import.meta.url),
  'LiberationSans-Bold.ttf': new URL('./LiberationSans-Bold.ttf', import.meta.url),
  'LiberationSans-Italic.ttf': new URL('./LiberationSans-Italic.ttf', import.meta.url),
  'LiberationSans-BoldItalic.ttf': new URL('./LiberationSans-BoldItalic.ttf', import.meta.url),
  'LiberationSerif-Regular.ttf': new URL('./LiberationSerif-Regular.ttf', import.meta.url),
  'LiberationSerif-Bold.ttf': new URL('./LiberationSerif-Bold.ttf', import.meta.url),
  'LiberationSerif-Italic.ttf': new URL('./LiberationSerif-Italic.ttf', import.meta.url),
  'LiberationSerif-BoldItalic.ttf': new URL('./LiberationSerif-BoldItalic.ttf', import.meta.url),
  'LiberationMono-Regular.ttf': new URL('./LiberationMono-Regular.ttf', import.meta.url),
  'LiberationMono-Bold.ttf': new URL('./LiberationMono-Bold.ttf', import.meta.url),
  'LiberationMono-Italic.ttf': new URL('./LiberationMono-Italic.ttf', import.meta.url),
  'LiberationMono-BoldItalic.ttf': new URL('./LiberationMono-BoldItalic.ttf', import.meta.url),
});
