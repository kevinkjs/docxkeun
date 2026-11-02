import { defineConfig } from 'vite';

export default defineConfig({
  root: 'playground',
  // Relative asset URLs, so the built demo works under https://kevinkjs.github.io/docxkeun/.
  base: './',
  // The docxkeun-fonts files are served next to index.html (./Carlito-Regular.ttf, ...).
  publicDir: '../packages/fonts',
  server: { port: 5199 },
});
