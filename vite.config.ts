import { defineConfig } from 'vite';

export default defineConfig({
  root: 'playground',
  // The docxkeun-fonts files are served from the site root (/Carlito-Regular.ttf, ...).
  publicDir: '../packages/fonts',
  server: { port: 5199 },
});
