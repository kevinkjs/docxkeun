import { defineConfig } from 'vite';

export default defineConfig({
  root: 'playground',
  // The bundled open fonts are served from the site root (/Carlito-Regular.ttf, ...).
  publicDir: '../fonts',
  server: { port: 5199 },
});
