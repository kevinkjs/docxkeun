import { fileURLToPath } from 'node:url';
import { defineConfig, type Plugin } from 'vite';
import { createRelay } from './server/relay.mjs';

const src = (path: string) => fileURLToPath(new URL(`../../src/${path}`, import.meta.url));

/**
 * A strict Content-Security-Policy for the built site: scripts only from the site itself, and network
 * access only to the site (Claude calls go through its relay). Not in dev, where Vite's HMR needs more.
 */
function contentSecurityPolicy(): Plugin {
  const policy = [
    "default-src 'self'",
    "script-src 'self'",
    "style-src 'self' 'unsafe-inline'",
    "img-src 'self' blob: data:",
    "font-src 'self' blob: data:",
    "connect-src 'self'",
    "object-src 'none'",
    "base-uri 'none'",
    "form-action 'none'",
  ].join('; ');
  return {
    name: 'docxkeun-chat-csp',
    apply: 'build',
    transformIndexHtml: (html) => html.replace('<head>', `<head>\n    <meta http-equiv="Content-Security-Policy" content="${policy}">`),
  };
}

/** The Claude API relay (server/relay.mjs) on the dev server, as `pnpm start:chat` has it. */
function claudeRelay(): Plugin {
  return {
    name: 'docxkeun-chat-relay',
    configureServer(server) {
      const relay = createRelay();
      server.middlewares.use((req, res, next) => void relay(req, res, next));
    },
  };
}

export default defineConfig({
  // Relative asset URLs, so the build works from any folder or sub-path.
  base: './',
  // The docxkeun-fonts files, served next to index.html.
  publicDir: '../fonts',
  resolve: {
    // docxkeun from this repository's sources (no build step needed).
    alias: [
      { find: /^docxkeun$/, replacement: src('index.ts') },
      { find: /^docxkeun\/editor$/, replacement: src('api/index.ts') },
      { find: /^docxkeun\/viewer$/, replacement: src('render/index.ts') },
      { find: /^docxkeun\/security$/, replacement: src('security/index.ts') },
      { find: /^docxkeun\/pdf$/, replacement: src('pdf/index.ts') },
      { find: /^docxkeun\/fonts$/, replacement: src('fonts/index.ts') },
    ],
  },
  plugins: [contentSecurityPolicy(), claudeRelay()],
  server: { port: 5260 },
  build: { target: 'es2022' },
});
