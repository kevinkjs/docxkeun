import { existsSync, readFileSync } from 'node:fs';
import { defineConfig, type Plugin } from 'vite';

/**
 * Dev-only: serve fonts installed on this machine to the playground, so layout can use real
 * font files without committing any. Requests look like /__fonts/Arial.ttf.
 */
function localFonts(): Plugin {
  const dirs = ['/System/Library/Fonts/Supplemental', '/System/Library/Fonts', '/Library/Fonts'];
  return {
    name: 'docxkit-local-fonts',
    configureServer(server) {
      server.middlewares.use('/__fonts/', (req, res, next) => {
        const name = decodeURIComponent((req.url ?? '').replace(/^\//, '').split('?')[0] ?? '');
        if (!/^[\w .-]+\.(ttf|otf|ttc)$/i.test(name)) return next();
        const dir = dirs.find((d) => existsSync(`${d}/${name}`));
        if (!dir) {
          res.statusCode = 404;
          res.end();
          return;
        }
        res.setHeader('Content-Type', 'font/ttf');
        res.end(readFileSync(`${dir}/${name}`));
      });
    },
  };
}

export default defineConfig({
  root: 'playground',
  plugins: [localFonts()],
  server: { port: 5199 },
});
