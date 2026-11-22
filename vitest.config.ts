import { fileURLToPath } from 'node:url';
import { defineConfig } from 'vitest/config';

const src = (path: string) => fileURLToPath(new URL(`./src/${path}`, import.meta.url));

export default defineConfig({
  resolve: {
    // packages/mcp imports docxkeun by name; test it against the sources, not a build.
    alias: [
      { find: /^docxkeun$/, replacement: src('index.ts') },
      { find: /^docxkeun\/editor$/, replacement: src('api/index.ts') },
      { find: /^docxkeun\/security$/, replacement: src('security/index.ts') },
      { find: /^docxkeun\/pdf$/, replacement: src('pdf/index.ts') },
      { find: /^docxkeun\/fonts$/, replacement: src('fonts/index.ts') },
      { find: /^docxkeun\/layout$/, replacement: src('layout/index.ts') },
    ],
  },
  test: {
    include: ['test/**/*.test.ts', 'packages/*/test/**/*.test.ts'],
    environment: 'node',
    testTimeout: 30_000,
  },
});
