#!/usr/bin/env node
/**
 * docxkeun-mcp: an MCP server (stdio) for Word documents.
 *
 *   docxkeun-mcp [--role editor] [--author "AI assistant"] [--fonts DIR] [--preview] [FOLDER ...]
 */
import { parseArgs } from 'node:util';
import { StdioServerTransport } from '@modelcontextprotocol/server/stdio';
import type { Role } from 'docxkeun/security';
import { parseColors, startPreview, type PreviewOptions } from './preview.js';
import { createServer, ROLES, VERSION } from './server.js';

const HELP = `docxkeun-mcp ${VERSION}: an MCP server that lets AI assistants read, edit, comment on,
protect and export Word documents (.docx).

Usage: docxkeun-mcp [options] [folder ...]

The server only reads and writes files inside the given folders (default: the current folder).

Options:
  --role <role>      What the AI may do (default: editor)
                       viewer     read and export PDFs
                       commenter  viewer + add, reply to and resolve comments
                       editor     commenter + change text, formatting, links and bookmarks
                       owner      editor + locks, document protection and passwords
  --author <name>    Name on comments the AI adds and on its caret in the live preview
                     (default: "AI assistant")
  --ai-color <color> Colour of the AI's caret and highlights in the live preview: a CSS colour,
                     or two separated by a comma for a gradient ("#0b57d0" or "#0b57d0,#34a853")
  --fonts <folder>   Extra .ttf/.otf fonts for page counts and PDF export (repeatable)
  --preview          Also serve a live preview at http://127.0.0.1:5230 (or the next free port)
                     that redraws documents whenever they change
  --preview-port <n> First port to try for the preview (default 5230)
  --preview-only     Run only the live preview (no MCP server) until stopped,
                     e.g. to watch a folder while you edit in Word
  -h, --help         Show this help
  -v, --version      Show the version

Claude Code:     claude mcp add docxkeun -- npx -y docxkeun-mcp ~/Documents
Claude Desktop:  {"mcpServers": {"docxkeun": {"command": "npx", "args": ["-y", "docxkeun-mcp", "/Users/you/Documents"]}}}
`;

async function main(): Promise<void> {
  const { values, positionals } = parseArgs({
    allowPositionals: true,
    options: {
      role: { type: 'string' },
      author: { type: 'string' },
      'ai-color': { type: 'string' },
      fonts: { type: 'string', multiple: true },
      preview: { type: 'boolean' },
      'preview-port': { type: 'string' },
      'preview-only': { type: 'boolean' },
      help: { type: 'boolean', short: 'h' },
      version: { type: 'boolean', short: 'v' },
    },
  });
  if (values.help) {
    process.stdout.write(HELP);
    return;
  }
  if (values.version) {
    process.stdout.write(`${VERSION}\n`);
    return;
  }
  const role = (values.role ?? 'editor') as Role;
  if (!ROLES.includes(role)) throw new Error(`Unknown role "${values.role}" (use ${ROLES.join(', ')})`);
  const folders = positionals.length ? positionals : [process.cwd()];
  let previewPort: number | undefined;
  if (values['preview-port'] !== undefined) {
    previewPort = Number(values['preview-port']);
    if (!Number.isInteger(previewPort) || previewPort < 0 || previewPort > 65535) throw new Error(`Not a port: ${values['preview-port']}`);
  }
  const color = values['ai-color'] === undefined ? undefined : parseColors(values['ai-color']);
  // The caret carries the same name as the AI's comments.
  const previewOptions: PreviewOptions = {
    ...(previewPort === undefined ? {} : { port: previewPort }),
    ai: { name: values.author?.trim() || 'AI assistant', ...(color ? { color } : {}) },
  };
  if (values['preview-only']) {
    const preview = await startPreview(folders, previewOptions);
    process.stdout.write(`docxkeun live preview: ${preview.url}\n`);
    const stop = () => void preview.close().finally(() => process.exit(0));
    process.once('SIGINT', stop);
    process.once('SIGTERM', stop);
    return;
  }
  const preview = values.preview || previewPort !== undefined ? await startPreview(folders, previewOptions) : undefined;
  const server = await createServer({
    folders,
    role,
    ...(values.author ? { author: values.author } : {}),
    ...(values.fonts ? { fontFolders: values.fonts } : {}),
    ...(preview ? { preview } : {}),
  });
  await server.connect(new StdioServerTransport());
  if (preview) {
    // The web server would keep the process alive after the client goes away: stop with stdin.
    const stop = () => void preview.close().finally(() => process.exit(0));
    process.stdin.once('end', stop);
    process.stdin.once('close', stop);
  }
  process.stderr.write(`docxkeun-mcp ${VERSION} running (role: ${role}; folders: ${folders.join(', ')}${preview ? `; preview: ${preview.url}` : ''})\n`);
}

main().catch((e: unknown) => {
  process.stderr.write(`docxkeun-mcp: ${e instanceof Error ? e.message : String(e)}\n`);
  process.exit(1);
});
