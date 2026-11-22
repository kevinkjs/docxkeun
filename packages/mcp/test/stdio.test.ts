/**
 * The built server (dist/cli.js) over stdio, the way Claude Code and Claude Desktop start it,
 * plus checks that the package metadata agrees with the rest of the repository.
 * The stdio test needs `pnpm build` and `pnpm -C packages/mcp build` first; CI runs them before the tests.
 */
import { spawn } from 'node:child_process';
import { existsSync, readFileSync } from 'node:fs';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { Client } from '@modelcontextprotocol/client';
import { StdioClientTransport } from '@modelcontextprotocol/client/stdio';
import { describe, expect, it } from 'vitest';
import { buildDocx, p } from '../../../test/helpers/docx.js';

const cli = fileURLToPath(new URL('../dist/cli.js', import.meta.url));
const rootDist = fileURLToPath(new URL('../../../dist/index.js', import.meta.url));
const built = existsSync(cli) && existsSync(rootDist);

const json = (path: string) => JSON.parse(readFileSync(new URL(path, import.meta.url), 'utf8'));

/** Whether a version satisfies a caret range such as ^0.1.3. */
function satisfiesCaret(version: string, range: string): boolean {
  const [ma, mi, pa] = range.replace(/^\^/, '').split('.').map(Number) as [number, number, number];
  const [va, vi, vp] = version.split('.').map(Number) as [number, number, number];
  if (va !== ma) return false;
  if (ma === 0) return vi === mi && vp >= pa;
  return vi > mi || (vi === mi && vp >= pa);
}

describe('package', () => {
  const pkg = json('../package.json');

  it('depends on the docxkeun and docxkeun-fonts in this repository', () => {
    expect(satisfiesCaret(json('../../../package.json').version, pkg.dependencies.docxkeun)).toBe(true);
    expect(satisfiesCaret(json('../../fonts/package.json').version, pkg.dependencies['docxkeun-fonts'])).toBe(true);
  });
});

describe.skipIf(!built)('stdio', () => {
  const pkg = json('../package.json');

  it('starts, lists tools and answers a call', async () => {
    const dir = await mkdtemp(join(tmpdir(), 'docxkeun-mcp-stdio-'));
    try {
      await writeFile(join(dir, 'hello.docx'), await buildDocx({ body: p('Hello from stdio') }));
      const transport = new StdioClientTransport({ command: process.execPath, args: [cli, '--role', 'viewer', dir], stderr: 'pipe' });
      const client = new Client({ name: 'stdio-test', version: '1.0.0' });
      await client.connect(transport);
      try {
        expect(client.getServerVersion()).toMatchObject({ name: 'docxkeun', version: pkg.version });
        const { tools } = await client.listTools();
        expect(tools.map((t) => t.name)).toContain('read_document');
        expect(tools.map((t) => t.name)).not.toContain('replace_text');
        const r = await client.callTool({ name: 'read_document', arguments: { path: 'hello.docx' } });
        expect((r.content as Array<{ text: string }>)[0]!.text).toContain('[1] Hello from stdio');
      } finally {
        await client.close();
      }
    } finally {
      await rm(dir, { recursive: true, force: true });
    }
  });
});

describe.skipIf(!built)('preview only', () => {
  it('serves the live preview without an MCP client until stopped', async () => {
    const dir = await mkdtemp(join(tmpdir(), 'docxkeun-mcp-watch-'));
    const child = spawn(process.execPath, [cli, '--preview-only', '--preview-port', '0', dir], { stdio: ['ignore', 'pipe', 'pipe'] });
    try {
      await writeFile(join(dir, 'watched.docx'), await buildDocx({ body: p('Watched') }));
      const url = await new Promise<string>((ok, fail) => {
        let out = '';
        child.stdout.on('data', (d: Buffer) => {
          out += d.toString();
          const m = /docxkeun live preview: (\S+)/.exec(out);
          if (m) ok(m[1]!);
        });
        child.once('exit', (code) => fail(new Error(`exited with ${code}: ${out}`)));
      });
      expect(await (await fetch(`${url}/api/documents`)).json()).toMatchObject([{ path: 'watched.docx' }]);
      const exited = new Promise<number | null>((ok) => child.once('exit', ok));
      child.kill('SIGTERM');
      expect(await exited).toBe(0);
    } finally {
      child.kill();
      await rm(dir, { recursive: true, force: true });
    }
  });
});
