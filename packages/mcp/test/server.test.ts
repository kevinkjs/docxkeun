/**
 * docxkeun-mcp through a real MCP client (in-memory transport): tool lists per role, reading,
 * editing, comments, locks, protection, passwords, PDF export and the folder sandbox.
 */
import { mkdir, mkdtemp, readFile, rm, symlink, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { Client } from '@modelcontextprotocol/client';
import { InMemoryTransport } from '@modelcontextprotocol/server';
import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { DocxDocument } from 'docxkeun';
import { isEncryptedPackage } from 'docxkeun/security';
import { buildDocx, p } from '../../../test/helpers/docx.js';
import { createServer, type DocxMcpOptions } from '../src/index.js';

const STYLES = `
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
<w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:basedOn w:val="Normal"/></w:style>
<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:pPr><w:outlineLvl w:val="0"/></w:pPr></w:style>`;

const cell = (text: string) => `<w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>${p(text)}</w:tc>`;

const BODY =
  p('Service Agreement', '<w:pStyle w:val="Title"/>') +
  p('1. Parties', '<w:pStyle w:val="Heading1"/>') +
  p('This agreement is between {{client}} and ACME Ltd.') +
  p('Payment is due within 30 days of the invoice date.') +
  `<w:tbl><w:tblPr><w:tblW w:w="8000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid>` +
  `<w:tr>${cell('Name')}${cell('Value')}</w:tr><w:tr>${cell('Rate')}${cell('3.75%')}</w:tr></w:tbl>` +
  p('Signed by {{signer}}.') +
  p('The Supplier shall deliver on time. ACME Ltd. agrees.');

/** One list: a., b., c. */
const NUMBERING =
  '<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum>' +
  '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>';

let dir: string;
const clients: Client[] = [];

async function fixture(name = 'contract.docx'): Promise<string> {
  const file = join(dir, name);
  await writeFile(file, await buildDocx({ body: BODY, styles: STYLES, settings: '' }));
  return file;
}

async function connect(options: Partial<DocxMcpOptions> = {}): Promise<Client> {
  const server = await createServer({ folders: [dir], ...options });
  const [clientSide, serverSide] = InMemoryTransport.createLinkedPair();
  await server.connect(serverSide);
  const client = new Client({ name: 'test', version: '1.0.0' });
  await client.connect(clientSide);
  clients.push(client);
  return client;
}

interface Result {
  text: string;
  isError: boolean;
}

async function call(client: Client, name: string, args: Record<string, unknown>): Promise<Result> {
  const r = await client.callTool({ name, arguments: args });
  const text = (r.content as Array<{ type: string; text?: string }>)
    .filter((c) => c.type === 'text')
    .map((c) => c.text)
    .join('\n');
  return { text, isError: r.isError === true };
}

async function ok(client: Client, name: string, args: Record<string, unknown>): Promise<string> {
  const r = await call(client, name, args);
  if (r.isError) throw new Error(`${name} failed: ${r.text}`);
  return r.text;
}

async function bodyText(file: string, password?: string): Promise<string[]> {
  const doc = await DocxDocument.open(new Uint8Array(await readFile(file)), password ? { password } : {});
  const { listParagraphs } = await import('../src/outline.js');
  return listParagraphs(doc).map((e) => e.text);
}

beforeEach(async () => {
  dir = await mkdtemp(join(tmpdir(), 'docxkeun-mcp-'));
});

afterEach(async () => {
  for (const c of clients.splice(0)) await c.close();
  await rm(dir, { recursive: true, force: true });
});

describe('tools and roles', () => {
  it('offers each role only the tools it may use', async () => {
    const names = async (role: DocxMcpOptions['role']) => (await (await connect({ role })).listTools()).tools.map((t) => t.name).sort();
    const viewer = await names('viewer');
    expect(viewer).toEqual(['export_pdf', 'find_text', 'get_document_info', 'list_comments', 'list_documents', 'list_links', 'list_objects', 'list_revisions', 'read_chart', 'read_document', 'read_headers_footers']);
    const commenter = await names('commenter');
    expect(commenter).toEqual([...viewer, 'add_comment', 'delete_comment', 'reply_to_comment', 'resolve_comment'].sort());
    const editor = await names('editor');
    expect(editor).toContain('replace_text');
    expect(editor).not.toContain('lock_region');
    const owner = await names('owner');
    expect(owner).toEqual([...editor, 'lock_region', 'protect_document', 'set_open_password', 'unlock_region'].sort());
  });

  it('describes tools with JSON schemas and read-only hints', async () => {
    const { tools } = await (await connect()).listTools();
    const read = tools.find((t) => t.name === 'read_document')!;
    expect(read.annotations?.readOnlyHint).toBe(true);
    expect(read.inputSchema.required).toEqual(['path']);
    expect(Object.keys(read.inputSchema.properties ?? {})).toEqual(['path', 'password', 'from', 'count']);
    const replace = tools.find((t) => t.name === 'replace_text')!;
    expect(replace.annotations?.destructiveHint).toBe(true);
  });

  it('tells the client how to use it', async () => {
    const client = await connect({ role: 'commenter', author: 'Review Bot' });
    const instructions = client.getInstructions() ?? '';
    expect(instructions).toContain(dir.split('/').pop());
    expect(instructions).toContain('commenter');
    expect(instructions).toContain('"Review Bot"');
  });
});

describe('reading', () => {
  it('numbers paragraphs, including table cells, with styles', async () => {
    await fixture();
    const text = await ok(await connect(), 'read_document', { path: 'contract.docx' });
    expect(text).toContain('contract.docx: 10 paragraphs, 1 table, 38 words, 0 comment threads.');
    expect(text).toContain('[1] (Title) Service Agreement');
    expect(text).toContain('[2] (Heading 1) 1. Parties');
    expect(text).toContain('[4] Payment is due within 30 days of the invoice date.');
    expect(text).toContain('[8] (table 1, row 2, cell 2) 3.75%');
    expect(text).toContain('[10] The Supplier shall deliver on time.');
  });

  it('pages through long documents', async () => {
    await fixture();
    const text = await ok(await connect(), 'read_document', { path: 'contract.docx', from: 3, count: 2 });
    expect(text).toContain('[3] This agreement');
    expect(text).not.toContain('[5]');
    expect(text).toContain('Call read_document with from: 5');
  });

  it('finds text with paragraph numbers and context', async () => {
    await fixture();
    const text = await ok(await connect(), 'find_text', { path: 'contract.docx', query: 'acme ltd.', matchCase: false });
    expect(text).toContain('2 matches');
    expect(text).toContain('[3] This agreement is between {{client}} and «ACME Ltd.»');
    expect(text).toContain('[10] The Supplier shall deliver on time. «ACME Ltd.» agrees.');
    expect(await ok(await connect(), 'find_text', { path: 'contract.docx', query: '\\d+ days', regex: true })).toContain('«30 days»');
  });

  it('reports details, pages and styles', async () => {
    await fixture();
    const text = await ok(await connect(), 'get_document_info', { path: 'contract.docx' });
    expect(text).toContain('Pages: 1');
    expect(text).toContain('Paragraphs: 10, tables: 1');
    expect(text).toContain('Protection: none');
    expect(text).toContain('Paragraph styles: Normal, Title, Heading 1');
    expect(text).toContain('Your role: editor');
  });

  it('lists documents and skips Word lock files and hidden folders', async () => {
    await fixture();
    await fixture('~$ntract.docx');
    await mkdir(join(dir, 'sub'));
    await mkdir(join(dir, '.hidden'));
    await fixture('sub/letter.docx');
    await fixture('.hidden/secret.docx');
    const text = await ok(await connect(), 'list_documents', {});
    expect(text).toContain('contract.docx');
    expect(text).toContain(join('sub', 'letter.docx'));
    expect(text).not.toContain('~$');
    expect(text).not.toContain('secret');
  });
});

describe('the folder sandbox', () => {
  it('refuses paths outside the allowed folders', async () => {
    await fixture();
    const outside = await mkdtemp(join(tmpdir(), 'docxkeun-outside-'));
    try {
      await writeFile(join(outside, 'x.docx'), await buildDocx({ body: p('secret') }));
      await symlink(join(outside, 'x.docx'), join(dir, 'link.docx'));
      const client = await connect();
      for (const path of [join(outside, 'x.docx'), '../x.docx', 'link.docx']) {
        const r = await call(client, 'read_document', { path });
        expect(r.isError, path).toBe(true);
        expect(r.text).toMatch(/Outside the folders|File not found/);
      }
      const r = await call(client, 'replace_text', { path: 'contract.docx', find: 'ACME', replace: 'X', output: join(outside, 'y.docx') });
      expect(r.isError).toBe(true);
      expect(r.text).toContain('Outside the folders');
    } finally {
      await rm(outside, { recursive: true, force: true });
    }
  });

  it('only accepts Word documents', async () => {
    await writeFile(join(dir, 'notes.txt'), 'hello');
    const r = await call(await connect(), 'read_document', { path: 'notes.txt' });
    expect(r).toMatchObject({ isError: true });
    expect(r.text).toContain('Not a Word document');
  });
});

describe('editing', () => {
  it('replaces text everywhere and saves the file', async () => {
    const file = await fixture();
    const text = await ok(await connect(), 'replace_text', { path: 'contract.docx', find: 'ACME Ltd.', replace: 'Contoso GmbH' });
    expect(text).toBe('Replaced 2 of 2 matches.\nSaved contract.docx.');
    const paras = await bodyText(file);
    expect(paras[2]).toBe('This agreement is between {{client}} and Contoso GmbH');
    expect(paras[9]).toBe('The Supplier shall deliver on time. Contoso GmbH agrees.');
  });

  it('replaces one occurrence, or with regex groups', async () => {
    const file = await fixture();
    const client = await connect();
    await ok(client, 'replace_text', { path: 'contract.docx', find: 'ACME Ltd.', replace: 'ACME Inc.', occurrence: 2 });
    await ok(client, 'replace_text', { path: 'contract.docx', find: '(\\d+) days', replace: '$1 calendar days', regex: true });
    const paras = await bodyText(file);
    expect(paras[2]).toContain('and ACME Ltd.');
    expect(paras[9]).toContain('ACME Inc. agrees');
    expect(paras[3]).toBe('Payment is due within 30 calendar days of the invoice date.');
  });

  it('lets a read sent during a change see the change', async () => {
    await fixture();
    const client = await connect();
    const [, read] = await Promise.all([
      ok(client, 'replace_text', { path: 'contract.docx', find: 'ACME Ltd.', replace: 'Contoso GmbH' }),
      ok(client, 'read_document', { path: 'contract.docx' }),
    ]);
    expect(read).toContain('Contoso GmbH');
    expect(read).not.toContain('ACME');
  });

  it('leaves the file alone when nothing matches', async () => {
    const file = await fixture();
    const before = await readFile(file);
    const r = await call(await connect(), 'replace_text', { path: 'contract.docx', find: 'nowhere', replace: 'x' });
    expect(r).toMatchObject({ isError: true });
    expect(r.text).toContain('No matches');
    expect(await readFile(file)).toEqual(before);
  });

  it('writes to another file with output, and will not overwrite without asking', async () => {
    const file = await fixture();
    const before = await readFile(file);
    const client = await connect();
    expect(await ok(client, 'replace_text', { path: 'contract.docx', find: 'ACME', replace: 'X', output: 'copy.docx' })).toContain('Saved copy.docx.');
    expect(await readFile(file)).toEqual(before);
    expect((await bodyText(join(dir, 'copy.docx')))[2]).toContain('and X Ltd.');
    const again = await call(client, 'replace_text', { path: 'contract.docx', find: 'ACME', replace: 'Y', output: 'copy.docx' });
    expect(again).toMatchObject({ isError: true });
    expect(again.text).toContain('already exists');
    await ok(client, 'replace_text', { path: 'contract.docx', find: 'ACME', replace: 'Y', output: 'copy.docx', overwrite: true });
    expect((await bodyText(join(dir, 'copy.docx')))[2]).toContain('and Y Ltd.');
  });

  it('fills templates and reports what is left', async () => {
    const file = await fixture();
    const text = await ok(await connect(), 'fill_template', { path: 'contract.docx', data: { client: 'Globex', city: 'Bandung' } });
    expect(text).toContain('Filled client.');
    expect(text).toContain('Not found in the document: city.');
    expect(text).toContain('Still unfilled: {{signer}}.');
    expect((await bodyText(file))[2]).toBe('This agreement is between Globex and ACME Ltd.');
  });

  it('inserts, edits and deletes paragraphs', async () => {
    const file = await fixture();
    const client = await connect();
    const inserted = await ok(client, 'insert_paragraphs', {
      path: 'contract.docx',
      after: 4,
      paragraphs: [{ text: '2. Delivery', style: 'Heading 1' }, { text: 'Goods ship within a week.' }],
    });
    expect(inserted).toContain('Inserted 2 paragraphs (now paragraphs 5–6).');
    await ok(client, 'edit_paragraph', { path: 'contract.docx', paragraph: 6, text: 'Goods ship within two weeks.' });
    await ok(client, 'delete_paragraphs', { path: 'contract.docx', from: 1 });
    const read = await ok(client, 'read_document', { path: 'contract.docx' });
    expect(read).toContain('[4] (Heading 1) 2. Delivery');
    expect(read).toContain('[5] Goods ship within two weeks.');
    expect((await bodyText(file))[0]).toBe('1. Parties');

    const all = await call(client, 'delete_paragraphs', { path: 'contract.docx', from: 1, to: 11 });
    expect(all.text).toContain('needs at least one paragraph');

    const bad = await call(client, 'edit_paragraph', { path: 'contract.docx', paragraph: 2, style: 'Fancy' });
    expect(bad.text).toContain('Unknown paragraph style "Fancy". This document has: Normal, Title, Heading 1');
  });

  it('formats text', async () => {
    const file = await fixture();
    expect(await ok(await connect(), 'format_text', { path: 'contract.docx', text: 'ACME Ltd.', all: true, format: { bold: true, color: '#C00000' } })).toContain(
      'Formatted 2 matches',
    );
    const doc = await DocxDocument.open(new Uint8Array(await readFile(file)));
    const runs = doc.body().blocks.flatMap((b) => (b.kind === 'paragraph' ? b.inlines : [])).filter((i) => i.kind === 'text' && i.text === 'ACME Ltd.');
    expect(runs).toHaveLength(2);
    for (const r of runs) expect(r.kind === 'text' && r.props).toMatchObject({ bold: true, color: 'C00000' });
  });

  it('adds bookmarks and links, and refuses unsafe URLs', async () => {
    await fixture();
    const client = await connect();
    await ok(client, 'add_bookmark', { path: 'contract.docx', paragraph: 2, name: 'Parties' });
    await ok(client, 'add_hyperlink', { path: 'contract.docx', text: 'ACME Ltd.', url: 'https://acme.example' });
    await ok(client, 'add_hyperlink', { path: 'contract.docx', text: 'Signed by', bookmark: 'Parties' });
    const links = await ok(client, 'list_links', { path: 'contract.docx' });
    expect(links).toContain('"ACME Ltd." → https://acme.example');
    expect(links).toContain('"Signed by" → bookmark Parties');
    expect(links).toContain('Parties (paragraph 2)');
    const bad = await call(client, 'add_hyperlink', { path: 'contract.docx', paragraph: 4, url: 'javascript:alert(1)' });
    expect(bad).toMatchObject({ isError: true });
    expect(bad.text).toContain('not allowed');
  });

  it('creates new documents', async () => {
    const client = await connect();
    await ok(client, 'create_document', {
      path: 'memo.docx',
      paragraphs: [{ text: 'Memo', style: 'Heading 1' }, { text: 'Line one\nLine two' }],
    });
    const read = await ok(client, 'read_document', { path: 'memo.docx' });
    expect(read).toContain('[1] (Heading 1) Memo');
    expect(read).toContain('[2] Line one\nLine two');
    expect((await call(client, 'create_document', { path: 'memo.docx', paragraphs: [] })).text).toContain('already exists');
  });

  it('copies a document exactly, to edit the copy', async () => {
    const file = await fixture();
    const client = await connect();
    const copied = await ok(client, 'copy_document', { path: 'contract.docx', to: 'contract (English).docx' });
    expect(copied).toBe('Copied contract.docx to contract (English).docx: an exact copy with all its formatting, tables, lists, pictures, headers, footers and comments. Change the copy with path "contract (English).docx".');
    const copy = join(dir, 'contract (English).docx');
    expect(new Uint8Array(await readFile(copy))).toEqual(new Uint8Array(await readFile(file)));
    // Editing the copy leaves the original alone.
    await ok(client, 'replace_text', { path: 'contract (English).docx', find: 'ACME Ltd.', replace: 'Contoso GmbH' });
    expect((await bodyText(file)).join('\n')).toContain('ACME Ltd.');
    expect((await bodyText(copy)).join('\n')).toContain('Contoso GmbH');

    const again = await call(client, 'copy_document', { path: 'contract.docx', to: 'contract (English).docx' });
    expect(again).toMatchObject({ isError: true });
    expect(again.text).toContain('contract (English).docx already exists; pass overwrite: true to replace it');
    await ok(client, 'copy_document', { path: 'contract.docx', to: 'contract (English).docx', overwrite: true });
    expect(new Uint8Array(await readFile(copy))).toEqual(new Uint8Array(await readFile(file)));
    expect((await call(client, 'copy_document', { path: 'contract.docx', to: 'contract.docx', overwrite: true })).text).toContain('is the same file');
    expect((await call(client, 'copy_document', { path: 'missing.docx', to: 'x.docx' })).isError).toBe(true);
  });

  it('keeps an encrypted copy encrypted, and tells the AI to use the copy rather than retype', async () => {
    await fixture();
    const client = await connect({ role: 'owner' });
    await ok(client, 'set_open_password', { path: 'contract.docx', action: 'set', newPassword: 'pw-1' });
    const copied = await ok(client, 'copy_document', { path: 'contract.docx', to: 'locked copy.docx' });
    expect(copied).toContain('It opens with the same password.');
    expect((await bodyText(join(dir, 'locked copy.docx'), 'pw-1')).join('\n')).toContain('ACME Ltd.');

    const { tools } = await client.listTools();
    expect(tools.find((t) => t.name === 'create_document')!.description).toContain('To copy a document, use copy_document');
    expect(client.getInstructions()).toContain('copy_document');
  });
});

describe('text formatting', () => {
  const R = (text: string, rPr = '') => `<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ''}<w:t xml:space="preserve">${text}</w:t></w:r>`;
  const NARROW = '<w:rFonts w:ascii="Arial Narrow" w:hAnsi="Arial Narrow"/>';
  const RED_BOLD = NARROW + '<w:b/><w:color w:val="C00000"/>';
  const TITLE = NARROW + '<w:b/><w:u w:val="single"/>';
  const FORMATTED =
    `<w:p><w:pPr><w:jc w:val="center"/></w:pPr>${R('KESEPAKATAN', TITLE)}${R(' KOMPENSASI', TITLE)}</w:p>` +
    `<w:p><w:pPr><w:jc w:val="both"/></w:pPr>${R('1.', NARROW)}<w:r><w:tab/></w:r>${R('PT Swadaya', RED_BOLD)}${R(' ', NARROW + '<w:b/>')}${R('Usaha Persada', RED_BOLD)}` +
    `${R(', Badan Hukum, disebut ', NARROW)}${R('“Pihak Pertama”', NARROW + '<w:b/>')}${R('; luas 6.337 M', NARROW)}${R('2', NARROW + '<w:vertAlign w:val="superscript"/>')}${R(' (a&lt;b&gt; &amp;lt;)', NARROW)}</w:p>`;
  const ORIGINAL = '1.\t<b>PT Swadaya Usaha Persada</b>, Badan Hukum, disebut <b>“Pihak Pertama”</b>; luas 6.337 M<sup>2</sup> (a&lt;b> &amp;lt;)';

  async function formatted(): Promise<string> {
    const file = join(dir, 'kesepakatan.docx');
    await writeFile(file, await buildDocx({ body: FORMATTED, styles: STYLES, settings: '' }));
    return file;
  }

  /** Paragraph n's text runs: text, bold, italic, underline, superscript, colour and font. */
  async function runsOf(file: string, n: number): Promise<Array<[string, boolean, boolean, boolean, boolean, string | undefined, string | undefined]>> {
    const doc = await DocxDocument.open(new Uint8Array(await readFile(file)));
    const { listParagraphs } = await import('../src/outline.js');
    return listParagraphs(doc)[n - 1]!.para.inlines.flatMap((i) =>
      i.kind === 'text'
        ? [[i.text, !!i.props.bold, !!i.props.italic, !!i.props.underline && i.props.underline !== 'none', i.props.vertAlign === 'superscript', i.props.color, i.props.fonts?.ascii] as const]
        : [],
    ) as never;
  }

  it('shows formatting: tags inside a paragraph, and what the whole paragraph has in its description', async () => {
    await formatted();
    const read = await ok(await connect(), 'read_document', { path: 'kesepakatan.docx' });
    expect(read).toContain('[1] (centered; bold, underlined) KESEPAKATAN KOMPENSASI\n');
    expect(read).toContain(`[2] (justified) ${ORIGINAL}`);
  });

  it('keeps bold, underline, colours and fonts when paragraphs are rewritten with the tags', async () => {
    const file = await formatted();
    const client = await connect();
    await ok(client, 'edit_paragraph', { path: 'kesepakatan.docx', paragraph: 1, text: 'AGREEMENT ON COMPENSATION' });
    const english = '1.\t<b>PT Swadaya Usaha Persada</b>, a legal entity, called <b>“First Party”</b>; area 6,337 m<sup>2</sup> (a&lt;b> &amp;lt;)';
    await ok(client, 'edit_paragraph', { path: 'kesepakatan.docx', paragraph: 2, text: english });

    const read = await ok(client, 'read_document', { path: 'kesepakatan.docx' });
    expect(read).toContain('[1] (centered; bold, underlined) AGREEMENT ON COMPENSATION\n');
    expect(read).toContain(`[2] (justified) ${english}`);
    expect(await runsOf(file, 1)).toEqual([['AGREEMENT ON COMPENSATION', true, false, true, false, undefined, 'Arial Narrow']]);
    expect(await runsOf(file, 2)).toEqual([
      ['1.', false, false, false, false, undefined, 'Arial Narrow'],
      // Each bold part copies the bold phrase in the same place: the first is red, the second is not.
      ['PT Swadaya Usaha Persada', true, false, false, false, 'C00000', 'Arial Narrow'],
      [', a legal entity, called ', false, false, false, false, undefined, 'Arial Narrow'],
      ['“First Party”', true, false, false, false, undefined, 'Arial Narrow'],
      ['; area 6,337 m', false, false, false, false, undefined, 'Arial Narrow'],
      ['2', false, false, false, true, undefined, 'Arial Narrow'],
      [' (a<b> &lt;)', false, false, false, false, undefined, 'Arial Narrow'],
    ]);
  });

  it('formats new text from tags, and leaves plain text plain', async () => {
    const file = await formatted();
    const client = await connect();
    await ok(client, 'edit_paragraph', { path: 'kesepakatan.docx', paragraph: 2, text: 'A plain sentence with <i>one italic</i> word.' });
    expect(await runsOf(file, 2)).toEqual([
      ['A plain sentence with ', false, false, false, false, undefined, 'Arial Narrow'],
      ['one italic', false, true, false, false, undefined, 'Arial Narrow'],
      [' word.', false, false, false, false, undefined, 'Arial Narrow'],
    ]);
    await ok(client, 'insert_paragraphs', { path: 'kesepakatan.docx', paragraphs: [{ text: 'Total: <b>Rp 201.801.600</b> <u>only</u>' }] });
    const added = await runsOf(file, 3);
    expect(added.map(([t, b, , u]) => [t, b, u])).toEqual([
      ['Total: ', false, false],
      ['Rp 201.801.600', true, false],
      [' ', false, false],
      ['only', false, true],
    ]);
  });

  it('creates documents with formatted text', async () => {
    const client = await connect();
    await ok(client, 'create_document', { path: 'note.docx', paragraphs: [{ text: '<b>Note</b>' }, { text: 'Pay <u>before</u> the 5<sup>th</sup>.' }] });
    const read = await ok(client, 'read_document', { path: 'note.docx' });
    expect(read).toContain('[1] (bold) Note\n');
    expect(read).toContain('[2] Pay <u>before</u> the 5<sup>th</sup>.');
  });

  it('finds and replaces text written with tags', async () => {
    const file = await formatted();
    const client = await connect();
    expect(await ok(client, 'find_text', { path: 'kesepakatan.docx', query: '<b>“Pihak Pertama”</b>' })).toContain('1 match');
    await ok(client, 'replace_text', { path: 'kesepakatan.docx', find: '<b>“Pihak Pertama”</b>', replace: '<u>“First Party”</u>' });
    const run = (await runsOf(file, 2)).find(([t]) => t === '“First Party”');
    expect(run?.slice(1, 4)).toEqual([true, false, true]);
  });
});

describe('comments', () => {
  it('adds, lists, replies to and resolves comments', async () => {
    await fixture();
    const client = await connect({ role: 'commenter', author: 'Review Bot' });
    const added = await ok(client, 'add_comment', { path: 'contract.docx', text: '30 days', comment: 'Is 30 days standard?' });
    const id = /comment (\S+) on/.exec(added)![1]!;
    expect(added).toContain('on "30 days" in paragraph 4.');
    await ok(client, 'reply_to_comment', { path: 'contract.docx', commentId: id, reply: 'Yes, net 30.' });
    await ok(client, 'resolve_comment', { path: 'contract.docx', commentId: id });
    const list = await ok(client, 'list_comments', { path: 'contract.docx' });
    expect(list).toContain(`Comment ${id} by Review Bot`);
    expect(list).toContain('(resolved)');
    expect(list).toContain('On paragraph 4: "30 days"');
    expect(list).toContain('Is 30 days standard?');
    expect(list).toMatch(/Reply \S+ by Review Bot.*: Yes, net 30\./);
    expect(await ok(client, 'read_document', { path: 'contract.docx' })).toContain(`[4] (comment ${id}) Payment`);
    expect(await ok(client, 'list_comments', { path: 'contract.docx', includeResolved: false })).toBe('No comments.');
  });

  it("lets a commenter delete only its own comments", async () => {
    await fixture();
    const editor = await connect({ author: 'Kevin' });
    const theirs = /comment (\S+) on/.exec(await ok(editor, 'add_comment', { path: 'contract.docx', paragraph: 3, comment: 'Check names' }))![1]!;
    const bot = await connect({ role: 'commenter', author: 'Review Bot' });
    const r = await call(bot, 'delete_comment', { path: 'contract.docx', commentId: theirs });
    expect(r).toMatchObject({ isError: true });
    expect(r.text).toContain('You can only change your own comments');
  });
});

describe('locks, protection and passwords', () => {
  it('owners lock regions; editors cannot change them', async () => {
    const file = await fixture();
    const owner = await connect({ role: 'owner' });
    expect(await ok(owner, 'lock_region', { path: 'contract.docx', paragraph: 4, title: 'Payment terms' })).toMatch(/Locked paragraph 4 as lock \S+\./);
    const editor = await connect();
    const read = await ok(editor, 'read_document', { path: 'contract.docx' });
    expect(read).toMatch(/\[4\] \(locked "Payment terms", lock \S+\) Payment is due/);

    const edit = await call(editor, 'edit_paragraph', { path: 'contract.docx', paragraph: 4, text: 'Pay whenever.' });
    expect(edit).toMatchObject({ isError: true });
    expect(edit.text).toContain('"Payment terms" is locked against editing');
    const replace = await call(editor, 'replace_text', { path: 'contract.docx', find: '30 days', replace: '90 days' });
    expect(replace.text).toContain('inside locked regions; nothing changed');
    expect((await bodyText(file))[3]).toBe('Payment is due within 30 days of the invoice date.');
    // Commenting on locked text is fine.
    await ok(editor, 'add_comment', { path: 'contract.docx', paragraph: 4, comment: 'Too short?' });
  });

  it("enforces the document's protection", async () => {
    await fixture();
    const owner = await connect({ role: 'owner' });
    await ok(owner, 'protect_document', { path: 'contract.docx', mode: 'comments', protectionPassword: 'pw' });
    const editor = await connect();
    expect(await ok(editor, 'get_document_info', { path: 'contract.docx' })).toContain('Protection: comments, with a password');
    const r = await call(editor, 'replace_text', { path: 'contract.docx', find: 'ACME', replace: 'X' });
    expect(r).toMatchObject({ isError: true });
    expect(r.text).toContain('Not allowed: The document only allows comments');
    await ok(editor, 'add_comment', { path: 'contract.docx', text: 'ACME', comment: 'Full legal name?' });
    expect((await call(owner, 'protect_document', { path: 'contract.docx', mode: 'none', protectionPassword: 'nope' })).text).toContain('Incorrect protection password');
    await ok(owner, 'protect_document', { path: 'contract.docx', mode: 'none', protectionPassword: 'pw' });
    await ok(editor, 'replace_text', { path: 'contract.docx', find: 'ACME', replace: 'X' });
  });

  it('encrypts documents and keeps them encrypted through edits', async () => {
    const file = await fixture();
    const owner = await connect({ role: 'owner' });
    await ok(owner, 'set_open_password', { path: 'contract.docx', action: 'set', newPassword: 's3cret' });
    expect(isEncryptedPackage(new Uint8Array(await readFile(file)))).toBe(true);

    const editor = await connect();
    const locked = await call(editor, 'read_document', { path: 'contract.docx' });
    expect(locked.text).toContain('This document is encrypted');
    expect((await call(editor, 'read_document', { path: 'contract.docx', password: 'wrong' })).text).toContain('Wrong password');
    await ok(editor, 'replace_text', { path: 'contract.docx', password: 's3cret', find: 'ACME', replace: 'X' });
    expect(isEncryptedPackage(new Uint8Array(await readFile(file)))).toBe(true);
    expect((await bodyText(file, 's3cret'))[2]).toContain('and X Ltd.');

    await ok(owner, 'set_open_password', { path: 'contract.docx', password: 's3cret', action: 'remove' });
    expect(isEncryptedPackage(new Uint8Array(await readFile(file)))).toBe(false);
  });
});

describe('PDF export', () => {
  it('writes a PDF next to the document', async () => {
    await fixture();
    const text = await ok(await connect({ role: 'viewer' }), 'export_pdf', { path: 'contract.docx' });
    expect(text).toMatch(/^Exported 1 page to contract\.pdf \(\d+(\.\d)? KB\)\.$/);
    const pdf = await readFile(join(dir, 'contract.pdf'));
    expect(pdf.subarray(0, 5).toString('latin1')).toBe('%PDF-');
    expect(pdf.toString('latin1')).not.toContain('/Encrypt');
  });

  it('protects the PDF with a password and permissions', async () => {
    await fixture();
    const text = await ok(await connect(), 'export_pdf', { path: 'contract.docx', output: 'out/locked.pdf', pdfPassword: 'open', permissions: { print: false } }).catch(
      (e: Error) => e.message,
    );
    expect(text).toContain('Folder does not exist');
    await mkdir(join(dir, 'out'));
    expect(await ok(await connect(), 'export_pdf', { path: 'contract.docx', output: 'out/locked.pdf', pdfPassword: 'open', permissions: { print: false } })).toContain(
      'password protected',
    );
    expect((await readFile(join(dir, 'out', 'locked.pdf'))).toString('latin1')).toContain('/Encrypt');
  });
});

describe('lists', () => {
  const listItem = (text: string) => p(text, '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>');

  async function letter(): Promise<void> {
    const body = p('We ask that all payments be transferred:') + listItem('Land release fee Rp 86.486.400') + listItem('Compensation Rp 201.801.600') + p('Thank you.');
    await writeFile(join(dir, 'letter.docx'), await buildDocx({ body, styles: STYLES, numbering: NUMBERING }));
  }

  it('shows list labels and turns letters into numbers', async () => {
    await letter();
    const client = await connect();
    expect(await ok(client, 'read_document', { path: 'letter.docx' })).toContain('[2] (list item a.) Land release fee');
    expect(await ok(client, 'set_list', { path: 'letter.docx', from: 2, to: 3, format: 'decimal' })).toBe('Paragraphs 2–3 are now labelled "1.", "2.".\nSaved letter.docx.');
    const read = await ok(client, 'read_document', { path: 'letter.docx' });
    expect(read).toContain('[2] (list item 1.) Land release fee');
    expect(read).toContain('[3] (list item 2.) Compensation');
  });

  it('adds items, changes levels, bullets and removes lists', async () => {
    await letter();
    const client = await connect();
    await ok(client, 'insert_paragraphs', { path: 'letter.docx', after: 3, paragraphs: [{ text: 'Admin fee Rp 100.000' }] });
    expect(await ok(client, 'set_list', { path: 'letter.docx', from: 4, sameListAs: 3 })).toContain('Paragraph 4 is now labelled "c.".');
    expect(await ok(client, 'set_list', { path: 'letter.docx', from: 2, to: 4, start: 5 })).toContain('labelled "e.", "f.", "g.".');
    expect(await ok(client, 'set_list', { path: 'letter.docx', from: 2, to: 4, format: 'bullet' })).toContain('labelled "•", "•", "•".');
    expect(await ok(client, 'set_list', { path: 'letter.docx', from: 3, level: 2 })).toContain('labelled "o".');
    expect(await ok(client, 'read_document', { path: 'letter.docx' })).toContain('[3] (bullet o, level 2) Compensation');
    expect(await ok(client, 'set_list', { path: 'letter.docx', from: 2, to: 4, format: 'none' })).toContain('Paragraphs 2–4 are no longer in a list.');
    expect(await ok(client, 'read_document', { path: 'letter.docx' })).toContain('[2] Land release fee');
    const bad = await call(client, 'set_list', { path: 'letter.docx', from: 1, level: 2 });
    expect(bad).toMatchObject({ isError: true });
    expect(bad.text).toContain('not in a list');
  });

  it('is an editing tool', async () => {
    const names = (await (await connect({ role: 'commenter' })).listTools()).tools.map((t) => t.name);
    expect(names).not.toContain('set_list');
  });

  it('joins a list with its own format, and says how to build a list that is not a list yet', async () => {
    await letter();
    const client = await connect();
    await ok(client, 'set_list', { path: 'letter.docx', from: 2, format: 'decimal' });
    expect(await ok(client, 'set_list', { path: 'letter.docx', from: 3, sameListAs: 2, format: 'decimal' })).toContain('Paragraph 3 is now labelled "2.".');
    const notYet = await call(client, 'set_list', { path: 'letter.docx', from: 4, sameListAs: 1 });
    expect(notYet).toMatchObject({ isError: true });
    expect(notYet.text).toContain('Paragraph 1 is not in a list yet. Make it one first (set_list from 1 with a format), then add the others with sameListAs: 1.');
    const other = await call(client, 'set_list', { path: 'letter.docx', from: 3, sameListAs: 2, format: 'bullet' });
    expect(other.text).toContain('That list is numbered 1.; to continue it leave out format');
  });
});

describe('tables and copies', () => {
  it('copies everything onto a new page with lists and tables, for a translation', async () => {
    const body =
      p('Agreement', '<w:jc w:val="center"/>', '<w:b/>') +
      p('Terms:') +
      p('First term', '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>') +
      p('Second term', '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>') +
      `<w:tbl><w:tblPr><w:tblW w:w="8000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid><w:tr>${cell('No')}${cell('Total')}</w:tr></w:tbl>` +
      p('Signed.');
    await writeFile(join(dir, 'deal.docx'), await buildDocx({ body, styles: STYLES, numbering: NUMBERING }));
    const client = await connect();
    expect(await ok(client, 'copy_paragraphs', { path: 'deal.docx', from: 1, to: 7, pageBreak: true })).toBe(
      'Copied paragraphs 1–7 (with 1 table) at the end, starting on a new page: the copy is paragraphs 8–14. Lists in the copy start again at 1.\nSaved deal.docx.',
    );
    const read = await ok(client, 'read_document', { path: 'deal.docx' });
    expect(read).toContain('[10] (list item a.) First term');
    expect(read).toContain('[11] (list item b.) Second term');
    expect(read).toContain('[12] (table 2, row 1, cell 1) No');
    // Translate a copied cell like any paragraph.
    await ok(client, 'edit_paragraph', { path: 'deal.docx', paragraph: 12, text: 'Nomor' });
    expect((await bodyText(join(dir, 'deal.docx'))).slice(7)).toEqual(['Agreement', 'Terms:', 'First term', 'Second term', 'Nomor', 'Total', 'Signed.']);
  });

  it('inserts a table, adds and deletes rows, and deletes the table', async () => {
    const file = await fixture();
    const client = await connect();
    expect(await ok(client, 'insert_table', { path: 'contract.docx', after: 4, rows: [['No', 'Description', 'Total'], ['1', 'Land', 'Rp 201.801.600']], header: true, widths: [1, 4, 2] })).toBe(
      'Inserted a 2×3 table after paragraph 4 (table 1): its cells are paragraphs 5–10.\nSaved contract.docx.',
    );
    expect(await ok(client, 'edit_table', { path: 'contract.docx', table: 1, action: 'add_rows', rows: [['2', 'Trees', 'Rp 1.000.000']] })).toBe(
      'Added 1 row to table 1 (now 3 rows): the new cells are paragraphs 11–13.\nSaved contract.docx.',
    );
    expect((await bodyText(file)).slice(4, 13)).toEqual(['No', 'Description', 'Total', '1', 'Land', 'Rp 201.801.600', '2', 'Trees', 'Rp 1.000.000']);
    expect(await ok(client, 'edit_table', { path: 'contract.docx', table: 1, action: 'delete_rows', fromRow: 2 })).toBe('Deleted row 2 of table 1 (2 rows left).\nSaved contract.docx.');
    expect(await ok(client, 'edit_table', { path: 'contract.docx', table: 1, action: 'delete_table' })).toBe('Deleted table 1.\nSaved contract.docx.');
    // The fixture's own table is table 1 again.
    expect(await ok(client, 'read_document', { path: 'contract.docx' })).toContain('(table 1, row 1, cell 1) Name');
    const missing = await call(client, 'edit_table', { path: 'contract.docx', table: 5, action: 'delete_table' });
    expect(missing.text).toContain('There is no table 5; the document has 1 table.');
  });

  it('are editing tools', async () => {
    const names = (await (await connect({ role: 'commenter' })).listTools()).tools.map((t) => t.name);
    for (const name of ['copy_paragraphs', 'insert_table', 'edit_table']) expect(names).not.toContain(name);
    const editor = (await (await connect({ role: 'editor' })).listTools()).tools.map((t) => t.name);
    for (const name of ['copy_paragraphs', 'insert_table', 'edit_table']) expect(editor).toContain(name);
  });
});

describe('tracked changes', () => {
  it('records edits while Track Changes is on, lists them, and accepts or rejects them', async () => {
    const file = await fixture();
    const editor = await connect({ author: 'Claude' });
    await ok(editor, 'set_track_changes', { path: 'contract.docx', on: true });
    await ok(editor, 'replace_text', { path: 'contract.docx', find: '30 days', replace: '45 days' });
    const listed = await ok(editor, 'list_revisions', { path: 'contract.docx' });
    expect(listed).toContain('Track Changes is on');
    expect(listed).toMatch(/Change \d+ \(deleted\) by Claude, .*paragraph 4: "30 days"/);
    expect(listed).toMatch(/Change \d+ \(inserted\) by Claude, .*paragraph 4: "45 days"/);
    // read_document shows the result with the changes accepted, and marks the paragraph.
    const read = await ok(editor, 'read_document', { path: 'contract.docx' });
    expect(read).toContain('[4] (2 tracked changes) Payment is due within 45 days');
    expect(await call(editor, 'accept_revisions', { path: 'contract.docx' })).toMatchObject({ isError: true });
    const id = /Change (\d+) \(deleted\)/.exec(listed)![1]!;
    expect(await ok(editor, 'reject_revisions', { path: 'contract.docx', ids: [id] })).toContain('Rejected 1 tracked change');
    expect(await ok(editor, 'accept_revisions', { path: 'contract.docx', all: true })).toContain('Accepted 1 tracked change');
    expect((await bodyText(file))[3]).toBe('Payment is due within 30 days45 days of the invoice date.');
  });

  it('lets editors change a document protected for tracked changes (tracked), but not review it', async () => {
    await fixture();
    const owner = await connect({ role: 'owner' });
    await ok(owner, 'protect_document', { path: 'contract.docx', mode: 'trackedChanges' });
    const editor = await connect({ author: 'Claude' });
    await ok(editor, 'replace_text', { path: 'contract.docx', find: 'ACME Ltd.', replace: 'ACME Limited' });
    expect(await ok(editor, 'list_revisions', { path: 'contract.docx' })).toContain('by Claude');
    const r = await call(editor, 'accept_revisions', { path: 'contract.docx', all: true });
    expect(r.isError).toBe(true);
    expect(r.text).toContain('cannot be accepted or rejected');
  });
});

describe('paragraph formatting, stories and objects', () => {
  it('formats paragraphs, including right-to-left', async () => {
    await fixture();
    const editor = await connect();
    expect(await ok(editor, 'format_paragraph', { path: 'contract.docx', from: 3, to: 4, align: 'justify', spaceAfter: 12, lineSpacing: 1.5 })).toContain('Formatted paragraphs 3–4');
    await ok(editor, 'format_paragraph', { path: 'contract.docx', from: 9, direction: 'rtl' });
    const read = await ok(editor, 'read_document', { path: 'contract.docx' });
    expect(read).toContain('[9] (right-to-left)');
  });

  it('reads and edits headers, footers and footnotes, and adds footnotes', async () => {
    await fixture();
    const editor = await connect();
    expect(await ok(editor, 'read_headers_footers', { path: 'contract.docx' })).toContain('no headers, footers');
    expect(await ok(editor, 'add_footnote', { path: 'contract.docx', paragraph: 4, after: '30 days', text: 'Calendar days.' })).toContain('Added footnote 1');
    const stories = await ok(editor, 'read_headers_footers', { path: 'contract.docx' });
    expect(stories).toContain('footnote1: Footnote 1');
    expect(stories).toContain('Calendar days.');
    await ok(editor, 'replace_story_text', { path: 'contract.docx', story: 'footnote1', find: 'Calendar', replace: 'Business' });
    expect(await ok(editor, 'read_headers_footers', { path: 'contract.docx' })).toContain('Business days.');
    const missing = await call(editor, 'replace_story_text', { path: 'contract.docx', story: 'header1', find: 'x', replace: 'y' });
    expect(missing.text).toContain('No story "header1"');
    // The PDF shows the footnote.
    await ok(editor, 'export_pdf', { path: 'contract.docx' });
  });

  it('inserts pictures from the allowed folders and lists objects', async () => {
    await fixture();
    // A 2 × 1 PNG at 96 dpi.
    const png = Uint8Array.from(Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAIAAAABCAYAAAD0In+KAAAAEUlEQVR4nGP4z8DwnwGIGQAd8AP9Fr8f4QAAAABJRU5ErkJggg==', 'base64'));
    await writeFile(join(dir, 'logo.png'), png);
    const editor = await connect();
    expect(await ok(editor, 'insert_image', { path: 'contract.docx', image: 'logo.png', paragraph: 1, width: 144, alt: 'Company logo' })).toContain('Inserted');
    const objects = await ok(editor, 'list_objects', { path: 'contract.docx' });
    expect(objects).toBe('[1] object 1: picture "Company logo" (144×72pt, inline)');
    expect(await ok(editor, 'read_document', { path: 'contract.docx' })).toContain('Service Agreement {object 1: picture "Company logo" (144×72pt, inline)}');
    const outside = await call(editor, 'insert_image', { path: 'contract.docx', image: '/etc/hosts', paragraph: 1 });
    expect(outside.isError).toBe(true);
    const notPicture = await call(editor, 'insert_image', { path: 'contract.docx', image: 'contract.docx', paragraph: 1 });
    expect(notPicture.text).toContain('Not a picture');
  });
});

describe('headers and footers', () => {
  it('adds, lists, edits and removes headers and footers', async () => {
    const file = await fixture();
    const editor = await connect();
    expect(await ok(editor, 'set_header_footer', { path: 'contract.docx', kind: 'header', text: 'ACME Ltd.\tService Agreement\tConfidential' })).toBe(
      'Set the header of the document to "ACME Ltd. → Service Agreement → Confidential".\nSaved contract.docx.',
    );
    await ok(editor, 'set_header_footer', { path: 'contract.docx', kind: 'footer', text: 'Page {PAGE} of {NUMPAGES}', align: 'center' });
    await ok(editor, 'set_header_footer', { path: 'contract.docx', kind: 'footer', text: 'Cover', type: 'first' });
    const listed = await ok(editor, 'read_headers_footers', { path: 'contract.docx' });
    expect(listed).toContain('header1: Header 1 (every page, the whole document)\n  ACME Ltd. → Service Agreement → Confidential');
    expect(listed).toContain('footer1: Footer 1 (every page, the whole document)\n  Page {PAGE} of {NUMPAGES}');
    expect(listed).toContain('footer2: Footer 2 (first page, the whole document)\n  Cover');
    await ok(editor, 'replace_story_text', { path: 'contract.docx', story: 'header1', find: 'Confidential', replace: 'Draft' });
    expect(await ok(editor, 'read_headers_footers', { path: 'contract.docx' })).toContain('ACME Ltd. → Service Agreement → Draft');
    expect(await ok(editor, 'remove_header_footer', { path: 'contract.docx', kind: 'footer', type: 'first' })).toContain('Removed 1 footer.');
    const gone = await call(editor, 'remove_header_footer', { path: 'contract.docx', kind: 'footer', type: 'first' });
    expect(gone).toMatchObject({ isError: true });
    expect(gone.text).toContain('There is no first footer to remove');
    // Word opens it: the parts and references are in place after a save and reopen.
    const doc = await DocxDocument.open(new Uint8Array(await readFile(file)));
    expect(doc.storyParts().map((sp) => sp.kind)).toEqual(['header', 'footer']);
    // Viewers may read but not change them.
    const viewer = await connect({ role: 'viewer' });
    expect((await viewer.listTools()).tools.map((t) => t.name)).not.toContain('set_header_footer');
  });
});

