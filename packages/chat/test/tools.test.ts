/** The chat's tools: definitions sent to the API, and running them on the open document. */
import { describe, expect, it } from 'vitest';
import { DocxDocument } from 'docxkeun';
import { DocxEditor } from 'docxkeun/editor';
import { buildDocx, p } from '../../../test/helpers/docx.js';
import { renderMarkdown } from '../src/markdown.js';
import { createToolbox } from '../src/tools.js';
import { listParagraphs } from '../../mcp/src/outline.js';

const NUMBERING =
  '<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="%1."/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum>' +
  '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>';
const item = (t: string) => p(t, '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>');

describe('chat tools', () => {
  it('offers the editor operations with clean JSON schemas', () => {
    const { definitions } = createToolbox(() => []);
    const names = definitions.map((d) => d.name);
    expect(names).toContain('set_list');
    expect(names).toContain('add_comment');
    expect(names).not.toContain('lock_region');
    for (const d of definitions) {
      expect(d.input_schema.type).toBe('object');
      expect(d.input_schema).not.toHaveProperty('$schema');
      expect(d.eager_input_streaming).toBe(true);
    }
    const setList = definitions.find((d) => d.name === 'set_list')!;
    expect(setList.input_schema.required).toEqual(['from']);
    expect((setList.input_schema.properties as Record<string, { enum?: string[] }>).format!.enum).toContain('decimal');
  });

  it('changes the open document, which the viewer redraws from', async () => {
    const doc = await DocxDocument.open(await buildDocx({ body: p('Amounts:') + item('Land fee') + item('Compensation'), numbering: NUMBERING }));
    const changes: string[] = [];
    doc.onChange((part) => changes.push(part));
    const letter = { doc, name: 'letter.docx' };
    const toolbox = createToolbox(() => [letter]);
    const result = await toolbox.run('set_list', { from: 2, to: 3, format: 'decimal' });
    expect(result).toEqual({ content: 'Paragraphs 2–3 are now labelled "1.", "2.".', isError: false, kind: 'change', document: letter });
    expect(changes).toContain('word/document.xml');
    const comment = await toolbox.run('add_comment', { text: 'Land fee', comment: 'Check the amount' });
    expect(comment).toMatchObject({ isError: false, kind: 'add', comment: '0' });
    const read = await toolbox.run('read_document', {});
    expect(read.content).toContain('[2] (list item 1.; comment 0) Land fee');
  });

  it('explains refusals from locks and protection', async () => {
    const doc = await DocxDocument.open(await buildDocx({ body: p('Fixed terms') + p('Free text'), settings: '' }));
    const owner = new DocxEditor(doc);
    owner.lockParagraphs(owner.paragraphs()[0]!.el, undefined, { title: 'Agreed' });
    const toolbox = createToolbox(() => [{ doc, name: 'x.docx' }]);
    const locked = await toolbox.run('edit_paragraph', { paragraph: 1, text: 'Changed' });
    expect(locked).toMatchObject({ isError: true });
    expect(locked.content).toContain('"Agreed" is locked against editing');
    expect((await toolbox.run('nope', {})).content).toBe('There is no tool named nope.');
  });
});

describe('chat tools with two documents open', () => {
  const open = async (name: string, ...texts: string[]) => ({ doc: await DocxDocument.open(await buildDocx({ body: texts.map((t) => p(t)).join('') })), name });
  const text = (d: { doc: DocxDocument }) => listParagraphs(d.doc).map((e) => e.text);

  it('lets every tool name its document', () => {
    for (const d of createToolbox(() => []).definitions) {
      const props = d.input_schema.properties as Record<string, { type?: string; description?: string }>;
      expect(props.document, d.name).toMatchObject({ type: 'string' });
      expect(props.document!.description).toMatch(/file name/i);
      expect(d.input_schema.required ?? [], d.name).not.toContain('document');
    }
  });

  it('runs on the document the call names, and only that one', async () => {
    const contract = await open('contract.docx', 'Payment within 30 days.');
    const template = await open('template.docx', 'Payment within [N] days.');
    const toolbox = createToolbox(() => [contract, template]);
    const result = await toolbox.run('replace_text', { document: 'template.docx', find: '[N]', replace: '45' });
    expect(result).toMatchObject({ isError: false, kind: 'change', document: template });
    expect(text(template)).toEqual(['Payment within 45 days.']);
    expect(text(contract)).toEqual(['Payment within 30 days.']);
    // The name is matched without regard to case, and the other document reads as itself.
    const read = await toolbox.run('read_document', { document: 'CONTRACT.DOCX' });
    expect(read).toMatchObject({ isError: false, document: contract });
    expect(read.content).toContain('30 days');
  });

  it('refuses a call that does not say which of two documents, and changes nothing', async () => {
    const a = await open('a.docx', 'Alpha.');
    const b = await open('b.docx', 'Beta.');
    const toolbox = createToolbox(() => [a, b]);
    const result = await toolbox.run('edit_paragraph', { paragraph: 1, text: 'Changed.' });
    expect(result).toMatchObject({ isError: true, kind: 'change' });
    expect(result.content).toBe('Two documents are open: a.docx and b.docx. Say which one with `document`.');
    expect(result).not.toHaveProperty('document');
    expect([...text(a), ...text(b)]).toEqual(['Alpha.', 'Beta.']);
  });

  it('refuses a name that is not open, listing the ones that are', async () => {
    const a = await open('a.docx', 'Alpha.');
    const toolbox = createToolbox(() => [a]);
    const result = await toolbox.run('read_document', { document: 'missing.docx' });
    expect(result).toMatchObject({ isError: true });
    expect(result.content).toBe('No open document is called missing.docx. Open: a.docx.');
  });

  it('says which document a call is about while it is still being written', async () => {
    const a = await open('a.docx', 'Alpha.');
    const b = await open('b.docx', 'Beta.');
    const two = createToolbox(() => [a, b]);
    expect(two.target({ document: 'b.docx', paragraph: 1 })).toBe(b);
    expect(two.target({ paragraph: 1 })).toBeUndefined();
    expect(two.target(undefined)).toBeUndefined();
    expect(createToolbox(() => [a]).target({})).toBe(a);
  });
});

describe('renderMarkdown', () => {
  it('formats replies and escapes everything else', () => {
    expect(renderMarkdown('**Done.** I changed *two* lines in `letter.docx`.')).toBe('<p><strong>Done.</strong> I changed <em>two</em> lines in <code>letter.docx</code>.</p>');
    expect(renderMarkdown('Changes:\n- one\n- two\n\n1. first\n2. second')).toBe('<p>Changes:</p><ul><li>one</li><li>two</li></ul><ol><li>first</li><li>second</li></ol>');
    expect(renderMarkdown('<img src=x onerror=alert(1)> & "quotes"')).toBe('<p>&lt;img src=x onerror=alert(1)&gt; &amp; &quot;quotes&quot;</p>');
    expect(renderMarkdown('a\nb')).toBe('<p>a<br>b</p>');
  });
});
