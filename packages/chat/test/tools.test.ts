/** The chat's tools: definitions sent to the API, and running them on the open document. */
import { describe, expect, it } from 'vitest';
import { DocxDocument } from 'docxkeun';
import { DocxEditor } from 'docxkeun/editor';
import { buildDocx, p } from '../../../test/helpers/docx.js';
import { renderMarkdown } from '../src/markdown.js';
import { createToolbox } from '../src/tools.js';

const NUMBERING =
  '<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="%1."/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum>' +
  '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>';
const item = (t: string) => p(t, '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>');

describe('chat tools', () => {
  it('offers the editor operations with clean JSON schemas', () => {
    const { definitions } = createToolbox(() => undefined);
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
    const toolbox = createToolbox(() => ({ doc, name: 'letter.docx' }));
    const result = await toolbox.run('set_list', { from: 2, to: 3, format: 'decimal' });
    expect(result).toEqual({ content: 'Paragraphs 2–3 are now labelled "1.", "2.".', isError: false, kind: 'change' });
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
    const toolbox = createToolbox(() => ({ doc, name: 'x.docx' }));
    const locked = await toolbox.run('edit_paragraph', { paragraph: 1, text: 'Changed' });
    expect(locked).toMatchObject({ isError: true });
    expect(locked.content).toContain('"Agreed" is locked against editing');
    expect((await toolbox.run('nope', {})).content).toBe('There is no tool named nope.');
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
