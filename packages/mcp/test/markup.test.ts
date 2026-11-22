import { describe, expect, it } from 'vitest';
import { DocxDocument, type DocParagraph } from 'docxkeun';
import { buildDocx } from '../../../test/helpers/docx.js';
import { hasMarkup, markupOf, parseMarkup, plainText } from '../src/markup.js';

const seg = (text: string, ...tags: string[]) => ({ text, tags: new Set(tags) });

async function paragraph(xml: string): Promise<DocParagraph> {
  const doc = await DocxDocument.open(await buildDocx({ body: xml }));
  return doc.body().blocks[0] as DocParagraph;
}

describe('formatting tags', () => {
  it('parses nested tags and escapes, leniently', () => {
    expect(parseMarkup('<b>Bold <i>both</i></b> plain')).toEqual([seg('Bold ', 'b'), seg('both', 'b', 'i'), seg(' plain')]);
    // A stray closing tag is ignored, an open one runs to the end, other "<" and "&" are text.
    expect(parseMarkup('a</b>b <u>c')).toEqual([seg('ab '), seg('c', 'u')]);
    expect(parseMarkup('x &lt;b&gt; &amp;lt; AT&T a < b')).toEqual([seg('x <b> &lt; AT&T a < b')]);
    expect(parseMarkup('m<sup>2</sup> H<sub>2</sub>O <s>old</s>')).toEqual([seg('m'), seg('2', 'sup'), seg(' H'), seg('2', 'sub'), seg('O '), seg('old', 's')]);
    expect(plainText('<b>PT</b> &amp; <i>Co</i>')).toBe('PT & Co');
    expect(hasMarkup('a < b & c')).toBe(false);
    expect(hasMarkup('<u>x</u>')).toBe(true);
  });

  it('tags runs, joins phrases Word split into runs, and keeps underlined blanks', async () => {
    const r = (text: string, rPr = '') => `<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ''}<w:t xml:space="preserve">${text}</w:t></w:r>`;
    // A bold space between two bold runs belongs to the phrase; an underlined gap stays underlined.
    const para = await paragraph(`<w:p>${r('Signed by ')}${r('PT', '<w:b/>')}${r(' ', '<w:b/>')}${r('Swadaya', '<w:b/>')}${r(': ')}${r('          ', '<w:u w:val="single"/>')}${r(' &lt;b&gt;')}</w:p>`);
    const text = markupOf(para);
    expect(text).toBe('Signed by <b>PT Swadaya</b>: <u>          </u> &lt;b>');
    expect(parseMarkup(text).map((s) => s.text).join('')).toBe('Signed by PT Swadaya:            <b>');
  });

  it('nests tags instead of repeating them', async () => {
    const r = (text: string, rPr = '') => `<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ''}<w:t xml:space="preserve">${text}</w:t></w:r>`;
    const para = await paragraph(`<w:p>${r('seluas ')}${r('6.337 M', '<w:b/>')}${r('2', '<w:b/><w:vertAlign w:val="superscript"/>')}${r(' (enam ribu)', '<w:b/>')}${r(' dan ')}${r('x', '<w:i/><w:u w:val="single"/>')}${r('y', '<w:u w:val="single"/>')}</w:p>`);
    expect(markupOf(para)).toBe('seluas <b>6.337 M<sup>2</sup> (enam ribu)</b> dan <i><u>x</u></i><u>y</u>');
  });
});
