/** The document the demo edits: a short sales update with a rough summary, a table and a list. */
import { createDocx } from 'docxkeun';
import { DocxEditor } from 'docxkeun/editor';

export async function sampleDocument() {
  const doc = await createDocx();
  const editor = new DocxEditor(doc);
  const title = editor.paragraphs()[0].el;
  editor.appendText(title, 'Q3 sales update');
  editor.setParagraphStyle(title, 'Heading1');
  editor.insertParagraph('end', 'Lumbung Coffee · for the October leadership meeting', { format: { italic: true, color: '595959' } });
  editor.insertParagraph('end', 'Summary', { style: 'Heading2' });
  editor.insertParagraph('end', 'This quarter our sales was continued to grow in all of the regions that we are operating in, and the online store have became the biggest channel of all the channels for the first time ever since we has launched it.');
  editor.insertParagraph('end', 'Sales by region', { style: 'Heading2' });
  const lead = editor.insertParagraph('end', 'Cups sold, in thousands:');
  editor.insertTable({ after: lead }, [
    ['Region', 'Q1', 'Q2', 'Q3'],
    ['Java', '120', '135', '152'],
    ['Sumatra', '80', '86', '97'],
    ['Bali', '45', '52', '71'],
  ], { header: true, widths: [2, 1, 1, 1] });
  editor.insertParagraph('end', 'Next steps', { style: 'Heading2' });
  const steps = ['Open more stores', 'Hire for the online store'].map((t) => editor.insertParagraph('end', t));
  await editor.setList(steps, { format: 'bullet' });
  return doc.save();
}
