import { writeFileSync } from 'node:fs';
import { it } from 'vitest';
import { DocxDocument } from '../src/model/index.js';
import { exportPdf } from '../src/pdf/index.js';
import { sampleDocx } from '../playground/sample.js';
import { testRegistry } from './helpers/fonts.js';

/** Writes the playground sample as PDF to test/.out for manual inspection. */
it.skipIf(!process.env.DOCXKIT_WRITE_SAMPLES)('exports the sample document', async () => {
  const doc = await DocxDocument.open(await sampleDocx());
  writeFileSync(new URL('./.out/sample.pdf', import.meta.url), await exportPdf(doc, { fonts: testRegistry() }));
});
