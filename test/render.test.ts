import { describe, expect, it } from 'vitest';
import { layoutDocument } from '../src/layout/index.js';
import { DocxDocument } from '../src/model/index.js';
import { pageToSvg } from '../src/render/svg.js';
import { safeUrl } from '../src/render/viewer.js';
import { readComments } from '../src/comments/read.js';
import { buildDocx, p } from './helpers/docx.js';
import { haveSystemFonts, testRegistry } from './helpers/fonts.js';

describe('rendering safety', () => {
  it('allows only listed link schemes', () => {
    const allowed = ['http:', 'https:', 'mailto:'];
    expect(safeUrl('https://example.com/a?b=1', allowed)).toBe('https://example.com/a?b=1');
    expect(safeUrl('mailto:a@b.c', allowed)).toBe('mailto:a@b.c');
    expect(safeUrl('javascript:alert(1)', allowed)).toBeUndefined();
    expect(safeUrl('JaVaScRiPt:alert(1)', allowed)).toBeUndefined();
    expect(safeUrl('data:text/html,<script>', allowed)).toBeUndefined();
    expect(safeUrl('file:///etc/passwd', allowed)).toBeUndefined();
    expect(safeUrl('relative/path.docx', allowed)).toBeUndefined();
  });

  it.skipIf(!haveSystemFonts)('escapes document text in SVG output', async () => {
    const doc = await DocxDocument.open(await buildDocx({ body: p('&lt;script&gt;alert("x")&lt;/script&gt; &amp; "quotes"') }));
    const { pages } = layoutDocument(doc, { fonts: testRegistry() });
    const svg = pageToSvg(pages[0]!, { fontFamily: () => 'x"><script>', imageUrl: () => undefined });
    expect(svg).not.toContain('<script');
    expect(svg).toContain('&lt;script&gt;alert(&quot;x&quot;)&lt;/script&gt; &amp; &quot;quotes&quot;');
    expect(svg).toContain('font-family="x&quot;&gt;&lt;script&gt;"');
  });
});

describe('comments', () => {
  it('reads threads, resolved state and durable ids', async () => {
    const W15 = 'xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"';
    const comments =
      '<w:comment w:id="0" w:author="A" w:date="2026-01-01T00:00:00Z"><w:p w14:paraId="0000000A"><w:r><w:t>Root</w:t></w:r></w:p></w:comment>' +
      '<w:comment w:id="1" w:author="B"><w:p w14:paraId="0000000B"><w:r><w:t>Reply</w:t></w:r></w:p></w:comment>';
    const ext = `<?xml version="1.0"?><w15:commentsEx ${W15}><w15:commentEx w15:paraId="0000000A" w15:done="1"/><w15:commentEx w15:paraId="0000000B" w15:paraIdParent="0000000A"/></w15:commentsEx>`;
    const ids = '<?xml version="1.0"?><w16cid:commentsIds xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid"><w16cid:commentId w16cid:paraId="0000000A" w16cid:durableId="12345678"/></w16cid:commentsIds>';
    const doc = await DocxDocument.open(
      await buildDocx({
        body: p('x'),
        comments,
        rels: [
          ['rIdCx', 'http://schemas.microsoft.com/office/2011/relationships/commentsExtended', 'commentsExtended.xml'],
          ['rIdCi', 'http://schemas.microsoft.com/office/2016/09/relationships/commentsIds', 'commentsIds.xml'],
        ],
        parts: { 'word/commentsExtended.xml': { data: ext }, 'word/commentsIds.xml': { data: ids } },
      }),
    );
    const [root, reply] = readComments(doc);
    expect(root).toMatchObject({ id: '0', author: 'A', text: 'Root', resolved: true, durableId: '12345678' });
    expect(reply).toMatchObject({ id: '1', parentId: '0', resolved: false, text: 'Reply' });
  });
});
