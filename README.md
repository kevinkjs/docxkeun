# docxkit

Read, render, edit and convert `.docx` files in the browser (and Node), with **zero runtime dependencies**. Everything is implemented in this package: zip, XML, the document model, a page layout engine, an SVG viewer, a PDF writer, and the cryptography for password-protected files.

| Area | What you get |
|---|---|
| Read / write | Open `.docx` (also `.dotx`/`.docm`), edit through a code API, save. Parts you did not touch are written back byte-for-byte; unknown markup is preserved. |
| Render | Paginated, Word-like layout: styles, lists, tabs, tables, sections, columns, headers/footers, footnotes, page numbers, inline images. One SVG per page with real, selectable text. |
| PDF | Native PDF from the same layout: embedded font subsets, searchable text, links, bookmarks as named destinations, an outline from headings, comments as annotations. |
| Permissions | Word protection (read-only / comments / tracked changes / forms) with editable ranges, an app-level `AccessPolicy`, password-to-open `.docx` encryption, PDF passwords and permission flags. |
| Comments, anchors, links | Threaded comments (all five Word comment parts), bookmarks, internal and external hyperlinks, field-code links. |

## Quick start

```ts
import { openDocx, FontRegistry } from 'docxkit';
import { DocxViewer } from 'docxkit/viewer';

const fonts = new FontRegistry();
await fonts.addFromUrl('/fonts/Carlito-Regular.ttf');   // see "Fonts" below
await fonts.addFromUrl('/fonts/Carlito-Bold.ttf');

const viewer = new DocxViewer(document.getElementById('viewer')!, { fonts });
await viewer.load(fileFromInput);                       // Uint8Array | ArrayBuffer | Blob
viewer.on('linkclick', (e) => console.log(e.target));   // call e.preventDefault() to handle it yourself
viewer.scrollToAnchor('payment');                       // bookmarks
```

Without registered fonts the viewer still works: it measures the viewer's system fonts with the canvas API. Exact layout and PDF export need font files.

### Editing

```ts
import { DocxEditor } from 'docxkit/editor';

const editor = new DocxEditor(viewer.document!);       // the viewer re-renders on every change

editor.fillTemplate({ name: 'PT Contoh', amount: 'USD 1,200' }); // {{name}} even when split across runs, plus content controls by tag
const [match] = editor.find('royalty rate');
editor.formatRange(match, { bold: true, color: 'C00000' });
await editor.addHyperlink(match, { url: 'https://example.com/terms' });
editor.addBookmark(editor.find('Payment terms')[0], 'payment');
await editor.addHyperlink(editor.find('see payment')[0], { bookmark: 'payment' });

// Comments on the user's selection in the viewer
const range = editor.rangeFromSelection(viewer.getSelection()!);
const id = await editor.addComment(range, { author: 'Rina Pratama', text: 'Confirm with finance' });
await editor.replyToComment(id, { author: 'Budi', text: 'Confirmed.' });
await editor.resolveComment(id);

const bytes = await viewer.document!.save();
```

Also available: `insertText`, `appendText`, `replaceRange`, `replaceAll`, `deleteRange` (across paragraphs), `insertParagraph`, `deleteParagraph`, `setParagraphStyle`, `removeHyperlink`, `hyperlinks()`, `bookmarks()`, `removeBookmark`, `editComment`, `deleteComment`, `addEditableRange`, `removeEditableRange`.

### PDF

```ts
import { exportPdf } from 'docxkit/pdf';
import { pdfSecurity } from 'docxkit/security';

const pdf = await exportPdf(doc, { fonts, layout: viewer.layout });   // reuse the on-screen layout

const locked = await exportPdf(doc, {
  fonts,
  encryption: pdfSecurity({ userPassword: 'open-me', permissions: { copy: false, print: 'lowQuality' } }),
});
```

### Permissions

```ts
import { AccessPolicy, setProtection, readProtection } from 'docxkit/security';

// Word's own protection (stored in the file, honoured by Word)
await setProtection(doc, 'readOnly', { password: 'owner-secret' });
editor.addEditableRange(editor.find('[counterparty notes]')[0], { group: 'everyone' });
editor.addEditableRange(editor.find('Budi only')[0], { editor: 'budi@example.com' });

// App-level roles, combined with the document's protection
const policy = new AccessPolicy(doc, { user: { id: 'budi@example.com', name: 'Budi' }, role: 'editor' });
const guarded = new DocxEditor(doc, policy);       // every call is checked; throws PermissionError
policy.can('comment');                             // for enabling UI
new DocxViewer(el, { fonts, highlightRange: policy.highlightRange }); // shade what this user may edit
```

Roles: `viewer` (read, export), `commenter` (comments; edit/delete own), `editor` (edits, subject to document protection and editable ranges), `owner` (everything, including protection).

### Password-to-open `.docx`

```ts
import 'docxkit/security';                         // enables encrypted files in open/save

const doc = await openDocx(bytes, { password });   // PasswordRequiredError / WrongPasswordError
const encrypted = await doc.save({ password: 'new-password' });
```

Encryption is MS-OFFCRYPTO Agile (AES-256-CBC, SHA-512, 100,000 rounds, HMAC integrity), the same scheme Word uses. Office 2007 "Standard" encryption can be read.

### React

```tsx
import * as React from 'react';
import { createDocxViewerComponent } from 'docxkit/react';
const DocxView = createDocxViewerComponent(React);  // no React dependency in docxkit itself

<DocxView file={bytes} fonts={fonts} onReady={(viewer) => setViewer(viewer)} />
```

## Fonts

Word documents name fonts (Calibri, Times New Roman...) but rarely contain them, and browsers cannot read the system's font files. For exact layout and for PDF export, register font files with a `FontRegistry`:

- Office fonts map automatically to open, metric-compatible substitutes when those are registered: Calibri → Carlito, Cambria → Caladea, Arial → Liberation Sans, Times New Roman → Liberation Serif, Courier New → Liberation Mono (all SIL OFL).
- `registry.alias('Corporate Sans', 'Liberation Sans')` maps any other name.
- Fonts embedded in the document are supported: `await registerEmbeddedFonts(doc, registry)`.
- `registry.fallbacks = ['Carlito']` sets what unknown fonts use.

TrueType, OpenType (CFF) and `.ttc` collections are supported. PDF embedding respects a font's "restricted license" flag.

## Security model, honestly

- **Password-to-open `.docx` and PDF user passwords are real encryption** (AES-256).
- **Word document protection, editable ranges, PDF permission flags and `AccessPolicy` are restrictions, not encryption.** Word and PDF viewers honour them, but anyone can strip them from a file, and client-side checks can be bypassed. If your server stores documents, re-check permissions there before accepting changes.
- Documents are untrusted input. All text is escaped when rendered, external links only open with allowed schemes (`http:`, `https:`, `mailto:`, `tel:` by default), images referenced by URL are not loaded unless you opt in, and the XML parser never expands external entities.

## Status and limitations

Working and tested (77 tests): zip/XML/OPC round-trips, styles and numbering, layout and pagination, viewer, editing, comments, links, bookmarks, PDF export, document protection, access policy, `.docx` encryption, PDF encryption.

Verified against independent implementations so far:
- Crypto primitives against FIPS/RFC test vectors and Node's crypto.
- PDF output and AES-256 PDF encryption opened by Chrome's PDF viewer (with and without an open password).

**Still to verify against files made by Microsoft Word:** layout fidelity (line and page breaks), the Word protection password hash (the Excel-compatible low half is confirmed), and opening docxkit-encrypted files in Word / decrypting Word-encrypted files.

Not yet supported:
- Text wrapping around floating images and text boxes. Floating images are placed but text does not flow around them; DrawingML shapes and text boxes are not drawn.
- Right-to-left text and complex-script shaping (Arabic, Hebrew, Indic), ligatures and GPOS kerning.
- EMF/WMF, GIF and BMP images (drawn as placeholders in PDF), charts, SmartArt, equations (shown as plain text).
- Recording tracked changes. Existing insertions are shown and deletions hidden.
- Column balancing on continuous section breaks.

## Development

```bash
pnpm install
pnpm test        # vitest
pnpm dev         # playground at http://localhost:5199 (serves this Mac's system fonts in dev only)
pnpm build       # dist/
```

Tests that need fonts use the fonts installed on the machine (never committed) and are skipped when they are missing.
