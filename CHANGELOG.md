# Changelog

All notable changes to docxkeun and docxkeun-fonts. Versions follow [semantic versioning](https://semver.org); while the version is 0.x, minor versions may contain breaking changes.

## Unreleased

### docxkeun

- The AI's marks (`viewer.activity`) no longer slow the viewer down as they build up. Drawing them read each page's position, and the viewer's height for the change map, once per mark, right after adding the previous one, so the browser laid out the whole page again for every mark. In a long AI review (a 30-page contract with about 60 changes) redrawing the marks took about 110 ms after every change and kept growing until the page froze; it now reads each position once per pass and takes about 6 ms.
- Deleting a paragraph with Track Changes on now removes it outright only when its author inserted both the paragraph and all of its text. A paragraph someone else typed into is deleted as a tracked change, so rejecting the deletion brings their text back. Before, it disappeared without a trace whenever the paragraph mark was the author's own insertion.

### docxkeun chat

- **Instructions:** a new button opens a box for your own instructions to Claude (up to 4000 characters, for example "You are a legal assistant that helps me check documents for my clients."). They are added after the built-in instructions for every message, apply from your next message, and are remembered in the browser; opening another document keeps them. Locked regions and the document's protection still apply.

### docxkeun-mcp

- `--instructions <file>` gives the AI your own instructions, from a text or Markdown file of up to 4000 characters. They are added after the server's own instructions, and the AI app decides how to use them (Claude Code adds them to its system prompt). A file that is missing, unreadable or too long stops the server with one line saying why; edits to the file apply after the app reconnects.
- `createServer()` takes `instructions` (the text itself), and `MAX_INSTRUCTIONS` and `instructionsSection` are exported.
- A new `suggester` role (`--role suggester`, or `role: 'suggester'` in `createServer()`) for an AI whose every change a person reviews: it reads, comments, and changes text, lists and tables only as tracked changes signed with its author, whatever the document's own Track Changes setting. It cannot delete comments, accept or reject changes, switch tracking off, write to any file but the one it was given (its tools take no `output` or `overwrite`), or use tools whose edits are never tracked (pictures, charts, shapes and other objects, footnotes, links, bookmarks, headers and footers). It refuses to format text that carries another author's pending formatting change, since docxkeun would sign the change with that author's name. `SUGGESTER_OPERATIONS` and the `McpRole` type are exported.

## 0.4.0 (2026-11-24)

### docxkeun

- The `words` effect style now colours only the letters. Its shimmer and the marks it leaves are drawn on the text itself, so shaded table cells, highlighted text and pictures in the paragraph keep their own colours, and the changed words no longer sit in a tinted glow. Before, the shimmer was a coloured band blended over the line, which tinted anything that was not white, and the marks added a blurred box behind the words. With `marks: false`, the changed letters now light up for a moment instead of a tinted band passing over them.

- **Charts:** `DocxEditor.insertChart()` adds a column, bar, line, area, pie, doughnut or scatter chart (stacked and 100% stacked for column, bar and area), with a title, axis titles and legend. It looks like a chart Word has just inserted (theme colours, a white chart area without a border, 9 pt grey text), and its numbers are written into an embedded workbook, so "Edit Data" in Word shows them. `chartInfo()` reads a chart's data back as a table, and `updateChart()` changes its title, legend or data while keeping its look, and rewrites the workbook; an added series is styled like the last one in the next accent colour (keeping transparency, but not the last one's single-point formatting, trendline or error bars). A chart linked to an outside file only has its cached numbers changed. Word does not track chart changes, so neither does this. Number formats carry through: a series' `formatCode` (such as `0%`) and a chart's `categoryFormat` (such as `mmm-yy` for dates) are written to the chart and to the workbook's cells, `chartInfo()` reads them back, an update keeps each series' format by position, and numeric or date categories stay numbers instead of turning into text. So `categories` in `insertChart()`, `chartInfo()` and `updateChart()` holds strings, numbers (dates as Excel serial numbers) or `null` for a missing one; TypeScript code reading `chartInfo().categories` should allow for all three. The chart types and `ChartDataError` are exported from `docxkeun/editor`.
- **Objects:** `objects()` lists pictures, charts, diagrams, shapes, text boxes and groups with numbers. `resizeObject()`, `setObjectAlt()`, `deleteObject()`, `moveObject()` and `placeObject()` change them; resizing, alt text and floating position are applied directly, since Word does not track them, while creating, deleting, moving and switching between inline and floating are tracked. Before, only pictures could be inserted and given alt text. For an object that already floats, `placeObject()` keeps what the placement leaves out (x, y, `from`, `wrap`), so its wrap or reference can change on its own; making an inline object float still needs x and y.
- **Shapes and text boxes:** `insertShape()` adds a text box or a shape (rectangle, rounded rectangle, oval, triangle, diamond, arrows, chevron, pentagon, star, callouts, line), inline or floating (square wrap, above and below, in front of or behind the text), with text, fill and outline colours. As in Word, only pictures go inside a text box: `insertShape()`, `insertChart()` and `moveObject()` refuse shapes, text boxes, charts, SmartArt and other graphics there.
- **Breaking:** `paragraphs()` now also lists the paragraphs inside text boxes, right after the paragraph that holds the box. `find()`, `replaceAll()` and `fillTemplate()` reach text in boxes too. Code that indexes `paragraphs()` by position will see its numbers shift in documents with text boxes. Before, text in boxes could only be changed through `storyParts()`.
- Copied content (`copyBlocks()`) now gives a copied chart its own chart part and workbook (Word 2016 charts such as waterfalls too), so changing one copy no longer changes the other. `updateChart()` first gives a chart its own part when copies made by docxkeun 0.3.0 or older still share it.
- `replaceRange()` takes `{ keepObjects: true }`: pictures, charts, shapes and text boxes inside the replaced range then stay, right after the new text in their original order, and are not marked deleted with Track Changes on. It returns how many objects it kept. Without the option, objects in the range go with the old text, as when someone selects and types over them.
- The AI's marks (`aiActivity`) also notice changed charts and objects, not only text, and text changed inside a text box or shape (marked on the paragraph holding it).

### docxkeun chat

- Claude can now insert charts (column, bar, line, area, pie, doughnut, scatter), read and update their data, insert shapes and text boxes, and resize, move, describe or delete pictures, charts, shapes and text boxes. Text box paragraphs are numbered in line, so Claude edits them like any other paragraph. Before, it could only read these.

### docxkeun-mcp 0.3.0

- New tools: `insert_chart`, `read_chart`, `update_chart`, `insert_shape`, `edit_object` and `delete_object`. `read_document` and `list_objects` give every picture, chart, shape and text box an object number that these tools use. `read_chart` names number formats other than General (and says when the categories are dates), and `update_chart` keeps them.
- Text box paragraphs are now numbered in line, right after the paragraph holding the box, so `edit_paragraph`, `replace_text`, `format_text` and the other paragraph tools work on them. Paragraph numbers after a text box shift when a box with text is added, moved or deleted; read the document again after those tools. `replace_story_text` still works.
- `edit_object` changes a floating object's wrap or reference on its own: x and y left out keep its position. `list_objects` says where a floating object sits and how text wraps around it ("floating, 36 pt right and 12 pt down from the margin, square wrap").
- `edit_paragraph` and `replace_text` keep the pictures, charts and text boxes that sat in the middle of the text they rewrite, right after the new text, and say how many they kept. Before, rewriting such a paragraph deleted them, text box and all.
- The live preview shows tracked changes as Word's All Markup view does: inserted text underlined and deleted text struck through, in a colour per author. Before, it always showed documents as if every change were accepted, so an AI's tracked edits looked like plain edits. A **Tracked changes** checkbox switches back to that view, and the browser remembers it. The AI's effects and the selection chip find the right paragraphs in both views, and `get_selection` reads a selection made where deleted text shows, saying when the text includes deleted words. Activity events now also carry read_document's paragraph numbers (`read`).

## 0.3.0 (2026-11-23)

### docxkeun

- **AI effect styles:** `aiActivity: { style }` and `viewer.activity.style` pick how the paragraph the AI is changing looks: `highlight` (the band, as before), `words` (the letters themselves shimmer), `underline`, `edge`, `spotlight`, `quiet` or `sparkles`. `ACTIVITY_STYLES` lists them with labels for a picker. The mark that stays after a change follows the style: a soft tint on the paragraph (`highlight`), glowing changed words (`words`), underlined changed words (`underline`), or the ✦ margin bar (the others); `marks: false` / `viewer.activity.marksShown` turn the marks off. Every style has a still form for `animate: false` and `prefers-reduced-motion`.
- The AI's marks now build up over a run of changes instead of showing only the latest one: each follows its paragraph as others are added or removed, a paragraph changed twice is marked against its text before the first change, and `clearMarks()` starts over. `ParagraphChange.from` gives a changed paragraph's index in the old text.
- The AI's marks and effects now follow zoom changes and the comment column's width (they kept their old positions before), and the change map's ticks show the text from before.
- **Tables:** `DocxEditor.insertTable()` adds a table that spans the text width, with an optional bold header row that repeats on each page, relative column widths and alignment. It uses Word's "Table Grid" style when the document has it, or plain borders. `insertTableRows()` adds rows formatted like the row they copy, `deleteTableRows()` and `deleteTable()` remove them, and `tableOf()` finds the table a cell paragraph is in. All of them record tracked changes when Track Changes is on, and rejecting an inserted table removes it whole.
- **Copying content:** `DocxEditor.copyBlocks()` copies paragraphs and tables with their formatting, lists and pictures, optionally starting on a new page. Lists in the copy start again at 1 (or continue with `restartLists: false`). Bookmarks, comment anchors and note references are left out and counted, and ids that must be unique are renewed.
- `DocxEditor.replaceRange()` also takes pieces of text, each formatted like a run you point to (`{ text, like, format }`), so rewritten text can keep which words were bold, underlined or coloured.
- `setList` with `sameListAs` now accepts the joined list's own `format`, and its errors say what to do instead: join a paragraph that is not in a list yet by making it one first; use `start` without `sameListAs` to number from a given value.

### docxkeun chat

- "Appearance" also picks the effect style and whether marks stay after each change, with a live sample. A reply's marks cover all its changes; opening a document or undoing a reply clears them.
- Claude now sees which words are bold, italic, underlined, struck through, superscript or subscript, and keeps them when it rewrites or translates text (it used to flatten a rewritten paragraph to one formatting).
- Claude can now copy content with its formatting and make and change tables, so "translate this document onto a new page" keeps the tables and lists instead of retyping them as plain text.

### docxkeun-mcp 0.2.0

- The live preview's header has a **Style** list and a **Marks** switch next to "AI effects", remembered in the browser. Marks build up while the AI works on a document and start over when you open another one.
- New tools: `copy_paragraphs` copies paragraphs and tables with their formatting and lists (for example onto a new page, to translate a document), `insert_table` makes a table, and `edit_table` adds or deletes rows or deletes a table.
- `set_list` explains how to join paragraphs that are not next to each other into one list, and accepts the list's own format with `sameListAs`.
- **Text formatting is visible and kept.** `read_document` tags text formatted differently from the rest of its paragraph (`<b>`, `<i>`, `<u>`, `<s>`, `<sup>`, `<sub>`) and names a paragraph's alignment and whole-paragraph formatting in its description, e.g. `[1] (centered; bold, underlined) AGREEMENT`. `edit_paragraph`, `insert_paragraphs`, `create_document` and `replace_text` take the same tags. When `edit_paragraph` rewrites a paragraph, each tagged part copies the formatting (font, size and colour too) of the matching phrase in the original, in order. Before, rewriting a paragraph gave all of it the first run's formatting, so bold names and amounts were lost in translations and small edits alike.
- New tool `copy_document` copies a document to a new file exactly as it is (an encrypted one stays encrypted), so "make a copy" or "translate it into a new file" keeps every bit of formatting. Before, the AI could only retype the text into `create_document`, which made a blank document and lost the text and paragraph formatting, lists and tables. `create_document`'s description now sends copies to `copy_document`.

## 0.2.0 (2026-11-22)

### docxkeun

- List editing: `DocxEditor.listItems()` returns each list paragraph's label as Word shows it (`"1."`, `"b)"`, `"•"`), and `DocxEditor.setList()` switches paragraphs between bullets, numbers, letters and roman numerals, restarts numbering, changes the list level, joins an existing list or removes a list. A changed list keeps the indentation and label formatting of the list it came from. `numbering.xml` is created when a document has none.
- **Tracked changes:** `new DocxEditor(doc, guard, { trackChanges: true, author })` records edits as Word revisions (insertions, deletions, formatting and paragraph changes, paragraph marks), signed and dated. Typing inside your own insertion extends it; deleting it removes it. `revisions()`, `acceptRevisions()` and `rejectRevisions()` (all, by id or by author) review them, including moves made in Word, and `setTrackChangesSetting()` turns Word's own Track Changes on or off. Edits follow the document's setting by default. Documents protected for tracked changes now accept edits (recorded) and refuse review. With `showDeletions`, insertions are underlined and deletions struck through in an author colour.
- **Drawings:** DrawingML shapes (153 preset shapes and custom geometry, theme fills and outlines, arrowheads, rotation and flips), text boxes (laid out as Word paragraphs, old VML ones too), groups and drawing canvases, charts (column, bar, line, area, pie, doughnut, scatter, bubble, radar, combinations, with axes, legends, data labels and number formats) and SmartArt (from Word's saved drawing) are drawn on screen and in PDFs.
- **Pictures:** GIF, BMP and TIFF are decoded (TIFF is shown in the browser as PNG; all three are embedded in PDFs), and EMF/WMF pictures are drawn as vectors. Pictures can be rotated and flipped. `DocxEditor.insertImage()` adds a picture at its natural size or a given one.
- **Equations:** Office Math is typeset: fractions, scripts, radicals, big operators with limits, delimiters that grow, matrices, accents, equation arrays and more, inline or on their own line. Text extraction gives Office's linear format (`x=(-b±√(b^2-4ac))/(2a)`).
- **Text wrapping:** text flows around floating pictures, shapes and text boxes (square, tight, through, top and bottom), on the paragraphs they meet.
- **Columns:** a multi-column section that ends with a continuous section break is balanced, as in Word.
- **Right-to-left text:** the Unicode bidi algorithm orders mixed Arabic, Hebrew and Latin text; right-to-left paragraphs (`w:bidi`) are laid out from the right margin with their indents and alignment mirrored. `DocxEditor.formatParagraph()` sets alignment, direction, indents, spacing and keep rules.
- **Headers, footers and notes:** `DocxDocument.storyParts()` lists header, footer and note parts, and `new DocxEditor(doc, guard, { part })` edits one. `DocxEditor.addFootnote()` adds a footnote (and the footnotes part when there is none).
- **In-page editing:** double-click a header or footer, or click a footnote or text box, to type in it. A formatting toolbar (`toolbar: true`), rich copy and paste (HTML from Word, Google Docs and web pages), Cmd/Ctrl+E/L/R/J alignment, Track Changes while typing, and a caret that steps over pictures. Layout re-uses the lines of unchanged paragraphs (`LayoutCache`, used by the viewer), so typing in long documents is several times faster.
- **AI assistants:** `viewer.selectionInfo()` and the `selectioninfo` event describe what the user selected (or where the caret is) with read_document's paragraph numbers, offsets, pages and text; `describeSelection()` turns it into a prompt line. `viewer.activity` shows an AI at work: a shimmer over the paragraphs it is changing and its own named caret, a scan line while it reads, then the change inked in with a fading glow, removed paragraphs floating away, change bars with ✦ in the margin and a change map on the right edge (`work`, `snapshot`, `reveal`, `idle`, `clearMarks`; follows `prefers-reduced-motion`). Optional: `aiActivity: false`, or `{ animate: false }` for highlights without motion and `{ marks: false }` without change bars; `viewer.activity.enabled` / `.animate` / `.marksShown` at run time. `{ name, color }` (and `viewer.activity.name` / `.color`) set the name on the caret and the colours, one CSS colour or two for a gradient; from CSS, `--docxkeun-ai-color-1` / `-2`. The caret's label turns dark on light colours.
- **Adding headers and footers:** `DocxEditor.setHeader()` / `setFooter()` create or replace them, for every section or one, on every page, the first page (w:titlePg) or even pages (Word's even/odd setting). Tabs split a line into left, centre and right parts, and `{PAGE}`, `{NUMPAGES}`, `{SECTIONPAGES}` become page number fields. A part shared with other sections is not overwritten, Track Changes records the text changes, and `removeHeader()` / `removeFooter()` delete parts nothing uses any more. `headersAndFooters()` lists them with their sections and text.
- New exports: `describeChart`, `readChart`, `renderChart`, `mathToText`, `layoutMath`, `renderMetafile`, `sniffImage`, `decodeRaster`, `imageInfo`, `LayoutCache`, `FormatToolbar`, `htmlToParagraphs`.
- **Viewer UI:** one shared look and frame for four optional parts. `toolbar` (fold to a "Show toolbar" strip, overflow menu), `navigation` (240px panel: back/forward, Outline, Bookmarks and Links tabs, search, per-tab previous/next, folds to a 36px rail), `comments` (one column of compact cards, margin bubbles on narrow viewers with `mode: 'auto'`, Reply/Resolve/Delete, "Show resolved") and `contextMenu` (right-click menu that follows what was clicked; Shift+right-click shows the browser's). Each has a controller on the viewer (`viewer.toolbar`, `.navigation`, `.comments`, `.contextMenu`), the `chromechange` event reports changes and nothing is stored. Theme with `--docxkeun-ui-bg`, `-fg`, `-muted`, `-line`, `-hover`, `-accent`, `-accent-soft`, `-radius`, `-shadow`, `-font` (the `--docxkeun-toolbar-*` variables still work). `scrollToComment(id)` now opens the thread, and `viewer.comments.width` is what fit-to-window apps subtract.
- **Undo for everything:** `viewer.editing.transaction(label, fn)` makes an action that touches several parts (comments, relationships, headers, numbering, styles) one undo step; a failure restores the document and emits `editerror`. Inside `fn`, changes go through the `editor` it receives; edits requested meanwhile wait their turn. `DocxDocument.checkpoint()` / `restore()` back it. Fixes ⌘Z after adding a comment or link leaving markers half-removed. Undo keeps 200 steps.
- **Editor:** `setImageAlt()`, `setHyperlinkTarget()` (URL or bookmark, same scheme rules as `addHyperlink`) and `clearFormatting()`.
- **Breaking (0.x):** `DocxEditor.removeHyperlink()` is now `async`; call it with `await`. It also drops the link's relationship when nothing else uses it, and a guard refusal rejects the promise instead of throwing.
- New exports (`docxkeun/viewer`): `NavigationController`, `NavTab`, `NavigationOptions`, `CommentsController`, `CommentThread`, `CommentsMode`, `CommentsOptions`, `relativeTime`, `filterTargets`, `currentHeading`, `ToolbarController`, `ContextMenuController`, `hitTarget`, `defaultItems`, `parseLinkTarget`, `MenuItem`, `MenuHandle`, `MenuContext`, `MenuTarget`, `MenuTargetKind`, `MenuEnv`, `MenuActions`, `ContextMenuOptions`, `ToolbarOptions`, `FormatChoice`, `ParagraphChoice`, `openMenu`, `promptPopover`.
- `viewer.navigation` is now a class instance (`NavigationController`). `const { headings, bookmarks, links } = viewer.navigation` works as before, but its lists are getters, so `{ ...viewer.navigation }` and `JSON.stringify(viewer.navigation)` no longer include them.

### docxkeun chat (new app, not published)

- A web app in `packages/chat`: open a `.docx`, chat with Claude about it with your own API key, and watch the changes appear live in the docxkeun viewer. Undo per reply, selection as context, download and PDF export.
- The viewer shimmers the paragraphs Claude is about to change while it is still writing the tool call (from the streamed tool input), then animates the change. The selection (or the caret's paragraph) goes with the next message, with paragraph numbers and page.
- The document has the viewer's navigation rail and comment cards, and the right-click menu has "Ask Claude about this", which puts the selection into the message box.
- "Appearance" sets the name on Claude's caret and its colours, remembered in the browser.
- The Claude calls go through a small relay on your computer (`pnpm dev:chat`, or `pnpm start:chat` for the build), which keeps the key in memory: it works for organizations that block browser calls to the API, and the browser never stores the key. The relay answers only its own page and forwards only Messages and Models calls.

### docxkeun-mcp 0.1.0 (new package)

- An MCP server for Word documents, so Claude Code, Claude Desktop and other MCP clients can work with `.docx` files through docxkeun.
- Tools to list, read, search and inspect documents; find and replace; fill templates; edit, insert and delete paragraphs; format text; change lists (bullets, numbers, letters, levels); add links and bookmarks; add, reply to, resolve and delete comments; create documents; lock regions; set protection and open passwords; and export PDFs.
- Only files inside the folders given on the command line can be read or written (symlinks and `..` cannot escape them). Writes go through a temporary file.
- `--preview` runs a live preview on 127.0.0.1: docxkeun's viewer redraws a document whenever a tool or anything else changes it, with a note of what changed. Read-only, local only, and it checks the Host header against DNS rebinding. `--preview-only` runs just the preview, to watch a folder.
- `--role viewer|commenter|editor|owner` decides which tools the AI gets, and docxkeun's `AccessPolicy` enforces it together with the document's protection and locked regions.
- The document operations behind the tools are exported (`OPERATIONS`, `operationsFor`, `contextFor`) for building other front ends, such as docxkeun chat.
- Tools for tracked changes (`list_revisions`, `accept_revisions`, `reject_revisions`, `set_track_changes`), headers, footers, notes and text boxes (`read_headers_footers`, `replace_story_text`), pictures (`insert_image`, `list_objects`), footnotes (`add_footnote`) and paragraph formatting (`format_paragraph`, including right-to-left). `read_document` shows pictures, charts (with their data), shapes, equations, text direction and tracked changes. Edits are signed with `--author` when Track Changes is on, and documents protected for tracked changes can now be edited.
- `set_header_footer` and `remove_header_footer`: headers and footers with page numbers, left / centre / right parts, first-page and even-page versions, per section. `read_headers_footers` shows page fields as `{PAGE}` and which pages and sections each one is for.
- The live preview's "AI effects" menu: animated, still or off. The AI's caret carries the `--author` name, and `--ai-color` sets its colours (one CSS colour, or two for a gradient).
- `get_selection` (and the resource `docxkeun://selection`): what the user selected in the live preview, so "this" and "here" mean something. The preview's header shows what is shared. Pages report positions only, with a per-server token; the text comes from the file.
- The live preview shows where the AI works while a tool runs (shimmer and caret, or a scan line for reads) and animates each change when it lands.
- The live preview has the navigation rail and comment cards (margin bubbles in a narrow pane), and a read-only right-click menu with Copy, Copy link address and Open link.
- It needs docxkeun 0.2.0.

## 0.1.3 (2026-11-02)

Documentation only. No code changes.

- The README shows a screenshot of the live demo.
- docxkeun-fonts is unchanged and stays at 0.1.1.

## 0.1.2 (2026-11-02)

Package metadata only. No code changes.

- A live demo at https://kevinkjs.github.io/docxkeun/, now the package homepage and linked from the README.
- More descriptive package description and search keywords.
- docxkeun-fonts is unchanged and stays at 0.1.1.

## 0.1.1 (2026-11-01)

First release published from GitHub Actions. No code changes.

- The packaged changelog now shows the 0.1.0 release date (0.1.0 on npm says "unreleased").
- docxkeun-fonts: version aligned with docxkeun; the fonts are unchanged.

## 0.1.0 (2026-11-01)

First public release (beta).

### docxkeun

- Read and write `.docx` with zero runtime dependencies: own zip, XML, OPC and document model. Parts you do not change are written back byte-for-byte.
- Word-like layout: styles, lists, tabs, tables (merged cells, repeated header rows), sections, columns, headers/footers, footnotes, page numbers, inline images.
- Browser viewer: one SVG per page with selectable text, comment sidebar, links, bookmarks, zoom.
- Typing directly on the pages: caret, selection, IME input, clipboard (plain text), shortcuts, undo/redo.
- Code API: find/replace across runs, `{{placeholder}}` and content-control filling, formatting, paragraphs, hyperlinks, bookmarks, threaded comments.
- Navigation between headings, bookmarks and hyperlinks, with back/forward history.
- Locked regions (Word content controls that cannot be edited) and editable ranges.
- Word document protection with passwords, an `AccessPolicy` for app roles, and password-to-open encryption (MS-OFFCRYPTO Agile, AES-256).
- Native PDF export from the same layout: embedded font subsets, searchable text, links, outline, comments, AES-256 encryption with permission flags.
- React binding without a React dependency.

### docxkeun-fonts

- Carlito, Caladea and Liberation Sans/Serif/Mono (SIL OFL 1.1), metric-compatible with Calibri, Cambria, Arial, Times New Roman and Courier New.
