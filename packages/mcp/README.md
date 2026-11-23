# docxkeun-mcp

An [MCP](https://modelcontextprotocol.io) server for Word documents. It lets Claude (Claude Code, Claude Desktop) and other MCP clients read, edit, comment on, protect and convert `.docx` files to PDF.

It is built on [docxkeun](https://github.com/kevinkjs/docxkeun), so it needs no Word, no LibreOffice and no cloud service: the documents are opened, changed and converted on your machine. It only touches files in the folders you give it, and the AI's role (viewer, commenter, editor or owner) is enforced along with the document's own protection and locked regions.

## Set up

Node.js 20.12 or newer is needed.

**Claude Code**

```bash
claude mcp add docxkeun -- npx -y docxkeun-mcp ~/Documents/contracts
```

**Claude Desktop**: add this to `claude_desktop_config.json` (Settings → Developer → Edit Config), then restart Claude:

```json
{
  "mcpServers": {
    "docxkeun": {
      "command": "npx",
      "args": ["-y", "docxkeun-mcp", "/Users/you/Documents/contracts"]
    }
  }
}
```

Claude Desktop starts servers without your shell's `PATH`; if it cannot find `npx`, use its full path (`which npx`).

**Other clients**: any MCP client that starts stdio servers can run `npx -y docxkeun-mcp <folder>`.

Then ask things like:

- "Read contract.docx and comment on anything unusual in the payment terms."
- "Fill in offer-letter.docx for Dana Putri, start date 1 December, and save it as offer-dana.docx."
- "Replace ACME Ltd. with Contoso GmbH everywhere in the agreements folder."
- "Change the a., b. items in the transfer letter to 1., 2."
- "Lock the signature block, protect the document so others can only comment, and export a PDF."

## Options

```
docxkeun-mcp [options] [folder ...]
```

| Option | Meaning |
|---|---|
| `folder ...` | The folders the server may read and write (default: the current folder). Subfolders are included. |
| `--role <role>` | What the AI may do: `viewer`, `commenter`, `editor` (default) or `owner`. |
| `--author <name>` | The name on comments the AI adds, and on its caret in the live preview (default "AI assistant"). |
| `--ai-color <color>` | The colour of the AI's caret, shimmer and change marks in the live preview: a CSS colour, or two separated by a comma for a gradient (`"#0b57d0"` or `"#0b57d0,#1e8e3e"`). |
| `--fonts <folder>` | Extra `.ttf`/`.otf` fonts for page counts and PDF export, e.g. Microsoft Office's fonts for exact Calibri layout. Repeatable. |
| `--preview` | Also run the [live preview](#live-preview) at `http://127.0.0.1:5230` (or the next free port). |
| `--preview-port <n>` | The first port to try for the preview. |
| `--preview-only` | Run only the live preview, without the MCP server, until you stop it: for watching a folder while you edit in Word. |

## Live preview

With `--preview`, the server also shows your documents in a web page, drawn by docxkeun's own viewer: pages, comments and list labels. The page redraws a document whenever it changes, whether through a tool or in Word, and shows a note of what the AI did ("Added comment 3 on …"). New comments scroll into view. A rail on the left lists headings, bookmarks and links; comments show as cards beside the page (small bubbles in a narrow pane). Right-click offers Copy, Copy link address and Open link, and Shift+right-click gives the browser's menu.

```bash
claude mcp add docxkeun -- npx -y docxkeun-mcp --preview ~/Documents/contracts
```

Ask Claude for the preview link (the `get_preview_link` tool), or open `http://127.0.0.1:5230`. In the Claude desktop app, open it in the browser pane next to the chat. "Follow changes" switches to whichever document the AI is editing, and "Fit width" keeps the page and its comments inside a narrow pane.

To watch a folder without an AI client, run `npx docxkeun-mcp --preview-only ~/Documents/contracts`; it redraws documents whenever they are saved.

The "AI effects" menu in the header picks how the AI's work shows: animated, still (highlights without motion) or off. **Style** picks the effect (highlight band, word shimmer, underline wave, edge glow, spotlight, quiet or sparkles), and **Marks** whether a mark stays on what the AI changed. The browser remembers all three.

While a tool runs, the preview shows where the AI is working: a shimmer over the paragraphs and its own caret ("Claude is editing…"), or a scan line while it reads. When the change lands, new text inks in, removed paragraphs float away, and change bars with a ✦ (hover: what the text was) stay in the margin until the next change.

**Selections.** Select text (or click in a paragraph) in the preview, and the AI can ask what you mean by "this" with `get_selection`: it answers with the document, paragraph numbers, page and the text in «», read from the file. In Claude Code you can also @-mention the `docxkeun://selection` resource. A chip in the preview's header shows what is shared.

The preview only listens on 127.0.0.1, only answers requests addressed to that host (so other websites cannot read it through DNS rebinding), serves only the Word documents in the allowed folders, and is read-only. Pages report their selection with a token only the preview's own page has, and only as positions: the text the AI sees always comes from the file. Encrypted documents are not shown, since the preview has no password for them. It stops when the MCP client closes the server.

## Tools

Each role gets only the tools it may use.

| Tool | What it does | Role |
|---|---|---|
| `list_documents` | Word documents in the allowed folders | viewer |
| `read_document` | Numbered paragraphs with styles, list labels, tables, locks, comments, tracked changes, text direction, and the pictures, charts, shapes and equations in them | viewer |
| `get_document_info` | Properties, page count, protection, encryption, locked regions, styles | viewer |
| `find_text` | Matches (text or regex) with paragraph numbers and context | viewer |
| `list_comments` | Comment threads with anchors and replies | viewer |
| `list_links` | Hyperlinks and bookmarks | viewer |
| `list_objects` | Pictures, charts (with their data), SmartArt, shapes, text boxes and equations | viewer |
| `read_headers_footers` | Headers, footers, footnotes, endnotes and text boxes, with ids for editing | viewer |
| `list_revisions` | Tracked changes with ids, authors, dates and paragraphs | viewer |
| `export_pdf` | PDF with embedded fonts, links, outline and comments; optional password and permissions | viewer |
| `get_preview_link` | A link to the live preview of a document (only with `--preview`) | viewer |
| `get_selection` | What the user selected (or clicked) in the live preview: paragraph numbers, page and text. Also the resource `docxkeun://selection` (only with `--preview`) | viewer |
| `add_comment`, `reply_to_comment`, `resolve_comment`, `delete_comment` | Comments (a commenter may delete only its own) | commenter |
| `create_document` | A new, blank document from paragraphs | editor |
| `copy_document` | An exact copy of a document in a new file (formatting, tables, lists, pictures, headers, footers, comments), for example to translate the copy | editor |
| `replace_text` | Find and replace, keeping formatting; regex with `$1` | editor |
| `fill_template` | `{{placeholder}}` fields and content controls | editor |
| `edit_paragraph`, `insert_paragraphs`, `delete_paragraphs` | Paragraph text and styles | editor |
| `copy_paragraphs` | Copy paragraphs and tables with their formatting and lists, optionally onto a new page (for example to translate a document); lists in the copy start again | editor |
| `insert_table`, `edit_table` | Make a table (header row, column widths, borders, alignment); add or delete rows, or delete a table. Cells are paragraphs, so the text tools edit them | editor |
| `insert_chart`, `update_chart`, `read_chart` | Make a chart (column, bar, line, area, pie, doughnut, scatter; stacked) with its data in an embedded workbook, change its title, legend or data, or read its data as a table | editor (`read_chart`: viewer) |
| `insert_shape` | A text box or a shape (rectangle, oval, arrow, star, callout, line…), inline or floating, with text, fill and outline | editor |
| `edit_object`, `delete_object` | Resize, move, re-place or describe (alt text) a picture, chart, shape or text box, or delete it | editor |
| `format_text` | Bold, italic, underline, colour, highlight, size, font | editor |
| `format_paragraph` | Alignment, right-to-left direction, indents, spacing, line spacing, keep and page-break rules | editor |
| `replace_story_text` | Edit text in a header, footer, footnote, endnote or text box | editor |
| `set_header_footer`, `remove_header_footer` | Add, replace or remove headers and footers: page numbers (`{PAGE}` of `{NUMPAGES}`), left / centre / right parts, first-page and even-page versions, per section | editor |
| `insert_image` | A picture file (PNG, JPEG, GIF, BMP, TIFF, EMF, WMF) from the allowed folders | editor |
| `add_footnote` | A footnote at a place in a paragraph | editor |
| `set_track_changes` | Turn Word's Track Changes on or off; while on, every edit is a tracked change signed with `--author` | editor |
| `accept_revisions`, `reject_revisions` | Review tracked changes by id, by author, or all | editor |
| `set_list` | Bullets, numbers (1. 1)), letters (a. A.) or roman numerals; restart numbering, change level, join paragraphs that are not next to each other into one list, or leave a list | editor |
| `add_hyperlink`, `add_bookmark` | Links to web addresses or bookmarks | editor |
| `lock_region`, `unlock_region` | Lock text or paragraphs against editing (Word content controls) | owner |
| `protect_document` | Word's editing restrictions: read-only, comments only, tracked changes only, forms only | owner |
| `set_open_password` | Encrypt with a password to open (AES-256), or remove it | owner |

`read_document` shows text formatting as tags: text formatted differently from the rest of its paragraph is wrapped in `<b>`, `<i>`, `<u>`, `<s>`, `<sup>` or `<sub>`, and a paragraph's alignment and whole-paragraph formatting are in its description (`[1] (centered; bold, underlined) AGREEMENT`). `edit_paragraph`, `insert_paragraphs`, `create_document` and `replace_text` take the same tags, so an AI that rewrites or translates a paragraph keeps its bold, underlined and other formatted words; each tagged part copies the font, size and colour of the matching phrase. Searches ignore the tags.

Objects (pictures, charts, shapes, text boxes) have numbers in `read_document` and `list_objects`, which the object tools take. The paragraphs of a text box are numbered in line, right after the paragraph holding the box, so the paragraph tools edit them; numbers after a box shift when one with text is added, moved or deleted, so read the document again.

Editing tools save to the file straight away; every one of them takes `output` to write a copy instead. Encrypted documents take `password`, and stay encrypted when they are changed.

## Safety

- **Folders.** Paths outside the given folders are refused, including through symlinks and `..`. Only Word documents (`.docx`, `.docm`, `.dotx`, `.dotm`) and PDFs (for export) are read or written. The server refuses to run with the whole file system as its folder.
- **Roles and protection.** Every change goes through docxkeun's `AccessPolicy`. A commenter cannot change text; only an owner can lock regions, change protection or set passwords. A document's own protection (for example "comments only") applies to every role except owner, and locked regions apply to everyone.
- **Writes.** Files are written to a temporary file and renamed, so a failed write never leaves a half-written document. Keep backups or use version control for documents you let an AI edit, and close a document in Word before an AI edits it, since Word does not notice changes made to an open file.
- **Your data and the AI.** The server itself sends nothing over the network, but your MCP client passes tool results (document text, comments) to its AI model. Passwords you give the AI to open a document are part of that conversation.
- **Approval.** MCP clients such as Claude Code and Claude Desktop ask before running a tool unless you allow it; tools that change files are marked as such.

## Limitations

The tools edit the body (including tables), headers, footers, footnotes, endnotes and text boxes, insert pictures and footnotes, and record and review tracked changes. A document protected for tracked changes can be edited (every edit is recorded) but not reviewed; only its owner can accept or reject changes after removing the protection. Charts, shapes and text boxes can be made and changed, and pictures resized and moved; SmartArt and equations can be read but not created or changed. Page counts and PDFs come from docxkeun's own layout and may differ slightly from Word's. See [docxkeun's limitations](https://github.com/kevinkjs/docxkeun#status-and-limitations).

## Use from code

```ts
import { createServer } from 'docxkeun-mcp';
import { StdioServerTransport } from '@modelcontextprotocol/server/stdio';

const server = await createServer({ folders: ['/srv/documents'], role: 'commenter', author: 'Review Bot' });
await server.connect(new StdioServerTransport());
```

`createServer` returns an `McpServer` from the [MCP TypeScript SDK](https://github.com/modelcontextprotocol/typescript-sdk), so it works with any of its transports.

## License

[MIT](LICENSE) © Kevin Jeremia
