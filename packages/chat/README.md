# docxkeun chat

A small web app for working on a Word document with Claude. Open a `.docx`, ask Claude to summarise, fix, comment on or restructure it, and watch every change appear in the document as it happens: docxkeun's own viewer redraws the pages, lists and comments after each edit.

The page runs in your browser; the Claude API calls are made by a small server on your computer (the relay), which holds your API key in memory. Organizations whose data-retention settings block API calls from browsers allow these, and your browser never stores the key.

## Run it

```bash
pnpm install
pnpm dev:chat
```

Open the address it prints (http://localhost:5260) and paste an API key from [console.anthropic.com](https://console.anthropic.com/settings/keys). Then open a `.docx` (or the sample) and start asking.

To skip the key prompt, start the server with the key in the environment:

```bash
ANTHROPIC_API_KEY=sk-ant-... pnpm dev:chat
```

To run the built app instead of the dev server:

```bash
pnpm build:chat
```

```bash
pnpm start:chat
```

## What it does

- **Live document:** the right side is docxkeun's viewer. Claude's edits show up there immediately; new comments scroll into view. You can also type in the document yourself. Headings, bookmarks and links are in the navigation rail on the left (click it to open), comments are cards beside the page, and right-clicking selected text offers "Ask Claude about this", which puts it in the message box.
- **Point at text:** select text (or click in a paragraph) and Claude knows what "this" or "here" means: the message carries the paragraph numbers, page and the selected text. The chip above the input shows it; × leaves it out.
- **See Claude work:** while Claude is still writing a tool call, the paragraphs it is about to change shimmer and its caret ("Claude is editing…") goes there; reads show a scan line. When the change lands, new text inks in, removed paragraphs float away, and change bars with a ✦ (hover: the text before) mark everything the reply touched, with ticks on the right edge to jump to them. The "Effects" menu switches this to still highlights or off, and "Appearance" changes the name on the caret, its colours, the effect (a highlight band, the words themselves shimmering, an underline wave, an edge glow, a spotlight, a quiet margin bar or sparkles) and whether a mark stays on what changed (all remembered in the browser).
- **Instructions:** your own instructions for Claude (for example "You are a legal assistant…"), added to its built-in ones for every message and remembered in the browser. Locks and protection still apply.
- **Undo:** each reply that changed the document gets an "Undo Claude's changes" button that puts the document back to how it was before that reply.
- **Tools:** Claude reads the document as numbered paragraphs, with bold, italic, underlined and other formatted words marked so it keeps them when it rewrites text, and can find and replace (keeping formatting), fill `{{placeholders}}`, edit, insert and delete paragraphs, copy paragraphs and tables with their formatting (for example to translate a document onto a new page), make tables and add or delete rows, format text, change lists (bullets, numbers, letters, levels), make and update charts, shapes and text boxes (and resize, move, describe or delete pictures and other objects), add links and bookmarks, and add, reply to and resolve comments. These are the same operations docxkeun-mcp offers, from `packages/mcp/src/operations.ts`.
- **Safety:** Claude acts as an *editor*: locked regions and the document's own protection (for example "comments only") are enforced, and refusals are explained in the chat.
- **Save:** "Download .docx" saves the edited document; "Export PDF" uses docxkeun's PDF writer.
- **Models:** Claude Opus 5.5 (default), Claude Sonnet 5.5 or Claude Haiku 5.5. Opus and Sonnet requests opt into server-side refusal fallbacks (`fallbacks: "default"`), so a request a safety classifier declines is retried on Anthropic's recommended fallback model instead of failing.

## Your API key and the relay

- The key you type is sent once to the relay on this computer, which checks it with the API and keeps it only in memory; it is forgotten when the server stops. "API key → Forget the key" clears it sooner. The browser stores nothing.
- The relay (`server/relay.mjs`) forwards only the calls the chat makes (Messages and Models) to `api.anthropic.com`, adds the key, and passes on only the request headers those calls need, never your browser's cookies or origin.
- It listens on your computer only and answers only its own page: requests from other websites, other host names, or without the page's own header are refused, so another site cannot spend your key through it.
- The built page has a Content-Security-Policy that allows network requests only to its own server, and loads no third-party scripts.

## Development

- `?mock` on the dev server (http://localhost:5260/?mock) replaces Claude with a scripted stand-in, to try the interface without a key or cost. It is never part of a build.
- Tests: `pnpm test` runs the agent loop against a scripted stream, the tools against real documents, and the relay as a real local server in front of a fake API.
- A live check against the real API (a few cents): `DOCXKEUN_LIVE=1 ANTHROPIC_API_KEY=sk-ant-... pnpm vitest run packages/chat/test/live.test.ts`.

The conversation loop is in `src/agent.ts` (streaming Messages API with tool use, input validation, refusal and fallback handling), the tools in `src/tools.ts`, the connection to the relay in `src/auth.ts`, the relay in `server/relay.mjs`, and the page in `src/main.ts`.
