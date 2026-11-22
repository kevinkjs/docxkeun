# Viewer chrome: toolbar, navigation panel, comments, right-click menu

Design for four parts of the docxkeun viewer (`docxkeun/viewer`), agreed in a design
conversation with mockups on 2026-11-21. Mockups: `.superpowers/brainstorm/49040-1795273159/content/`
(`toolbar-nav-v2.html`, `comments-v2.html`, `context-menu.html`), option **A** in each.

## Why

The viewer is used by the docxkeun playground (the public demo), the docxkeun chat app, the
docxkeun-mcp live preview, and by anyone who embeds docxkeun. Today:

- the formatting toolbar can only appear or disappear with editing, and looks plain;
- the navigation sidebar exists only in the playground (`playground/navigation.ts`), so the other
  apps don't have one;
- comment cards stack badly: every card shows its full text and all replies, one long thread pushes
  every later card far from its text, each page has its own 260px column so cards near a page's end
  get cut off, and narrow panes scroll sideways;
- right-click gives the browser's menu, which knows nothing about the document.

The owner wants a redesigned toolbar and navigation panel that an app can show and hide from code,
nicer comment stacking, and a right-click menu that apps can customise.

## Goals

1. **Toolbar** (redesign of the existing `FormatToolbar`): one grouped row that folds to a thin strip,
   and can be shown, hidden and folded from code.
2. **Navigation panel** (moved into the library): tabs for the outline, bookmarks and links, a search
   box, back/forward; folds to an icon rail; shown, hidden and folded from code.
3. **Comments** (rework of the comment column): compact cards in one document-wide column, an open
   thread that sits exactly beside its text, resolved threads tucked away, margin bubbles when there
   is no room for a column.
4. **Right-click menu** (new): a Word-like menu whose items depend on what was clicked, extensible by
   apps, off unless the app turns it on.
5. The playground, chat app and live preview adopt all four.

## Non-goals

- A multi-row or tabbed ribbon; a floating or selection-only toolbar (directions B and C were
  considered and set aside; a selection bubble could be added later).
- Editing an existing comment's text, @mentions, emoji reactions.
- Touch-specific gestures (an app can open the menu from code on long-press).
- Table structure commands (insert row/column, merge) in the menu: the editor has no API for them yet.

## Anti-goals (failure even if it works)

- Breaking current users: `toolbar: true`, `toolbar: <element>`, `showComments`, `editable`, the
  `commentclick` event and `viewer.root` as the scroll container must keep working.
- Taking the browser's right-click away from apps that didn't ask for the new menu.
- UI that covers document text (cards, bubbles, popovers and menus must not hide the text they are
  about).
- Making the viewer heavier for apps that use none of this: each part is built only when enabled.

## Constraints

- Zero runtime dependencies (the library rule). Plain DOM and injected CSS, like the existing
  toolbar and viewer; no framework.
- Works inside the chat app's Content-Security-Policy (no inline scripts; styles are injected
  `<style>` elements as today).
- Clipboard: browsers allow a page's own menu to write the clipboard after a user gesture
  (`navigator.clipboard.write`), but **reading** it (Paste) needs a permission prompt in Chromium and
  Safari and may be unavailable elsewhere. Paste must degrade (see the menu section).
- Accessibility: keyboard operable (toolbar roving focus, menu arrow keys, panel tabs), ARIA roles,
  visible focus, `prefers-reduced-motion` respected for any motion (card sliding, rail folding).
- Permissions: actions go through the existing `EditGuard` (`AccessPolicy.can()` gives a yes/no
  without throwing; `assert()` gives the reason as the error message).

## Approaches considered (for how it is built)

1. **Built into the viewer library** — chosen. Each part is its own module in `src/render/`, turned on
   by a viewer option and controlled through `viewer.toolbar`, `viewer.navigation`, `viewer.comments`,
   `viewer.contextMenu`. All apps get the same parts with no setup; the cost is a somewhat larger
   viewer bundle (mitigated by building parts only when enabled).
2. A separate `docxkeun-ui` package attaching to a viewer: leaner core, but two packages to version and
   the existing toolbar would move out (breaking `toolbar: true`).
3. Headless state and actions with a default renderer: most flexible for React/Vue apps, most work,
   and the default UI still has to be built.

## The design

### Shared shell

When the toolbar or the navigation panel is enabled, the viewer wraps itself in a frame inside the
host container (decision: the builder may adjust class names):

```
container
└─ .docxkeun-frame            (flex column, fills the container)
   ├─ .docxkeun-toolbar       (optional; or inside a host element given as `toolbar: el`)
   └─ .docxkeun-body          (flex row, min-height 0)
      ├─ .docxkeun-nav        (optional; panel or rail)
      └─ .docxkeun-viewer     (= viewer.root, the scroll container, unchanged role)
```

Without either part the DOM stays exactly as today. All four parts share one set of CSS variables
(`--docxkeun-ui-bg`, `-fg`, `-muted`, `-line`, `-hover`, `-accent`, `-accent-soft`, `-radius`,
`-shadow`, `-font`) defined on `.docxkeun-frame`/`.docxkeun-viewer`, light by default; apps theme
(including dark) by setting them. The existing `--docxkeun-toolbar-*` variables keep working as
aliases.

A new viewer event reports changes so apps can remember them:
`chromechange: { part: 'toolbar' | 'navigation' | 'comments' | 'contextMenu'; visible: boolean; collapsed?: boolean }`.
The library does not persist anything itself; the playground and apps store their choice.

### 1. Toolbar (direction A)

- One row, grouped with separators: undo · redo | paragraph style · font size | bold · italic ·
  underline · strikethrough · text colour · highlight | align left/centre/right/justify | bullets ·
  numbering · outdent · indent | left-to-right · right-to-left | Track Changes. Then a flexible space
  and a **⌃ fold button** at the right end.
- **Folded:** a thin full-width strip (about 18px) reading "⌄ Show toolbar"; clicking it unfolds.
- **Narrow widths:** groups that don't fit move into a "⋯" overflow dropdown at the end (rightmost
  groups first), so the row never wraps.
- Current behaviour kept: buttons mirror the caret's formatting (`caretchange`), keep focus in the
  page, and are disabled when there is no caret.
- **Options:** `toolbar: true | HTMLElement | { container?: HTMLElement; collapsed?: boolean; styles?: … }`
  (existing `ToolbarOptions.styles` kept).
- **API** (`viewer.toolbar`, present when enabled): `show()`, `hide()` (removed from view entirely;
  e.g. a read-only mode), `visible`, `collapsed` (get/set; same as the fold button), `element`.
  `setEditable(false)` hides it as today; `setEditable(true)` restores its previous visible state.

### 2. Navigation panel (direction A)

- Moves from `playground/navigation.ts` into the library (`src/render/navigation-panel.ts`), built on
  the existing `collectNavTargets`, `nextTarget`, `viewer.goTo`, `back()`, `forward()`,
  `canGoBack/canGoForward` and the `navigate` event.
- **Panel (about 240px):** back/forward buttons at the top; a segmented tab control **Outline ·
  Bookmarks · Links**; a search box filtering the current tab; the list (outline indented by heading
  level, current heading highlighted as you scroll); per-tab previous/next buttons (as the playground
  has); "Show hidden bookmarks" switch on the Bookmarks tab (as today). A **« fold** control at the
  bottom.
- **Rail (36px):** icons for Outline, Bookmarks, Links (and a Comments icon when comments exist, which
  scrolls to the next comment); clicking an icon unfolds the panel on that tab. Tooltips name them.
- **Narrow viewers** (frame narrower than about 720px — builder may tune): the panel starts folded to
  the rail.
- **Options:** `navigation: true | { tab?: 'outline' | 'bookmarks' | 'links'; collapsed?: boolean }`.
- **API** (`viewer.navigation`): `show(tab?)`, `hide()`, `visible`, `collapsed`, `tab`, `refresh()`
  (it also refreshes itself after each layout).
- **Compatibility** (my call, raised by the builder check): `viewer.navigation` is already public in
  0.1.3 and returns `{ headings, bookmarks, links }` (documented in the README, used by the
  playground). It stays one object that keeps those three properties and gains the panel controls
  above, so nothing breaks. The object always exists; the panel's DOM is built on the first `show()`
  (or at start when the `navigation` option is set). `viewer.toolbar` and `viewer.comments` follow the
  same pattern for consistency: always present, built on first use, so an app can show a part from
  code even if it wasn't enabled at construction.
- The playground drops its own `navigation.ts` and its "Navigation" button now drives
  `viewer.navigation`.

### 3. Comments (direction A, with B as the narrow fallback)

- **One column for the whole document** instead of one per page: a single positioned layer in the
  scroll container, to the right of the pages (the rows still reserve its width so pages stay
  centred as today). Cards are placed at their anchor's absolute y, so a card can sit across a page
  gap and is never cut off.
- **Resting card:** initials avatar (colour per author, reusing the revision author colours), name,
  relative time, and the first **two lines** of the comment (clamped). Replies fold into "N replies".
- **Opening a thread** (click the card or its highlighted text, or `viewer.comments.open(id)`): the
  card expands (full text, replies), moves so its top lines up exactly with its anchor, draws a thin
  connector line to the highlight, and the other cards move up (those above) and down (those below)
  so none overlap, with a short slide (none under reduced motion). The highlight turns darker (as
  today's `.active`). Esc or clicking elsewhere closes it.
- **Actions on an open card**, shown only when the viewer is editable and the guard allows
  commenting: Reply (inline text box, Post/Cancel), Resolve / Reopen, Delete (only own comments for
  commenters, as `AccessPolicy` rules). They call the existing `DocxEditor` comment methods.
- **New comment:** `viewer.comments.compose()` (used by the right-click menu and toolbar apps) puts a
  draft card with a text box beside the current selection; Post calls `addComment` with the author
  from editing options / guard user.
- **Resolved threads** are hidden; a small "Show resolved (N)" switch at the top of the column (only
  when there are resolved ones) shows them faded, as today.
- **Bubble mode:** when the viewer is too narrow for the column, the column becomes a 36px strip of round bubbles
  beside the lines that have comments ("2" when several anchors share a spot). Clicking a bubble opens
  the same card as a popover next to it, with ‹ › to step through the threads at that spot.
- **When `auto` switches** (my call, raised by the builder check): the rule must not depend on the
  current zoom, because the chat app and live preview fit the page to the window (with the old rule
  each mode keeps itself in place and the result depends on history). Rule: **cards** when the viewer
  is at least as wide as 75% of the widest page at 100% zoom plus the card column (about 300px),
  otherwise **bubbles**, with about 40px of hysteresis so a resize near the edge doesn't flip it back
  and forth. Apps that fit to width compute the fit after reading `viewer.comments.width`, which
  reflects the mode the rule picked. The builder may tune the 75% and 300px.
- **Options:** keep `showComments` (false hides everything); add
  `comments?: { mode?: 'auto' | 'cards' | 'bubbles'; showResolved?: boolean }` (default `auto`).
- **API** (`viewer.comments`): `show()`, `hide()`, `visible`, `mode`, `open(id)`, `close()`,
  `showResolved`, `compose()`, and `width` (the space the column takes now, for apps that fit the page
  to the window — the live preview and chat app currently hard-code 272px).
- `commentclick` keeps firing on open.

### 4. Right-click menu (direction A)

- **Off by default.** Enabled with `contextMenu: true | { items?: (at, defaults) => MenuItem[] }`.
  Shift + right-click always shows the browser's own menu. Right-clicks outside the pages are not
  touched.
- **Look:** a classic grouped list (about 220px), icon + label + shortcut on the right, separators,
  submenus opening to the side (flipped at screen edges), the menu itself flipped to stay on screen.
  Keyboard: arrows, Enter, Right/Left for submenus, Esc closes; focus returns to the page.
- **Default items** (`defaults`), by what was clicked:
  1. Always: Cut ⌘X, Copy ⌘C, Paste ⌘V, Paste as plain text ⇧⌘V (cut/paste only when editable).
  2. Text: Comment (opens the comment draft), Link… (a small popover with a URL / bookmark field),
     Format › (bold, italic, underline, strikethrough, superscript, subscript, clear formatting),
     Paragraph › (alignment, bullets, numbering, indent, direction, paragraph styles).
  3. On a **link:** a header line with its address, Open link, Copy link address, Edit link…, Remove
     link.
  4. On a **comment highlight:** a header "Comment by …", Reply, Resolve, Delete comment.
  5. On a **picture:** Alt text… (popover), Copy picture, Delete.
  6. In a **header or footer:** Edit header/footer (enters it, as double-click does), Remove
     header/footer (for that page type and section).
  7. The app's items last.
- **Disabled items:** shown greyed out when the guard refuses them (read-only, protection, locks),
  with the refusal reason as the tooltip; items that need editing are disabled in a view-only viewer.
- **Paste:** uses the async clipboard (HTML first, via the existing rich-paste path, else text). If the
  browser refuses or lacks it, the item shows "Paste (⌘V)" and choosing it just focuses the page so
  the keyboard shortcut works.
- **`at` (what was clicked):** `{ x, y, page, selection: SelectionInfo | undefined, target:
  { kind: 'text' | 'link' | 'comment' | 'image' | 'header' | 'footer'; link?: LinkTarget;
  commentId?: string; image?: ImageInline; part?: string } }`. A right-click outside the current
  selection first moves the caret there (as Word does).
- **`MenuItem`:** `{ id?, label, icon?, shortcut?, disabled?: boolean | string (reason), run(at) } |
  { label, items: MenuItem[] } (submenu) | '-' (separator)`.
- **API** (`viewer.contextMenu`): `enable()`, `disable()`, `open(x, y)` (e.g. from a "⋯" button or a
  long-press), `close()`.

### Undo (decided by the owner)

Everything the new parts do is undoable with ⌘Z / the toolbar's Undo and redoable, as in Word: posting,
replying to, resolving, reopening and deleting comments; adding, editing and removing links; alt text;
deleting pictures; removing headers and footers; and the toolbar's paragraph and list actions.

- One action is one undo step, including asynchronous ones (the step is recorded when the action
  finishes; a failed action records nothing and leaves the document as it was).
- An undo step restores **every part the action touched**, not only `document.xml`: the comment parts,
  relationships, header/footer parts, settings, numbering and styles, and the set of parts itself
  (a part the action created is removed again; one it deleted comes back). Redo is symmetrical.
- The viewer exposes this so apps' own buttons get the same behaviour, e.g.
  `await viewer.editing.transaction(async (editor) => { … })` (name is the builder's call). The
  playground's Comment / Link / Lock buttons move onto it.
- This also fixes an existing bug: today a ⌘Z after a comment or link was added restores an older
  `document.xml` and leaves comment markers or links half-removed.
- How the snapshot is taken (whole parts, or only those an action marks dirty) is the builder's call,
  as long as memory stays reasonable for long documents (the undo history is already capped at 200).

### App adoption

- **Playground:** toolbar on (foldable), navigation panel from the library (its own `navigation.ts`
  removed, header button kept), comments `auto`, context menu on with its existing actions (Comment,
  Link, Lock…) moved into it instead of `prompt()` where the new popovers cover them.
- **Chat app:** comments `auto`; context menu on with an extra **"Ask Claude about this"** item that
  puts the selection into the chat box; navigation available but folded to the rail; toolbar off as
  today (the chat drives edits) — builder may offer a toggle.
- **Live preview (read-only):** navigation (folded to the rail by default), comments `auto`, context
  menu with Copy / Copy link address only (it is read-only); fit-width uses `viewer.comments.width`.

## Left to the builder

- Exact sizes, spacing, icons (inline SVG, as the toolbar uses) and colours within the shared
  variables; the precise narrow-width breakpoints.
- Animation details for card sliding and folding (short; none under reduced motion).
- How the comment layout handles extreme cases (dozens of comments on one line): stacking order and
  scroll behaviour, as long as nothing overlaps and the open card stays beside its anchor.
- Internal structure: module boundaries, how the viewer constructs parts lazily, tests. The repo's
  tests run in Node without a DOM; adding a DOM test environment as a dev-only dependency is
  acceptable, and the apps must be checked in a browser.
- Whether the React wrapper needs anything beyond passing the new options through (it spreads
  `ViewerOptions` today).
