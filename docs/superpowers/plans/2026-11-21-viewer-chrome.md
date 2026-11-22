# Viewer Chrome Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** A foldable toolbar, a library navigation panel, a reworked comment column and an extensible right-click menu for `docxkeun/viewer`, with document-wide undo for everything they do, adopted by the playground, chat app and live preview.

**Architecture:** Each part is a DOM-free model module (pure functions, unit-tested in Node) plus a plain-DOM component in `src/render/`, created lazily behind an always-present controller on the viewer (`viewer.toolbar`, `.navigation`, `.comments`, `.contextMenu`). Undo moves to a DOM-free `UndoHistory` that can restore whole-package checkpoints (`DocxDocument.checkpoint()` / `restore()`), so menu and card actions run as transactions.

**Tech Stack:** TypeScript, plain DOM, injected CSS, Vitest (Node; `happy-dom` as a new dev-only test environment for DOM components).

**Spec:** `docs/superpowers/specs/2026-11-21-viewer-chrome-design.md` (read it first; mockups in `.superpowers/brainstorm/49040-1795273159/content/`, option A).

## Global Constraints

- `dependencies` in `package.json` stays empty (enforced by an existing test). Dev dependencies are allowed.
- Nothing existing breaks: `toolbar: true | HTMLElement`, `showComments`, `editable`, the `commentclick` event, `viewer.root` as the scroll container, `viewer.navigation.headings/bookmarks/links`.
- The right-click menu is **off by default**; Shift + right-click always shows the browser's menu; right-clicks outside the pages are untouched.
- Shared CSS variables on `.docxkeun-frame` and `.docxkeun-viewer`: `--docxkeun-ui-bg`, `--docxkeun-ui-fg`, `--docxkeun-ui-muted`, `--docxkeun-ui-line`, `--docxkeun-ui-hover`, `--docxkeun-ui-accent`, `--docxkeun-ui-accent-soft`, `--docxkeun-ui-radius`, `--docxkeun-ui-shadow`, `--docxkeun-ui-font`; existing `--docxkeun-toolbar-*` keep working as aliases. Light by default.
- Sizes from the spec: folded toolbar strip ≈ 18px; navigation panel ≈ 240px; rail 36px; navigation starts folded below a 720px frame width; comment cards clamp to 2 lines; bubble strip 36px; menu ≈ 220px wide.
- Comment mode rule: **cards** when `viewerWidth >= 0.75 × widestPageWidthAt100% + 300`, else **bubbles**, with 40px hysteresis.
- Motion (card sliding, folding) is short and **absent under `prefers-reduced-motion`**.
- The library persists nothing; it emits `chromechange: { part: 'toolbar' | 'navigation' | 'comments' | 'contextMenu'; visible: boolean; collapsed?: boolean }`.
- Undo history cap stays 200 entries.
- Keyboard and ARIA: toolbar `role="toolbar"` with roving focus; navigation tabs `role="tablist"`; menu `role="menu"`/`menuitem`, `aria-disabled`; visible focus rings.
- **Do not commit** unless the user asks: the "Checkpoint" steps run the checks instead.

## Review Focus

- **Empty documents** (no comments, headings, bookmarks or links): panels show a short empty state, the comment column takes no width, nothing throws. Tests: Task 5 (`renders empty states`), Task 7 (`no threads, no width`).
- **Many comments on one spot** (30 anchors at the same y): no two cards overlap, the open card's top equals its anchor, a bubble shows "30". Tests: Task 6 (`30 cards at one anchor`), Task 6 (`bubbleGroups merges near anchors`).
- **Read-only and protected documents:** menu items that edit are disabled with the guard's reason, card actions hidden, a refused transaction leaves the document unchanged. Tests: Task 9 (`disabled with the guard's reason`), Task 2 (`failed transaction restores and records nothing`).
- **Resize and zoom:** comment mode does not flicker at the threshold; toolbar overflow recomputes. Tests: Task 6 (`hysteresis`), Task 4 (`fitGroups`).
- **Mixed undo:** type → add a comment from the menu → type → undo ×3 → document identical to the start; redo ×3 → back; loading another document clears the history. Test: Task 2 (`mixed story and package undo`).

---

### Task 1: Package checkpoints

**Files:**
- Modify: `src/opc/package.ts` (add `checkpoint()` / `restore()`)
- Modify: `src/model/document.ts` (add `checkpoint()` / `restore()` wrapping the package plus document fields)
- Test: `test/checkpoint.test.ts`

**Interfaces:**
- Produces: `OpcPackage.checkpoint(): PackageCheckpoint`, `OpcPackage.restore(cp: PackageCheckpoint): void`; `DocxDocument.checkpoint(): DocumentCheckpoint`, `DocxDocument.restore(cp: DocumentCheckpoint): void` (both types opaque, exported as types from `src/model/document.ts`).

- [ ] **Step 1: Write the failing tests** in `test/checkpoint.test.ts` using `buildDocx`:
  - `restores comments, links and headers added after the checkpoint`: open a doc, `cp = doc.checkpoint()`, then `editor.addComment(...)`, `editor.addHyperlink(..., { url: 'https://example.com' })`, `editor.setHeader('H')`; `doc.restore(cp)`; expect `readComments(doc)` empty, `editor.hyperlinks()` empty, `doc.storyParts()` empty, `doc.pkg.partNames()` equal to the list before, and `blocksToText(doc.body().blocks)` equal to before.
  - `restores a deleted part`: `setHeader('H')`, `cp`, `removeHeader()`, `restore(cp)` → `headersAndFooters()` has the header with text `'H'` and its relationship resolves.
  - `checkpoints are independent of later edits`: edit after `cp`, restore, edit again, restore the same `cp` again → same state both times.
  - `save after restore`: `DocxDocument.open(await doc.save())` reads the restored state.
- [ ] **Step 2: Run** `npx vitest run test/checkpoint.test.ts` — expect FAIL (`checkpoint is not a function`).
- [ ] **Step 3: Implement.** `OpcPackage.checkpoint()` copies `parts` (each `Part` shallow-copied, its parsed `xml` deep-cloned with `cloneElement` on the root, `bytes` kept by reference since bytes are never mutated in place), `order`, `defaults`, `overrides`, `contentTypesDirty`, and `relsCache` (arrays copied, `Relationship` objects copied). `restore()` puts those back. `DocxDocument.checkpoint()` adds `parts` (the `WordParts` record), `relMaps`, `mediaCache`; `restore()` puts them back, re-points `documentXml` at the restored main-part XML, calls `reparseDefinitions()` (which also invalidates story caches) and notifies change listeners for the main part (one `touch`-style notification, so viewers relayout).
- [ ] **Step 4: Run** `npx vitest run test/checkpoint.test.ts` — expect PASS.
- [ ] **Step 5: Checkpoint:** `pnpm typecheck && npx vitest run` — all pass.

### Task 2: Undo history with transactions

**Files:**
- Create: `src/render/undo.ts` (DOM-free)
- Modify: `src/render/editing.ts` (use `UndoHistory`; add `transaction`; route `formatParagraphs`, `setStyle`, `setList`, `indent` through transactions; clear history on document load)
- Test: `test/undo.test.ts`

**Interfaces:**
- Consumes: `DocxDocument.checkpoint()/restore()` (Task 1).
- Produces:
  - `type UndoEntry = { kind: 'story'; part: string; story: EditStory; root: XmlElement; container?: number; anchor: Caret; focus: Caret } | { kind: 'package'; checkpoint: DocumentCheckpoint; story: EditStory; anchor: Caret; focus: Caret; label: string }`
  - `class UndoHistory { constructor(limit = 200); push(e: UndoEntry): void; undo(current: () => UndoEntry): UndoEntry | undefined; redo(current: () => UndoEntry): UndoEntry | undefined; clear(): void; readonly canUndo: boolean; readonly canRedo: boolean }` — `undo`/`redo` take a function producing the entry for the current state (pushed on the opposite stack), return the entry to restore.
  - `InlineEditing.transaction<T>(label: string, fn: (editor: DocxEditor) => T | Promise<T>): Promise<T>` — checkpoint, run, push one `package` entry on success; on throw restore the checkpoint, push nothing, emit `editerror`, rethrow.
  - `InlineEditing.canUndo` / `canRedo` getters (used by the toolbar).
- [ ] **Step 1: Write the failing tests** in `test/undo.test.ts` (DocxDocument + `UndoHistory`, no DOM; a small helper applies an entry: story entries replace that part's root, package entries call `doc.restore`):
  - `package entry undo and redo`: checkpoint, `addComment`, push package entry; `undo` → no comments; `redo` → the comment is back with the same text.
  - `failed transaction restores and records nothing`: with an `AccessPolicy` viewer guard, a transaction whose `fn` calls `addComment` throws `PermissionError`; document unchanged; `canUndo` false.
  - `mixed story and package undo`: story entry (typing in body) → package entry (comment) → story entry (typing) → undo ×3 → `doc.save()` bytes re-opened give the original text and no comments; redo ×3 → both typings and the comment.
  - `cap at 200`: push 205 entries → 200 undoable.
- [ ] **Step 2: Run** `npx vitest run test/undo.test.ts` — expect FAIL.
- [ ] **Step 3: Implement** `UndoHistory` in `src/render/undo.ts`; in `editing.ts` replace `undoStack`/`redoStack` with it (keep the 1-second typing coalescing in `apply()`), add `transaction()`, `canUndo`, `canRedo`, and make `undo()`/`redo()` restore either kind (package entries: `doc.restore(cp)`, then `refreshModel()`, restore story and selection). Clear the history in `onLayout` when `viewer.document` changed identity.
- [ ] **Step 4: Run** `npx vitest run test/undo.test.ts test/editing-stories.test.ts` — expect PASS.
- [ ] **Step 5: Browser check** (playground, `preview_start` name `playground`): select a list paragraph → toolbar numbered list → ⌘Z → the list is gone and numbering.xml-dependent labels are back; ⌘⇧Z → it returns.
- [ ] **Step 6: Checkpoint:** `pnpm typecheck && npx vitest run`.

### Task 3: Frame, shared theme and DOM test environment

**Files:**
- Modify: `package.json` (dev dependency `happy-dom`, pinned to a release at least two weeks old)
- Create: `src/render/ui/theme.ts`, `src/render/ui/icons.ts`
- Modify: `src/render/viewer.ts` (frame on demand; `chromechange` event type; controllers' home)
- Test: `test/ui-frame.test.ts` (first line `// @vitest-environment happy-dom`)

**Interfaces:**
- Produces:
  - `ensureUiStyle(): void` (injects the shared variables and base classes once, id `docxkeun-ui-style`).
  - `icon(name: IconName, label?: string): string` — inline SVG markup; `type IconName = 'undo' | 'redo' | 'alignLeft' | 'alignCenter' | 'alignRight' | 'justify' | 'bullets' | 'numbers' | 'outdent' | 'indent' | 'ltr' | 'rtl' | 'track' | 'fold' | 'unfold' | 'outline' | 'bookmark' | 'link' | 'comment' | 'search' | 'back' | 'forward' | 'cut' | 'copy' | 'paste' | 'more'` (move the existing toolbar icons here).
  - `DocxViewer.frame(): { frame: HTMLDivElement; toolbarSlot: HTMLDivElement; body: HTMLDivElement }` (internal, `@internal`): on first call wraps `root` as `container > .docxkeun-frame > [.docxkeun-toolbar-slot, .docxkeun-body > root]`; never called when no part needs it.
  - `ViewerEvents.chromechange` as in Global Constraints.
- [ ] **Step 1: Add** `happy-dom` as a dev dependency (`pnpm add -D happy-dom@<pinned>`); confirm the dependencies test still passes.
- [ ] **Step 2: Write the failing tests** in `test/ui-frame.test.ts`: `new DocxViewer(container)` leaves `container.firstElementChild === viewer.root`; after `viewer.frame()` the root's parent has class `docxkeun-body` inside `.docxkeun-frame`, calling it twice returns the same elements, and `viewer.root` is still the element with `overflow:auto` class `docxkeun-viewer`.
- [ ] **Step 3: Run** `npx vitest run test/ui-frame.test.ts` — FAIL.
- [ ] **Step 4: Implement** `theme.ts`, `icons.ts`, `frame()`, the event type.
- [ ] **Step 5: Run** the test — PASS. **Checkpoint:** `pnpm typecheck && npx vitest run`.

### Task 4: Toolbar redesign and controller

**Files:**
- Create: `src/render/ui/overflow.ts`
- Modify: `src/render/toolbar.ts` (rewrite as `ToolbarController` + view), `src/render/viewer.ts` (`viewer.toolbar` getter; options), `src/render/index.ts` (exports)
- Test: `test/toolbar-model.test.ts` (Node), `test/toolbar-dom.test.ts` (happy-dom)

**Interfaces:**
- Consumes: `ensureUiStyle`, `icon`, `viewer.frame()` (Task 3); `InlineEditing.canUndo/canRedo`, `transaction` (Task 2); existing `caretchange`.
- Produces:
  - `fitGroups(groupWidths: number[], available: number, moreButtonWidth: number): number` — how many leading groups stay in the row; the rest go to the "⋯" menu (all fit → no more button needed).
  - `interface ToolbarOptions { container?: HTMLElement; collapsed?: boolean; styles?: Array<{ id: string; name: string }> }`; `ViewerOptions.toolbar?: boolean | HTMLElement | ToolbarOptions`.
  - `class ToolbarController { show(): void; hide(): void; readonly visible: boolean; collapsed: boolean; readonly element: HTMLElement | undefined }` at `viewer.toolbar` (always present; DOM built on first `show()`, or at construction when the option is set and `editable` is on).
- [ ] **Step 1: Write failing model tests:** `fitGroups([40,60,80,50], 300, 28) === 4`; `fitGroups([40,60,80,50], 200, 28) === 2` (40+60+28 ≤ 200, +80 would not fit); `fitGroups([], 100, 28) === 0`; `fitGroups([500], 100, 28) === 0`.
- [ ] **Step 2: Write failing DOM tests** (stub viewer object with `root`, `on`, `editing` returning a stub with `caretFormat`, `canUndo`): toolbar has `role="toolbar"`; groups in the spec's order with separators; the fold button sets `collapsed = true` and shows a strip labelled "Show toolbar"; clicking the strip unfolds; `hide()` sets `element.hidden` and emits `chromechange { part: 'toolbar', visible: false }`; arrow keys move focus between buttons (roving `tabindex`).
- [ ] **Step 3: Run** both — FAIL.
- [ ] **Step 4: Implement.** Groups (spec order): undo·redo | style·size | B·I·U·S·colour·highlight | 4 alignments | bullets·numbering·outdent·indent | ltr·rtl | Track; spacer; fold button. A `ResizeObserver` measures group widths and calls `fitGroups`; overflowed groups render inside a "⋯" dropdown (reuse the menu component from Task 8 once it exists — until then a simple popover list; Task 8 Step 6 switches it). Paragraph/list buttons call `editing.formatParagraphs` / `setList` / `indent` (now transactions). `setEditable(false)` hides; `setEditable(true)` restores the previous `visible`. Keep `toolbar: true` and `toolbar: element` behaviour (element = `container`).
- [ ] **Step 5: Run** both tests — PASS.
- [ ] **Step 6: Browser check** (playground): fold/unfold; narrow the window → groups move into "⋯"; buttons still mirror bold at the caret.
- [ ] **Step 7: Checkpoint:** `pnpm typecheck && npx vitest run`.

### Task 5: Navigation panel in the library

**Files:**
- Create: `src/render/ui/nav-model.ts`, `src/render/navigation-panel.ts`
- Modify: `src/render/viewer.ts` (`viewer.navigation` becomes a `NavigationController` that still exposes `headings`, `bookmarks`, `links`), `src/render/index.ts`
- Test: `test/nav-model.test.ts` (Node), `test/navigation-panel.test.ts` (happy-dom)

**Interfaces:**
- Consumes: `collectNavTargets`, `nextTarget`, `NavTarget`, `DocLocation` (existing `src/render/navigation.ts`); `viewer.goTo`, `back`, `forward`, `canGoBack`, `canGoForward`, `location`, `navigate` and `layout` events; `frame()` (Task 3).
- Produces:
  - `filterTargets(list: NavTarget[], query: string): NavTarget[]` (case- and accent-insensitive substring on `label`, and on the URL/anchor for links).
  - `currentHeading(headings: NavTarget[], at: DocLocation): NavTarget | undefined` (last heading at or above `at`).
  - `type NavTab = 'outline' | 'bookmarks' | 'links'`; `ViewerOptions.navigation?: boolean | { tab?: NavTab; collapsed?: boolean }`.
  - `class NavigationController { readonly headings: NavTarget[]; readonly bookmarks: NavTarget[]; readonly links: NavTarget[]; show(tab?: NavTab): void; hide(): void; readonly visible: boolean; collapsed: boolean; tab: NavTab; refresh(): void }` at `viewer.navigation`.
- [ ] **Step 1: Write failing model tests:** `filterTargets` matches "pay" to "2. Payment terms", "é" queries match "e", links match on URL; `currentHeading` picks the right heading for locations before the first heading (undefined), between headings, and on a later page.
- [ ] **Step 2: Write failing DOM tests** with a stub host (targets, `goTo` spy, `location`): tabs `role="tablist"` with three tabs; typing in search filters the list; clicking an item calls `goTo(target, { history: true })`; `collapsed = true` renders a 36px rail with four buttons (outline, bookmarks, links, comments) and clicking "bookmarks" unfolds on that tab; `renders empty states` ("No headings", "No bookmarks", "No links"); frame width < 720 at first show starts collapsed.
- [ ] **Step 3: Run** both — FAIL.
- [ ] **Step 4: Implement** the panel (back/forward, tabs, search, list with outline indentation by `level`, per-tab previous/next, "Show hidden bookmarks", fold control) and the controller; refresh after each `layout`; highlight `currentHeading` on scroll (throttled). The compat getters read `collectNavTargets(viewer.layout)` as today. The rail's Comments button appears only when `layout.commentAnchors.size > 0`; clicking it calls `viewer.scrollToComment(id)` for the first anchor below `viewer.location` (wrapping to the first), which Task 7 makes open the thread too.
- [ ] **Step 5: Run** both — PASS; also `npx vitest run test/navigation.test.ts` (existing) — PASS.
- [ ] **Step 6: Checkpoint:** `pnpm typecheck && npx vitest run`.

### Task 6: Comment layout model

**Files:**
- Create: `src/render/ui/comment-layout.ts`
- Test: `test/comment-layout.test.ts`

**Interfaces:**
- Produces:
  - `interface CardBox { id: string; anchorY: number; height: number }`
  - `layoutCards(cards: CardBox[], openId?: string, gap = 8): Map<string, number>` — tops in the same coordinates as `anchorY`. Without `openId`: in anchor order, `top = max(anchorY, previousBottom + gap)`. With `openId`: the open card's top is exactly its `anchorY`; cards before it are placed upward (`top = min(anchorY, nextTop − height − gap)`), cards after it downward; if that pushes the first card above 0, shift all down by the overflow.
  - `commentMode(viewerWidth: number, widestPageAt100: number, previous: 'cards' | 'bubbles' | undefined): 'cards' | 'bubbles'` — threshold `0.75 × widestPageAt100 + 300`; switching to cards needs `≥ threshold + 40` when previous was bubbles, switching to bubbles needs `< threshold − 40` when previous was cards; no previous → plain threshold.
  - `bubbleGroups(anchors: Array<{ id: string; y: number }>, minGap = 22): Array<{ y: number; ids: string[] }>` — anchors closer than `minGap` merge; group `y` is the first anchor's.
- [ ] **Step 1: Write failing tests** (gap 8):
  - resting: anchors a 50, b 60, c 200; heights 80, 60, 40 → tops a 50, b 138, c 206.
  - open with room: anchors 100, 200, 300; heights 40, 50, 40; open b → 100, 200, 300.
  - open with neighbours close: anchors 100, 110, 120; heights 40, 50, 40; open b → a 62, b 110, c 168.
  - open without room above: anchors a 50, b 60, c 200; heights 80, 60, 40; open b → a 0, b 88, c 200 (cards above and the open one shift down by the overflow; cards below follow the open card).
  - `30 cards at one anchor` (anchor 300, height 40 each), open the 15th → no overlaps, all tops ≥ 0, the open card's top is 300 + overflow only if the 14 above don't fit (they need 14 × 48 = 672 > 300, so assert top = 672).
  - `hysteresis`: with page 612 → threshold 759; `commentMode(760, 612, undefined) === 'cards'`, `commentMode(780, 612, 'bubbles') === 'bubbles'`, `commentMode(800, 612, 'bubbles') === 'cards'`, `commentMode(730, 612, 'cards') === 'cards'`, `commentMode(718, 612, 'cards') === 'bubbles'`; `bubbleGroups merges near anchors` (y 10, 20, 50 → [{10,[a,b]},{50,[c]}]).
- [ ] **Step 2: Run** — FAIL. **Step 3: Implement** (when the first card above would start below 0, shift the cards above and the open card down by that amount, then lay out the cards below from the open card's bottom). **Step 4: Run** — PASS.
- [ ] **Step 5: Checkpoint:** `pnpm typecheck && npx vitest run`.

### Task 7: Comment column component

**Files:**
- Create: `src/render/comments.ts` (`CommentsController` + view)
- Modify: `src/render/viewer.ts` (remove `renderComments` and per-page gutters from `mountPage`; rows reserve `viewer.comments.width`; `viewer.comments` getter; `comments` option; `scrollToComment` and `commentclick` routed through the controller), `src/render/index.ts`
- Modify: `src/render/editing.ts` only if needed to expose author for new comments (use existing editing options `author` / guard user)
- Test: `test/comments-dom.test.ts` (happy-dom)

**Interfaces:**
- Consumes: `layoutCards`, `commentMode`, `bubbleGroups` (Task 6); `InlineEditing.transaction` (Task 2); `readComments`/`CommentInfo`; `layout.commentAnchors`, `page.comments`; `DocxEditor.addComment/replyToComment/resolveComment/deleteComment`; `REVISION_COLORS` author colours (reuse the reader's palette: export it from `src/model/content.ts` if not exported).
- Produces:
  - `interface CommentThread { comment: CommentInfo; replies: CommentInfo[]; anchor: { page: number; y: number; rects: Rect[] } }`
  - `class CommentsController { show(): void; hide(): void; readonly visible: boolean; mode: 'auto' | 'cards' | 'bubbles'; readonly effectiveMode: 'cards' | 'bubbles'; open(id: string): void; close(): void; showResolved: boolean; compose(): void; readonly width: number }` at `viewer.comments`.
  - `ViewerOptions.comments?: { mode?: 'auto' | 'cards' | 'bubbles'; showResolved?: boolean }` (`showComments: false` still hides all).
- [ ] **Step 1: Write failing DOM tests** (the controller takes a small host: threads, absolute anchor y per thread, `editable`, guard `can`, scroll container) — positions come from `layoutCards`, so tests assert DOM state, not pixels:
  - resting cards: author initials, 2-line clamp class, "2 replies" for a thread with two replies, resolved threads absent and a "Show resolved (1)" switch present.
  - `open('id')` → card has class `open` and `aria-expanded="true"`, replies rendered, connector element present; Esc closes.
  - actions: Reply/Resolve/Delete buttons only when `editable` and `can('comment')`; Reply posts through a host `transaction` spy.
  - `compose()` with a selection → a draft card with a textarea; Post calls the host's add-comment through `transaction`; Cancel removes it.
  - bubble mode (`mode = 'bubbles'`): one bubble per `bubbleGroups` group, text "2" for a merged group; clicking opens a popover card with ‹ › stepping between the two threads.
  - `no threads, no width`: `width === 0` and no column element.
- [ ] **Step 2: Run** — FAIL.
- [ ] **Step 3: Implement.** One absolutely positioned layer inside `viewer.root`, right of the pages (x = page row's right edge + 12px gap); anchor y = `pageElement(page).offsetTop + y × scale`; measure card heights after render, then place with `layoutCards`; re-place on `layout`, zoom and resize; `auto` uses `commentMode(root.clientWidth, widest page width × 96/72, previous)`; bubble strip 36px wide. Slide with `transition: top .18s` (none under reduced motion). Card colours: initials avatar per author. Clicking a highlight (`.docxkeun-comment-hl`, existing overlay) opens its thread. Emit `commentclick` on open and `chromechange` on show/hide/mode change.
- [ ] **Step 4: Run** — PASS; run `npx vitest run` (existing viewer tests unaffected).
- [ ] **Step 5: Browser check** (playground sample, which has a threaded comment): resting cards beside the highlight; open the thread → connector, neighbours move; narrow the pane below ≈ 760px → bubbles; widen past ≈ 800px → cards.
- [ ] **Step 6: Checkpoint:** `pnpm typecheck && npx vitest run`.

### Task 8: Menu and popover components

**Files:**
- Create: `src/render/ui/menu.ts`, `src/render/ui/popover.ts`
- Modify: `src/render/toolbar.ts` (its "⋯" overflow uses `openMenu`)
- Test: `test/menu-model.test.ts` (Node), `test/menu-dom.test.ts` (happy-dom)

**Interfaces:**
- Produces:
  - `type MenuItem = '-' | { id?: string; label: string; icon?: IconName | string; shortcut?: string; disabled?: boolean | string; run?: () => void | Promise<void>; items?: MenuItem[] }` (`disabled` string = reason shown as tooltip; `items` = submenu).
  - `positionMenu(at: { x: number; y: number }, size: { w: number; h: number }, viewport: { w: number; h: number }, margin = 4): { x: number; y: number }` — flips left/up when it would overflow, then clamps inside the margin.
  - `openMenu(items: MenuItem[], at: { x: number; y: number }, options?: { label?: string; onClose?: () => void }): MenuHandle` with `interface MenuHandle { close(): void; readonly element: HTMLElement }`.
  - `promptPopover(options: { title: string; value?: string; placeholder?: string; at: { x: number; y: number }; validate?: (v: string) => string | undefined }): Promise<string | undefined>` (resolves undefined on Esc/cancel).
- [ ] **Step 1: Write failing model tests** for `positionMenu`: fits → unchanged; near the right edge → `x = at.x − w`; near the bottom → `y = at.y − h`; larger than the viewport → clamped to the margin.
- [ ] **Step 2: Write failing DOM tests:** `role="menu"`, items `role="menuitem"`, separators `role="separator"`; ArrowDown/ArrowUp skip disabled items and separators; Enter runs and closes; ArrowRight opens a submenu, ArrowLeft closes it; Esc closes and calls `onClose`; a disabled item with a reason has `aria-disabled="true"` and `title` equal to the reason; clicking outside closes. `promptPopover` resolves the typed value on Enter, undefined on Esc, and shows `validate`'s message without resolving.
- [ ] **Step 3: Run** — FAIL. **Step 4: Implement** (shortcuts right-aligned; icons via `icon()` or a given string; focus returns to the previously focused element on close).
- [ ] **Step 5: Run** — PASS. **Step 6:** switch the toolbar's "⋯" to `openMenu`; re-run `test/toolbar-dom.test.ts` — PASS.
- [ ] **Step 7: Checkpoint:** `pnpm typecheck && npx vitest run`.

### Task 9: Right-click menu: targets, default items, controller, editor additions

**Files:**
- Create: `src/render/context-menu.ts` (`hitTarget`, `defaultItems`, `ContextMenuController`)
- Modify: `src/api/editor.ts` (add `setImageAlt`, `setHyperlinkTarget`), `src/render/viewer.ts` (`viewer.contextMenu`; `contextMenu` option), `src/render/editing.ts` (add `placeCaretAt(page: number, x: number, y: number): void`, using the existing `hitTest`; no-op when editing is off), `src/render/index.ts`
- Test: `test/context-menu.test.ts` (Node: `hitTarget`, `defaultItems`), `test/editor-links-alt.test.ts` (Node)

**Interfaces:**
- Consumes: `openMenu`, `promptPopover`, `MenuItem` (Task 8); `transaction` (Task 2); `CommentsController.open/compose` (Task 7); `SelectionInfo`, `describeSelection`; layout `Page` (`links`, `comments`, image items with `inline`, `storyLines`); `DocxEditor.removeHyperlink`, `addHyperlink`, `removeHeaderFooter`, `headersAndFooters`, `deleteRange`, comment methods; editing `toggle`, `formatParagraphs`, `setStyle`, `setList`, `indent`, `enterStory`.
- Produces:
  - `type MenuTargetKind = 'text' | 'link' | 'comment' | 'image' | 'header' | 'footer'`
  - `interface MenuTarget { kind: MenuTargetKind; link?: LinkTarget; commentId?: string; image?: ImageInline; part?: string }`
  - `interface MenuContext { x: number; y: number; page: number; selection: SelectionInfo | undefined; target: MenuTarget }`
  - `hitTarget(page: Page, x: number, y: number): MenuTarget` (priority: link → comment area → inline image item (not `frameOnly`) → header/footer story line → text).
  - `interface MenuEnv { editable: boolean; can(action: 'edit' | 'format' | 'comment' | 'link'): string | undefined /* refusal reason */; clipboardRead: boolean; act: MenuActions }` where `MenuActions` holds the callbacks the items run (copy, cut, paste, pastePlain, comment, link, format(f), paragraph(f), openLink, copyLink, editLink, removeLink, reply, resolve, deleteComment, altText, copyImage, deleteImage, editStory, removeStory).
  - `defaultItems(at: MenuContext, env: MenuEnv): MenuItem[]` — groups and labels exactly as the spec's list (Cut ⌘X, Copy ⌘C, Paste ⌘V, Paste as plain text ⇧⌘V; Comment, Link…, Format ›, Paragraph ›; per-target groups; Paste becomes "Paste (⌘V)" when `!env.clipboardRead`; Cut/Paste only when `editable`).
  - `ViewerOptions.contextMenu?: boolean | { items?: (at: MenuContext, defaults: MenuItem[]) => MenuItem[] }`; `class ContextMenuController { enable(): void; disable(): void; open(x: number, y: number): void; close(): void; readonly enabled: boolean }` at `viewer.contextMenu`.
  - `DocxEditor.setImageAlt(drawing: XmlElement, text: string): void` (sets `wp:docPr/@descr`, checks `'edit'`); `DocxEditor.setHyperlinkTarget(hyperlink: XmlElement, target: { url: string } | { bookmark: string }): Promise<void>` (same scheme rules as `addHyperlink`).
- [ ] **Step 1: Write failing Node tests:**
  - `hitTarget` on a laid-out fixture page: a point inside a link area → `link`; inside a comment area → `comment` with its id; on an inline picture → `image`; on a header story line → `header` with `part`; elsewhere → `text`.
  - `defaultItems` for text with editing on: labels in order `['Cut','Copy','Paste','Paste as plain text','-','Comment','Link…','Format','Paragraph']`, Format has submenu labels Bold…Clear formatting; for a link adds a header item showing the URL and `Open link`, `Copy link address`, `Edit link…`, `Remove link`; read-only (`editable: false`) → no Cut/Paste and Comment disabled; `disabled with the guard's reason`: `can('comment')` returns "The document only allows comments" → Comment item's `disabled` is that string; `clipboardRead: false` → label `Paste (⌘V)`.
  - `setImageAlt` writes `descr` and survives save; `setHyperlinkTarget` changes a URL link to another URL (new relationship) and to a bookmark (`w:anchor`, `r:id` removed); refuses `javascript:`.
- [ ] **Step 2: Run** — FAIL.
- [ ] **Step 3: Implement** `hitTarget`, `defaultItems`, the editor methods, and `ContextMenuController`: listens for `contextmenu` on `viewer.root` only when enabled; ignores events with `shiftKey` and outside `.docxkeun-page`; converts the point to page coordinates; if the point is outside the current selection, calls `editing.placeCaretAt(page, x, y)`; builds `at`, calls `options.items?.(at, defaults) ?? defaults`, opens `openMenu`. Actions: copy/cut via `navigator.clipboard.write` with `text/html` (existing `paragraphsToHtml`) and `text/plain`; paste via `navigator.clipboard.read`/`readText` into the existing rich-paste path; edits wrapped in `editing.transaction(label, …)`; Link… and Edit link… and Alt text… use `promptPopover`; Edit header/footer calls `editing.enterStory`; Remove header/footer finds its `type`/`section` via `headersAndFooters()` and calls `removeHeaderFooter` in a transaction.
- [ ] **Step 4: Run** — PASS.
- [ ] **Step 5: Browser check** (playground with `contextMenu: true`): right-click selected text → menu at the pointer; Shift+right-click → browser menu; on a link → link items; Comment → draft card; ⌘Z undoes the comment; read-only mode → editing items greyed with the reason on hover; keyboard navigation works.
- [ ] **Step 6: Checkpoint:** `pnpm typecheck && npx vitest run`.

### Task 10: Adopt in the playground, chat app and live preview

**Files:**
- Modify: `playground/main.ts`, `playground/index.html` (remove `#nav` pane and its CSS), delete `playground/navigation.ts`
- Modify: `packages/chat/src/main.ts`, `packages/chat/index.html`, `packages/chat/src/style.css` as needed
- Modify: `packages/mcp/src/preview-page.ts`
- Test: existing suites; browser checks below

**Interfaces:**
- Consumes: every controller and option from Tasks 4–9.
- [ ] **Step 1: Playground:** viewer options `{ editable: true, toolbar: true, navigation: true, comments: { mode: 'auto' }, contextMenu: { items: (at, d) => [...d, '-', { label: 'Lock…', run: … }, { label: 'Unlock', run: … }] } }`; the header "Navigation" button toggles `viewer.navigation.visible`; Comment/Link/Bold/Lock buttons call through `viewer.editing.transaction`; replace `prompt()` for comment text and link URL with `viewer.comments.compose()` and `promptPopover`; remember fold states from `chromechange` in `localStorage` (try/catch).
- [ ] **Step 2: Chat app:** `navigation: { collapsed: true }`, `comments: { mode: 'auto' }`, `contextMenu: { items: (at, d) => [...d, '-', { label: 'Ask Claude about this', icon: '✦', run: () => putSelectionInComposer(at.selection) }] }` where `putSelectionInComposer` reuses the existing selection chip and focuses the message box; fit-width uses `viewer.comments.width` instead of the hard-coded 272.
- [ ] **Step 3: Live preview:** `navigation: { collapsed: true }`, `comments: { mode: 'auto' }`, `contextMenu: { items: (at, d) => d.filter(i => i !== '-' && ['Copy', 'Copy link address', 'Open link'].includes(i.label)) }`; `applyFit` uses `viewer.comments.width`; script still parses (`node --check` on the extracted `PREVIEW_SCRIPT`, as done before).
- [ ] **Step 4: Run** `pnpm build && pnpm build:mcp && pnpm typecheck:mcp && pnpm typecheck:chat && pnpm build:chat && npx vitest run` — all pass. The React wrapper (`src/react/index.ts`) spreads `ViewerOptions` into the viewer, so the new options pass through; confirm it typechecks with them and needs no change.
- [ ] **Step 5: Browser checks:** playground (`preview_start` name `playground`): toolbar folds; navigation panel/rail works, search filters, back/forward; comments open/move/bubble at narrow width; right-click menu with Lock items; ⌘Z across a comment. Chat (`?mock`): rail present, "Ask Claude about this" fills the composer with the selection; no console errors. Live preview: rail, comments, read-only menu.
- [ ] **Step 6: Checkpoint:** all commands from Step 4 again.

### Task 11: Documentation

**Files:**
- Modify: `README.md` (new "Viewer UI" section: toolbar, navigation, comments, right-click menu, undo transactions, CSS variables, `chromechange`), `CHANGELOG.md` (Unreleased → docxkeun), `packages/chat/README.md`, `packages/mcp/README.md` (live preview: navigation rail, comments, read-only menu)
- [ ] **Step 1: Write** the README section with one code block covering the options and controllers named in Tasks 4, 5, 7, 9 and `viewer.editing.transaction`, and a sentence each on Shift+right-click, read-only behaviour and theming.
- [ ] **Step 2: Update** the changelog and the two package READMEs.
- [ ] **Step 3: Checkpoint:** `pnpm typecheck && npx vitest run && npm pack --dry-run` — pass.
