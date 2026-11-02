# Changelog

All notable changes to docxkeun and docxkeun-fonts. Versions follow [semantic versioning](https://semver.org); while the version is 0.x, minor versions may contain breaking changes.

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
