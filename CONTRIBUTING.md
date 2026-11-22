# Contributing

Thanks for helping. Bug reports with a small `.docx` that shows the problem are the most useful contribution: layout differences from Word are much easier to fix with the file and a screenshot (or PDF) of how Word shows it.

## Development

Requires Node 20.12+ and pnpm.

```bash
pnpm install
pnpm test        # vitest
pnpm typecheck
pnpm dev         # playground at http://localhost:5199
pnpm build       # dist/
```

The repository holds three published packages and one app (a pnpm workspace):

- `docxkeun` (the root): the library. It must keep **zero runtime dependencies**; a test fails if `dependencies` gains an entry. Development tools are fine as devDependencies.
- `docxkeun-fonts` (`packages/fonts`): the bundled open fonts. Tests and the playground use them.
- `docxkeun-mcp` (`packages/mcp`): an MCP server for AI clients, built on docxkeun. It may have dependencies (the MCP SDK and zod). It depends on docxkeun and docxkeun-fonts with normal version ranges, which pnpm links to the copies in this repository. `pnpm build && pnpm build:mcp && pnpm typecheck:mcp` builds and checks it; its tests are part of `pnpm test`.
- `docxkeun-chat` (`packages/chat`, private, not published): a web app for chatting with Claude about a document, built on docxkeun and the document operations in `packages/mcp/src/operations.ts`. `pnpm dev:chat` runs it with its API relay (`server/relay.mjs`); `?mock` uses a scripted stand-in instead of the API.

## Pull requests

- Add or update tests for behaviour changes (`test/*.test.ts`).
- Keep `pnpm typecheck` and `pnpm test` passing; CI runs them on Node 20, 22 and 24.
- Note user-visible changes in `CHANGELOG.md`.

## Releasing (maintainers)

1. Update versions in `package.json` (and in `packages/fonts/package.json` or `packages/mcp/package.json` if those changed) and the changelog. When docxkeun's minor version changes, update the `docxkeun` range in `packages/mcp/package.json` too (a test checks it).
2. Commit, then tag: `git tag v0.1.0 && git push origin v0.1.0`.
3. The release workflow tests, builds and publishes the packages to npm (with provenance once the repository is public). Packages whose version already exists on npm are skipped.
