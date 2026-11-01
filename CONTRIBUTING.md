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

The repository holds two packages:

- `docxkeun` (the root): the library. It must keep **zero runtime dependencies**; a test fails if `dependencies` gains an entry. Development tools are fine as devDependencies.
- `docxkeun-fonts` (`packages/fonts`): the bundled open fonts. Tests and the playground use them.

## Pull requests

- Add or update tests for behaviour changes (`test/*.test.ts`).
- Keep `pnpm typecheck` and `pnpm test` passing; CI runs them on Node 20, 22 and 24.
- Note user-visible changes in `CHANGELOG.md`.

## Releasing (maintainers)

1. Update versions in `package.json` (and `packages/fonts/package.json` if the fonts changed) and the changelog.
2. Commit, then tag: `git tag v0.1.0 && git push origin v0.1.0`.
3. The release workflow tests, builds and publishes both packages to npm with provenance. Packages whose version already exists on npm are skipped.
