# Contributing

## Development Setup

```bash
pnpm install
pnpm run fixtures:generate
pnpm run verify
```

Use Node.js 20 or newer for local development. The current test/build toolchain does not support Node 18.

To refresh the committed sanitized regression deck from a local source presentation:

```bash
PPTX_COMPLEX_SOURCE="/absolute/path/to/source-deck.pptx" pnpm run fixtures:sanitize-complex
```

This command is intended for maintainers. CI uses the committed sanitized fixture and does not need access to the original source deck.

## Development Workflow

- keep changes focused and backward-compatible unless the change is intentionally breaking
- add or update tests when parser or renderer behavior changes
- document new limitations or support changes in `README.md`
- open pull requests from topic branches instead of pushing directly to `main`
- run `pnpm run verify` before opening a pull request
- run `pnpm run pack:check` if you change package metadata or published files
- run `pnpm run test:browser` when you change browser-facing render behavior, bundling, or fixture handling
- if you update the complex regression deck, rerun `pnpm run verify` and `pnpm run test:browser` before pushing

## Release Process

1. Update `CHANGELOG.md`.
2. Bump the version in `package.json`.
3. Run `pnpm run verify`.
4. Commit the release changes.
5. Create and push a Git tag like `v0.1.0`.
6. Publish from GitHub Actions after npm trusted publishing is configured.

## Repository Rules

- `main` is branch-protected with required PR review, conversation resolution, linear history, and the `verify` status check
- release tags matching `v*` are protected against non-fast-forward updates and deletion

## Scope

This package targets browser-like environments. Changes that require Node-only or server-only behavior should stay optional and must not break the browser-first API.

The renderer is still under active development. In particular, grouped objects, charts, table graphic frames, and complex SmartArt/diagram layouts are not fully implemented yet, so changes in those areas should be explicit about current behavior and known regressions.
