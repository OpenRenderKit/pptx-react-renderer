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
- add exactly one release label to every pull request: `release:patch`, `release:minor`, `release:major`, or `release:skip`
- run `pnpm run verify` before opening a pull request
- run `pnpm run pack:check` if you change package metadata or published files
- run `pnpm run test:browser` when you change browser-facing render behavior, bundling, or fixture handling
- if you update the complex regression deck, rerun `pnpm run verify` and `pnpm run test:browser` before pushing

## Release Process

1. Open a pull request with exactly one release label.
2. Merge the pull request into `main`.
3. GitHub Actions verifies the repo, resolves the release version, publishes to npm, pushes a `vX.Y.Z` tag, and creates a GitHub release.
4. Use `release:skip` when a merge should not publish.

Before `1.0.0`, release labels are softened on purpose:

- `release:major` becomes a semver minor bump
- `release:minor` becomes a semver patch bump
- `release:patch` remains a patch bump

The first automated publish uses the version already in `package.json`. Later publishes use the most recent npm version as the base for the next bump.

Bootstrap note:

- either add an `NPM_TOKEN` repository secret so GitHub Actions can publish the first version
- or publish the first version manually, configure npm trusted publishing for `.github/workflows/publish.yml`, and then rely on merge-to-main releases afterward

## Repository Rules

- `main` is branch-protected with required PR review, conversation resolution, linear history, and the `verify` status check
- release tags matching `v*` are protected against non-fast-forward updates and deletion

## Scope

This package targets browser-like environments. Changes that require Node-only or server-only behavior should stay optional and must not break the browser-first API.

The renderer is still under active development. In particular, grouped objects, charts, table graphic frames, and complex SmartArt/diagram layouts are not fully implemented yet, so changes in those areas should be explicit about current behavior and known regressions.
