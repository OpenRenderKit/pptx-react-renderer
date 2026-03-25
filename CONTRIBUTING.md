# Contributing

## Development Setup

```bash
pnpm install
pnpm run verify
```

## Development Workflow

- keep changes focused and backward-compatible unless the change is intentionally breaking
- add or update tests when parser or renderer behavior changes
- run `pnpm run verify` before opening a pull request
- run `pnpm run pack:check` if you change package metadata or published files

## Release Process

1. Update `CHANGELOG.md`.
2. Bump the version in `package.json`.
3. Run `pnpm run verify`.
4. Commit the release changes.
5. Create and push a Git tag like `v0.1.0`.
6. Publish from GitHub Actions after npm trusted publishing is configured.

## Scope

This package targets browser-like environments. Changes that require Node-only or server-only behavior should stay optional and must not break the browser-first API.
