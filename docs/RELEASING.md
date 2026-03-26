# Releasing

## Current Release Flow

`pptx-react-renderer` publishes from GitHub Actions when pull requests merge into `main`.

The publish workflow is:

1. A pull request targets `main`.
2. The pull request carries exactly one release label:
   - `release:patch`
   - `release:minor`
   - `release:major`
   - `release:skip`
3. The PR must pass `verify` and `release-intent`.
4. On merge to `main`, `.github/workflows/publish.yml`:
   - resolves the current published npm version
   - computes the next release version
   - publishes to npm through trusted publishing
   - pushes a matching Git tag
   - creates a GitHub release

## Pre-1.0 Version Rules

Until `1.0.0`, the release labels are intentionally softened:

- `release:patch` => patch bump
- `release:minor` => patch bump
- `release:major` => minor bump
- `release:skip` => no publish

Examples from a current version of `0.1.1`:

- `release:patch` => `0.1.2`
- `release:minor` => `0.1.2`
- `release:major` => `0.2.0`

Once the package reaches `1.0.0`, the workflow uses normal semver behavior.

## Trusted Publishing

Trusted publishing is configured on npm for:

- package: `pptx-react-renderer`
- repository: `OpenRenderKit/pptx-react-renderer`
- workflow file: `publish.yml`

This means the steady-state release path does not require an `NPM_TOKEN` GitHub secret.

## If Publishing Breaks

1. Check the `publish` workflow logs in GitHub Actions.
2. Confirm the merged PR has exactly one release label.
3. Confirm npm still shows the trusted publisher:

```bash
npx npm@latest trust list pptx-react-renderer --json
```

4. Confirm the package exists and the latest version is what the workflow expects:

```bash
npm view pptx-react-renderer version dist-tags --json
```

5. Only if trusted publishing is broken and a release is urgent, use a short-lived fallback token path and remove it afterward.

## Current First Public Release

- npm: `0.1.1`
- GitHub release: `v0.1.1`
