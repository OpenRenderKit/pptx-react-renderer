# Repository Guidance

## Project State

- Package name: `pptx-react-renderer`
- npm package: `pptx-react-renderer@0.1.1`
- GitHub org/repo: `OpenRenderKit/pptx-react-renderer`
- This library is browser-first and still heavily under development.
- The long-term product goal is PowerPoint-like visual fidelity for text, SmartArt, images, placement, and color.

## Release Model

- `main` is the release branch.
- Normal code changes should go through a feature branch and pull request.
- Every PR must carry exactly one release label:
  - `release:patch`
  - `release:minor`
  - `release:major`
  - `release:skip`
- On merge to `main`, GitHub Actions resolves the next version from npm, publishes to npm, pushes a `vX.Y.Z` tag, and creates a GitHub release.
- Before `1.0.0`, release labels are intentionally softened:
  - `release:patch` => patch bump
  - `release:minor` => patch bump
  - `release:major` => minor bump
  - `release:skip` => no publish
- Trusted publishing is already configured for npm using `.github/workflows/publish.yml`.
- Do not add or depend on `NPM_TOKEN` unless trusted publishing breaks and a temporary bootstrap path is explicitly needed.

## Branching Expectations

- Maintain repo/admin documentation can be pushed directly to `main` when explicitly requested.
- Parser, renderer, testing, feature, or behavior changes should be done on a feature branch and opened as a PR.
- Keep `main` releasable.

## Verification

Run these before pushing meaningful changes:

```bash
pnpm run fixtures:generate
pnpm run verify
pnpm run pack:check
```

Run this when browser rendering, bundling, or fixture handling changes:

```bash
pnpm run test:browser
```

Run this when you need a real visual comparison against office-rendered slides:

```bash
pnpm run test:visual
```

## Fixtures

- Realistic committed fixtures live under `test/fixtures/real/`.
- `test/fixtures/real/complex-sanitized.pptx` is the main regression deck for SmartArt, tables, groups, diagrams, and mixed media.
- Visual-regression reference images are generated locally and in CI from the committed fixture; do not commit generated reference PNGs.
- Regenerate the sanitized complex fixture only from a local source deck using:

```bash
PPTX_COMPLEX_SOURCE="/absolute/path/to/source-deck.pptx" pnpm run fixtures:sanitize-complex
```

- Do not commit private source presentations.

## Known Technical Gaps

- Grouped shapes are not fully reconstructed.
- `graphicFrame` handling is still partial, especially charts.
- SmartArt/diagram fidelity is incomplete.
- Theme/color fidelity is still shallow compared with OOXML.
- Custom geometry support skips important cases such as arc conversion.
- Rendering is not yet 1:1 with PowerPoint.

## Coding Notes

- Prefer browser-safe behavior. Do not introduce Node-only runtime assumptions into the public API.
- Treat OOXML support work as correctness work first, polish work second.
- Add regression tests for every behavior change, ideally using the committed fixtures.
- When researching OOXML behavior, prefer primary sources such as ECMA-376 / ISO OOXML references, Microsoft Open Specifications, and official Open XML SDK documentation.
