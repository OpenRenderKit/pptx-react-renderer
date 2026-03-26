# Testing Strategy

This repo uses several test layers because parser correctness alone is not enough for a renderer whose target is visual fidelity.

## Test Layers

- `pnpm run test:unit`
  - Vitest parser and renderer unit tests.
  - Synthetic PPTX fixtures for focused behavior checks.
- `pnpm run test:browser`
  - Playwright smoke coverage against the committed sanitized deck.
  - Confirms the built package renders in a real browser.
- `pnpm run test:visual`
  - Case-based visual-regression run against reference slide images generated from real PPTX fixtures.
  - Compares selected rendered slides against LibreOffice-derived PNGs using pixel diffs.

## Visual Regression Model

The visual-regression path is intentionally based on real PPTX fixtures rather than hand-authored screenshots.

1. `scripts/generate-visual-fixtures.mjs` creates focused edge-case decks under `.tmp/visual-fixtures/`.
2. `scripts/render-reference-images.mjs` converts each selected PPTX fixture into a PDF with LibreOffice.
3. The PDF is rasterized to per-slide PNGs with `pdftoppm` at 96 DPI so the image dimensions match the renderer's 96-DPI slide model.
4. `test/visual.regression.spec.ts` renders the same PPTX through `pptx-react-renderer` in Chromium.
5. Playwright screenshots selected slides and compares them to the reference PNGs with `pixelmatch`.
6. Metrics, actual images, reference images, and diff images are written under `test-results/visual-regression/`.

This is not the same as comparing against native Microsoft PowerPoint output. On GitHub Actions we can reliably run LibreOffice on Linux; we cannot reliably run PowerPoint itself. The point of the test is to provide a realistic office-rendered baseline that is reproducible in CI.

## Visual Cases

The visual suite now mixes:

- the committed complex regression deck
- focused generated decks for:
  - text layout
  - bullets and auto-numbering
  - text insets and vertical anchoring
  - paragraph spacing and line spacing
  - theme colors
  - theme luminance transforms
  - grouped transforms
  - nested groups
  - table layout
  - image placement
  - image cropping
  - unsupported chart behavior

Each case compares only the slides that are meaningful for that case.

The suite also includes explicit metric-only tracking cases for:

- SmartArt / diagram fallback behavior from the committed complex deck
- nested grouped transforms that are not stable enough to gate yet
- image cropping and unsupported charts, where the library is intentionally incomplete today

The committed complex fixture still includes SmartArt, diagrams, grouped content, images, and tables. Some slides contain features that the renderer intentionally does not support well enough yet, especially chart and SmartArt fidelity.

The current complex-deck comparison stays limited to a curated subset:

- `1`
- `4`
- `5`
- `8`

These are useful because they stress text placement, images, tables, and overall layout without being dominated by currently unsupported chart behavior.

As fidelity improves, add more generated cases, expand the compared slide sets, and tighten thresholds.

## Local Prerequisites For Visual Regression

`pnpm run test:visual` requires:

- LibreOffice with `soffice`
- Poppler with `pdftoppm`

Examples:

```bash
# macOS
brew install --cask libreoffice
brew install poppler

# Ubuntu / Debian
sudo apt-get update
sudo apt-get install -y libreoffice poppler-utils
```

If `soffice` is not on `PATH`, set `SOFFICE_BIN`.

The reference-generation script also checks common macOS LibreOffice locations automatically, and `pdftoppm` can be overridden with `PDFTOPPM_BIN` if Poppler is installed outside the default `PATH`.

## CI Behavior

The `visual-regression` GitHub Actions job:

- installs LibreOffice and Poppler
- generates reference images from the committed PPTX fixture
- runs the Playwright visual comparison
- uploads diff artifacts and writes a job summary

CI currently runs visual regression in metric-only mode for normal development and release work. That means pull requests and routine version bumps still get full visual artifacts, diff images, and summary metrics without blocking merges.

Hard threshold enforcement is still available through `PPTX_VISUAL_ENFORCE=1` for local work and for future major release gates such as `1.0.0` or `2.0.0`, where a stricter visual freeze makes more sense.

As fidelity improves, keep tightening thresholds, and reserve CI-blocking visual gates for explicit major-release checkpoints rather than every `0.x` or `x.y.z` release.
