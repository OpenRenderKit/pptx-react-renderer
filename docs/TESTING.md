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
  - Visual-regression run against reference slide images generated from the committed PPTX fixture.
  - Compares selected rendered slides against LibreOffice-derived PNGs using pixel diffs.

## Visual Regression Model

The visual-regression path is intentionally based on the real fixture deck rather than hand-authored screenshots.

1. `scripts/render-reference-images.mjs` converts `test/fixtures/real/complex-sanitized.pptx` into a PDF with LibreOffice.
2. The PDF is rasterized to per-slide PNGs with `pdftoppm` at 96 DPI so the image dimensions match the renderer's 96-DPI slide model.
3. `test/visual.regression.spec.ts` renders the same PPTX through `pptx-react-renderer` in Chromium.
4. Playwright screenshots selected slides and compares them to the reference PNGs with `pixelmatch`.
5. Metrics, actual images, reference images, and diff images are written under `test-results/visual-regression/`.

This is not the same as comparing against native Microsoft PowerPoint output. On GitHub Actions we can reliably run LibreOffice on Linux; we cannot reliably run PowerPoint itself. The point of the test is to provide a realistic office-rendered baseline that is reproducible in CI.

## Why Only Selected Slides

The committed complex fixture includes SmartArt, diagrams, grouped content, images, and tables. Some slides still contain features that the renderer intentionally does not support well enough yet, especially chart and SmartArt fidelity.

The initial visual-regression suite compares a curated subset of slides:

- `1`
- `4`
- `5`
- `8`

These are useful because they stress text placement, images, tables, and overall layout without being dominated by currently unsupported chart behavior.

As fidelity improves, expand the compared slide set and tighten thresholds.

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

## CI Behavior

The `visual-regression` GitHub Actions job:

- installs LibreOffice and Poppler
- generates reference images from the committed PPTX fixture
- runs the Playwright visual comparison
- uploads diff artifacts and writes a job summary

Threshold enforcement is currently metric-only by default in CI. The job measures max and mean pixel-diff ratios but does not hard-fail on them until the thresholds are calibrated against real runs.

When the measured values stabilize, set `PPTX_VISUAL_ENFORCE=1` in CI and tighten the thresholds.
