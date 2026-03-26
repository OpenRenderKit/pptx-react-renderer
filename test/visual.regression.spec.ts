import { mkdir, readFile, writeFile } from "node:fs/promises";
import path from "node:path";
import { expect, test } from "@playwright/test";
import pixelmatch from "pixelmatch";
import { PNG } from "pngjs";

const fixtureName = process.env.PPTX_VISUAL_FIXTURE_NAME || "complex-sanitized";
const fixturePath = `/test/fixtures/real/${fixtureName}.pptx`;
const referenceDir = path.resolve(
  process.env.PPTX_VISUAL_REFERENCE_DIR || path.join(".tmp", "visual-reference", fixtureName),
);
const reportDir = path.resolve("test-results", "visual-regression", fixtureName);
const selectedSlides = (process.env.PPTX_VISUAL_SLIDES || "1,4,5,8")
  .split(",")
  .map((value) => Number.parseInt(value.trim(), 10))
  .filter((value) => Number.isFinite(value) && value > 0);
const maxDiffRatio = Number.parseFloat(process.env.PPTX_VISUAL_MAX_DIFF_RATIO || "0.45");
const meanDiffRatio = Number.parseFloat(process.env.PPTX_VISUAL_MEAN_DIFF_RATIO || "0.25");
const enforceThresholds = process.env.PPTX_VISUAL_ENFORCE === "1";

test("compares rendered slides against LibreOffice reference images", async ({ page }) => {
  await mkdir(reportDir, { recursive: true });

  await page.goto(
    `/test/browser/harness/?fixture=${encodeURIComponent(fixturePath)}&scale=1&slideNumbers=0&theme=light`,
  );
  await page.addStyleTag({
    content: `
      html, body { margin: 0; padding: 0; background: white; }
      #preview, .pptx-renderer-wrapper { margin: 0 !important; padding: 0 !important; }
      .pptx-slide { margin: 0 !important; box-shadow: none !important; }
    `,
  });

  await expect
    .poll(
      async () =>
        page.evaluate(
          () => (window as Window & { __PPTX_RENDER_STATUS__?: unknown }).__PPTX_RENDER_STATUS__,
        ),
      { message: "browser harness should finish rendering for visual regression" },
    )
    .toMatchObject({ state: "ready", slideCount: 12, showSlideNumbers: false });

  const slides = page.locator(".pptx-slide");
  await expect(slides).toHaveCount(12);

  const metrics = [];

  for (const slideNumber of selectedSlides) {
    const slideLabel = `slide-${String(slideNumber).padStart(2, "0")}`;
    const actualBuffer = await slides.nth(slideNumber - 1).screenshot({ animations: "disabled" });
    const referenceBuffer = await readFile(path.join(referenceDir, `${slideLabel}.png`));
    const actual = PNG.sync.read(actualBuffer);
    const reference = PNG.sync.read(referenceBuffer);

    expect(
      { width: actual.width, height: actual.height },
      `${slideLabel} dimensions should match the reference image`,
    ).toEqual({
      width: reference.width,
      height: reference.height,
    });

    const diff = new PNG({ width: actual.width, height: actual.height });
    const mismatchPixels = pixelmatch(
      actual.data,
      reference.data,
      diff.data,
      actual.width,
      actual.height,
      {
        threshold: 0.1,
        includeAA: true,
      },
    );
    const totalPixels = actual.width * actual.height;
    const diffRatio = mismatchPixels / totalPixels;

    metrics.push({
      slide: slideNumber,
      width: actual.width,
      height: actual.height,
      mismatchPixels,
      totalPixels,
      diffRatio,
    });

    await Promise.all([
      writeFile(path.join(reportDir, `${slideLabel}.actual.png`), actualBuffer),
      writeFile(path.join(reportDir, `${slideLabel}.reference.png`), referenceBuffer),
      writeFile(path.join(reportDir, `${slideLabel}.diff.png`), PNG.sync.write(diff)),
    ]);
  }

  const maxObservedDiffRatio = metrics.reduce((currentMax, metric) => Math.max(currentMax, metric.diffRatio), 0);
  const meanObservedDiffRatio =
    metrics.reduce((sum, metric) => sum + metric.diffRatio, 0) / Math.max(metrics.length, 1);

  const summary = {
    fixture: fixtureName,
    selectedSlides,
    enforceThresholds,
    thresholds: {
      maxDiffRatio,
      meanDiffRatio,
    },
    maxObservedDiffRatio,
    meanObservedDiffRatio,
    metrics,
  };

  await writeFile(path.join(reportDir, "metrics.json"), JSON.stringify(summary, null, 2));

  if (enforceThresholds) {
    expect(
      maxObservedDiffRatio,
      `max visual diff ratio ${maxObservedDiffRatio.toFixed(4)} exceeded ${maxDiffRatio.toFixed(4)}`,
    ).toBeLessThanOrEqual(maxDiffRatio);
    expect(
      meanObservedDiffRatio,
      `mean visual diff ratio ${meanObservedDiffRatio.toFixed(4)} exceeded ${meanDiffRatio.toFixed(4)}`,
    ).toBeLessThanOrEqual(meanDiffRatio);
  }
});
