import { mkdir, readFile, writeFile } from "node:fs/promises";
import path from "node:path";
import { expect, test } from "@playwright/test";
import pixelmatch from "pixelmatch";
import { PNG } from "pngjs";
import { getSelectedVisualCases } from "../scripts/visual-cases.mjs";

for (const visualCase of getSelectedVisualCases()) {
  test(`compares ${visualCase.id} against LibreOffice reference images`, async ({ page }) => {
    const referenceDir = path.resolve(".tmp", "visual-reference", visualCase.id);
    const reportDir = path.resolve("test-results", "visual-regression", visualCase.id);
    const shouldEnforce = process.env.PPTX_VISUAL_ENFORCE === "1" || visualCase.enforce;

    await mkdir(reportDir, { recursive: true });

    await page.goto(
      `/test/browser/harness/?fixture=${encodeURIComponent(visualCase.browserFixturePath)}&scale=1&slideNumbers=0&theme=light`,
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
        { message: `browser harness should finish rendering for ${visualCase.id}` },
      )
      .toMatchObject({ state: "ready", showSlideNumbers: false });

    const slides = page.locator(".pptx-slide");
    const minimumSlideCount = Math.max(...visualCase.slides);
    await expect
      .poll(async () => slides.count(), { message: `${visualCase.id} should render enough slides for comparison` })
      .toBeGreaterThanOrEqual(minimumSlideCount);

    const metrics = [];

    for (const slideNumber of visualCase.slides) {
      const slideLabel = `slide-${String(slideNumber).padStart(2, "0")}`;
      const actualBuffer = await slides.nth(slideNumber - 1).screenshot({ animations: "disabled" });
      const referenceBuffer = await readFile(path.join(referenceDir, `${slideLabel}.png`));
      const actual = PNG.sync.read(actualBuffer);
      const reference = PNG.sync.read(referenceBuffer);

      expect(
        { width: actual.width, height: actual.height },
        `${visualCase.id} ${slideLabel} dimensions should match the reference image`,
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
      fixture: visualCase.id,
      selectedSlides: visualCase.slides,
      enforceThresholds: shouldEnforce,
      thresholds: visualCase.thresholds,
      maxObservedDiffRatio,
      meanObservedDiffRatio,
      metrics,
    };

    await writeFile(path.join(reportDir, "metrics.json"), JSON.stringify(summary, null, 2));

    if (shouldEnforce) {
      expect(
        maxObservedDiffRatio,
        `${visualCase.id} max visual diff ratio ${maxObservedDiffRatio.toFixed(4)} exceeded ${visualCase.thresholds.maxDiffRatio.toFixed(4)}`,
      ).toBeLessThanOrEqual(visualCase.thresholds.maxDiffRatio);
      expect(
        meanObservedDiffRatio,
        `${visualCase.id} mean visual diff ratio ${meanObservedDiffRatio.toFixed(4)} exceeded ${visualCase.thresholds.meanDiffRatio.toFixed(4)}`,
      ).toBeLessThanOrEqual(visualCase.thresholds.meanDiffRatio);
    }
  });
}
