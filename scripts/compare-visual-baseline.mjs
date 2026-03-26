import { access, mkdir, readdir, readFile, writeFile } from "node:fs/promises";
import { constants } from "node:fs";
import path from "node:path";
import pixelmatch from "pixelmatch";
import { PNG } from "pngjs";

const baselineRoot = path.resolve(process.env.PPTX_VISUAL_BASELINE_DIR || ".tmp/visual-baseline");
const currentRoot = path.resolve(
  process.env.PPTX_VISUAL_CURRENT_DIR || path.join("test-results", "visual-regression"),
);
const outputRoot = path.resolve(
  process.env.PPTX_VISUAL_COMPARE_DIR || path.join("test-results", "visual-pr-comparison"),
);

await mkdir(outputRoot, { recursive: true });

const baselineDir = await resolveVisualRoot(baselineRoot);
const currentDir = await resolveVisualRoot(currentRoot);

const caseDirs = (await readdir(currentDir, { withFileTypes: true }))
  .filter((entry) => entry.isDirectory())
  .map((entry) => entry.name)
  .sort();

const comparison = {
  baselineDir,
  currentDir,
  generatedAt: new Date().toISOString(),
  cases: [],
};

for (const caseId of caseDirs) {
  const currentMetricsPath = path.join(currentDir, caseId, "metrics.json");
  const baselineMetricsPath = path.join(baselineDir, caseId, "metrics.json");
  const currentExists = await fileExists(currentMetricsPath);
  const baselineExists = await fileExists(baselineMetricsPath);

  if (!currentExists || !baselineExists) {
    comparison.cases.push({
      id: caseId,
      compared: false,
      reason: !baselineExists ? "missing-baseline" : "missing-current",
    });
    continue;
  }

  const currentMetrics = JSON.parse(await readFile(currentMetricsPath, "utf8"));
  const baselineMetrics = JSON.parse(await readFile(baselineMetricsPath, "utf8"));
  const outputCaseDir = path.join(outputRoot, caseId);

  await mkdir(outputCaseDir, { recursive: true });

  const slides = [];

  for (const metric of currentMetrics.metrics) {
    const slideLabel = `slide-${String(metric.slide).padStart(2, "0")}`;
    const currentActualPath = path.join(currentDir, caseId, `${slideLabel}.actual.png`);
    const baselineActualPath = path.join(baselineDir, caseId, `${slideLabel}.actual.png`);

    if (!(await fileExists(currentActualPath)) || !(await fileExists(baselineActualPath))) {
      continue;
    }

    const currentActual = PNG.sync.read(await readFile(currentActualPath));
    const baselineActual = PNG.sync.read(await readFile(baselineActualPath));
    const comparisonWidth = Math.min(currentActual.width, baselineActual.width);
    const comparisonHeight = Math.min(currentActual.height, baselineActual.height);
    const currentCropped = cropPng(currentActual, comparisonWidth, comparisonHeight);
    const baselineCropped = cropPng(baselineActual, comparisonWidth, comparisonHeight);
    const diff = new PNG({ width: comparisonWidth, height: comparisonHeight });
    const mismatchPixels = pixelmatch(
      currentCropped.data,
      baselineCropped.data,
      diff.data,
      comparisonWidth,
      comparisonHeight,
      {
        threshold: 0.1,
        includeAA: true,
      },
    );
    const totalPixels = comparisonWidth * comparisonHeight;
    const diffRatio = mismatchPixels / totalPixels;

    slides.push({
      slide: metric.slide,
      diffRatio,
      mismatchPixels,
      totalPixels,
      comparedWidth: comparisonWidth,
      comparedHeight: comparisonHeight,
    });

    await Promise.all([
      writeFile(path.join(outputCaseDir, `${slideLabel}.baseline.png`), PNG.sync.write(baselineCropped)),
      writeFile(path.join(outputCaseDir, `${slideLabel}.current.png`), PNG.sync.write(currentCropped)),
      writeFile(path.join(outputCaseDir, `${slideLabel}.diff.png`), PNG.sync.write(diff)),
    ]);
  }

  const maxDiffRatio = slides.reduce((max, slide) => Math.max(max, slide.diffRatio), 0);
  const meanDiffRatio = slides.reduce((sum, slide) => sum + slide.diffRatio, 0) / Math.max(slides.length, 1);
  const topSlides = [...slides].sort((left, right) => right.diffRatio - left.diffRatio).slice(0, 5);
  const summary = {
    caseId,
    compared: true,
    selectedSlides: currentMetrics.selectedSlides,
    slides,
    maxDiffRatio,
    meanDiffRatio,
    topSlides,
    baselineMetrics: {
      maxObservedDiffRatio: baselineMetrics.maxObservedDiffRatio,
      meanObservedDiffRatio: baselineMetrics.meanObservedDiffRatio,
    },
    currentMetrics: {
      maxObservedDiffRatio: currentMetrics.maxObservedDiffRatio,
      meanObservedDiffRatio: currentMetrics.meanObservedDiffRatio,
    },
  };

  comparison.cases.push(summary);
  await writeFile(path.join(outputCaseDir, "comparison.json"), JSON.stringify(summary, null, 2));
}

comparison.comparedCases = comparison.cases.filter((entry) => entry.compared).length;
comparison.totalCases = comparison.cases.length;
comparison.worstCases = [...comparison.cases]
  .filter((entry) => entry.compared)
  .sort((left, right) => right.maxDiffRatio - left.maxDiffRatio)
  .slice(0, 10)
  .map((entry) => ({
    caseId: entry.caseId,
    maxDiffRatio: entry.maxDiffRatio,
    meanDiffRatio: entry.meanDiffRatio,
  }));

await writeFile(path.join(outputRoot, "comparison-summary.json"), JSON.stringify(comparison, null, 2));
await writeFile(path.join(outputRoot, "comment.md"), buildComment(comparison));
await writeFile(path.join(outputRoot, "index.html"), buildHtml(comparison));

function buildComment(summary) {
  const lines = ["<!-- pptx-react-renderer-visual-pr-comment -->", "## Visual PR Diff", ""];

  if (summary.comparedCases === 0) {
    lines.push("No comparable visual baseline from `main` was found for this PR run.");
    return lines.join("\n");
  }

  lines.push(`Compared \`${summary.comparedCases}\` visual cases against the latest successful \`main\` visual artifact.`);
  lines.push("");
  lines.push("| Case | Max Diff | Mean Diff | Top Changed Slides |");
  lines.push("| --- | ---: | ---: | --- |");

  for (const entry of summary.worstCases) {
    const caseSummary = summary.cases.find((item) => item.caseId === entry.caseId);
    const topSlides = caseSummary.topSlides.map((slide) => `${slide.slide} (${slide.diffRatio.toFixed(4)})`).join(", ");
    lines.push(
      `| ${entry.caseId} | ${entry.maxDiffRatio.toFixed(4)} | ${entry.meanDiffRatio.toFixed(4)} | ${topSlides || "none"} |`,
    );
  }

  const missing = summary.cases.filter((entry) => !entry.compared).map((entry) => `${entry.id} (${entry.reason})`);
  if (missing.length > 0) {
    lines.push("");
    lines.push(`Missing baseline coverage: ${missing.join(", ")}`);
  }

  lines.push("");
  lines.push("Artifacts include an HTML side-by-side report under `test-results/visual-pr-comparison/index.html`.");
  return lines.join("\n");
}

function buildHtml(summary) {
  const cards = summary.cases
    .map((entry) => {
      if (!entry.compared) {
        return `
          <section class="case">
            <h2>${escapeHtml(entry.id)}</h2>
            <p>No comparison available: ${escapeHtml(entry.reason)}</p>
          </section>
        `;
      }

      const slides = entry.topSlides
        .map((slide) => {
          const slideLabel = `slide-${String(slide.slide).padStart(2, "0")}`;
          return `
            <section class="slide-card">
              <header>
                <h3>Slide ${slide.slide}</h3>
                <p>Diff ratio ${slide.diffRatio.toFixed(4)}</p>
              </header>
              <div class="images">
                <figure>
                  <img loading="lazy" src="./${entry.caseId}/${slideLabel}.baseline.png" alt="Baseline slide ${slide.slide}" />
                  <figcaption>Main baseline</figcaption>
                </figure>
                <figure>
                  <img loading="lazy" src="./${entry.caseId}/${slideLabel}.current.png" alt="Current slide ${slide.slide}" />
                  <figcaption>Current PR</figcaption>
                </figure>
                <figure>
                  <img loading="lazy" src="./${entry.caseId}/${slideLabel}.diff.png" alt="Diff slide ${slide.slide}" />
                  <figcaption>Diff</figcaption>
                </figure>
              </div>
            </section>
          `;
        })
        .join("\n");

      return `
        <section class="case">
          <header class="case-header">
            <div>
              <h2>${escapeHtml(entry.caseId)}</h2>
              <p>Compared slides: ${entry.selectedSlides.join(", ")}</p>
            </div>
            <div class="stats">
              <span>Max diff: ${entry.maxDiffRatio.toFixed(4)}</span>
              <span>Mean diff: ${entry.meanDiffRatio.toFixed(4)}</span>
            </div>
          </header>
          ${slides}
        </section>
      `;
    })
    .join("\n");

  return `<!DOCTYPE html>
<html lang="en">
  <head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <title>Visual PR Comparison</title>
    <style>
      :root {
        font-family: Inter, ui-sans-serif, system-ui, sans-serif;
        color-scheme: light;
        background: #f5f7fb;
        color: #152034;
      }
      * { box-sizing: border-box; }
      body { margin: 0; padding: 32px; }
      main { display: grid; gap: 24px; }
      .case { background: #fff; border-radius: 20px; padding: 24px; box-shadow: 0 24px 60px rgba(20, 32, 51, 0.08); }
      .case-header { display: flex; justify-content: space-between; gap: 24px; align-items: flex-start; }
      .stats { display: grid; gap: 6px; text-align: right; color: #42526b; }
      .slide-card { border-top: 1px solid #e3e8f0; padding-top: 20px; margin-top: 20px; }
      .images { display: grid; grid-template-columns: repeat(3, minmax(0, 1fr)); gap: 16px; margin-top: 16px; }
      figure { margin: 0; background: #f7f9fc; border: 1px solid #dbe3ef; border-radius: 16px; padding: 12px; }
      img { display: block; width: 100%; height: auto; border-radius: 10px; background: white; }
      figcaption { margin-top: 10px; text-align: center; color: #42526b; font-size: 14px; }
      @media (max-width: 1100px) { .images { grid-template-columns: 1fr; } .case-header { flex-direction: column; } .stats { text-align: left; } }
    </style>
  </head>
  <body>
    <main>
      <section class="case">
        <h1>Visual PR Comparison</h1>
        <p>Current PR renderer output compared against the latest successful <code>main</code> renderer artifact.</p>
      </section>
      ${cards}
    </main>
  </body>
</html>`;
}

async function resolveVisualRoot(root) {
  if (await pathExists(path.join(root, "slop-sample", "metrics.json"))) return root;
  if (await pathExists(path.join(root, "complex-sanitized", "metrics.json"))) return root;
  if (await pathExists(path.join(root, "test-results", "visual-regression", "slop-sample", "metrics.json"))) {
    return path.join(root, "test-results", "visual-regression");
  }
  if (await pathExists(path.join(root, "visual-regression", "slop-sample", "metrics.json"))) {
    return path.join(root, "visual-regression");
  }

  return root;
}

function cropPng(source, width, height) {
  if (source.width === width && source.height === height) {
    return source;
  }

  const cropped = new PNG({ width, height });

  for (let y = 0; y < height; y += 1) {
    const sourceStart = y * source.width * 4;
    const sourceEnd = sourceStart + width * 4;
    source.data.copy(cropped.data, y * width * 4, sourceStart, sourceEnd);
  }

  return cropped;
}

async function fileExists(filePath) {
  try {
    await readFile(filePath);
    return true;
  } catch {
    return false;
  }
}

async function pathExists(targetPath) {
  try {
    await access(targetPath, constants.R_OK);
    return true;
  } catch {
    return false;
  }
}

function escapeHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}
