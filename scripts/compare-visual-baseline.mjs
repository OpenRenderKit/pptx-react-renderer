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
  const baselineSlidesByNumber = new Map((baselineMetrics.metrics || []).map((metric) => [metric.slide, metric]));
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
    const branchDiffRatio = mismatchPixels / totalPixels;
    const baselineSlideMetrics = baselineSlidesByNumber.get(metric.slide);
    const baselineOfficeDiffRatio = Number(baselineSlideMetrics?.diffRatio || 0);
    const currentOfficeDiffRatio = Number(metric.diffRatio || 0);
    const officeDeltaDiffRatio = currentOfficeDiffRatio - baselineOfficeDiffRatio;

    slides.push({
      slide: metric.slide,
      branchDiffRatio,
      mismatchPixels,
      totalPixels,
      comparedWidth: comparisonWidth,
      comparedHeight: comparisonHeight,
      baselineOfficeDiffRatio,
      currentOfficeDiffRatio,
      officeDeltaDiffRatio,
    });

    await Promise.all([
      writeFile(path.join(outputCaseDir, `${slideLabel}.baseline.png`), PNG.sync.write(baselineCropped)),
      writeFile(path.join(outputCaseDir, `${slideLabel}.current.png`), PNG.sync.write(currentCropped)),
      writeFile(path.join(outputCaseDir, `${slideLabel}.diff.png`), PNG.sync.write(diff)),
    ]);
  }

  const maxDiffRatio = slides.reduce((max, slide) => Math.max(max, slide.branchDiffRatio), 0);
  const meanDiffRatio = slides.reduce((sum, slide) => sum + slide.branchDiffRatio, 0) / Math.max(slides.length, 1);
  const topSlides = [...slides].sort((left, right) => right.branchDiffRatio - left.branchDiffRatio).slice(0, 5);
  const officeMeanDelta =
    Number(currentMetrics.meanObservedDiffRatio || 0) - Number(baselineMetrics.meanObservedDiffRatio || 0);
  const officeMaxDelta =
    Number(currentMetrics.maxObservedDiffRatio || 0) - Number(baselineMetrics.maxObservedDiffRatio || 0);
  const topRegressedSlides = [...slides]
    .filter((slide) => slide.officeDeltaDiffRatio > 0)
    .sort((left, right) => right.officeDeltaDiffRatio - left.officeDeltaDiffRatio)
    .slice(0, 5);
  const topImprovedSlides = [...slides]
    .filter((slide) => slide.officeDeltaDiffRatio < 0)
    .sort((left, right) => left.officeDeltaDiffRatio - right.officeDeltaDiffRatio)
    .slice(0, 5);
  const summary = {
    caseId,
    compared: true,
    selectedSlides: currentMetrics.selectedSlides,
    slides,
    maxDiffRatio,
    meanDiffRatio,
    topSlides,
    officeMeanDelta,
    officeMaxDelta,
    topRegressedSlides,
    topImprovedSlides,
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
comparison.changedCases = comparison.cases.filter((entry) => entry.compared && entry.maxDiffRatio > 0).length;
comparison.regressedCases = comparison.cases
  .filter((entry) => entry.compared && entry.officeMeanDelta > 0)
  .sort((left, right) => right.officeMeanDelta - left.officeMeanDelta)
  .slice(0, 10)
  .map((entry) => ({
    caseId: entry.caseId,
    officeMeanDelta: entry.officeMeanDelta,
    officeMaxDelta: entry.officeMaxDelta,
    maxDiffRatio: entry.maxDiffRatio,
  }));
comparison.improvedCases = comparison.cases
  .filter((entry) => entry.compared && entry.officeMeanDelta < 0)
  .sort((left, right) => left.officeMeanDelta - right.officeMeanDelta)
  .slice(0, 10)
  .map((entry) => ({
    caseId: entry.caseId,
    officeMeanDelta: entry.officeMeanDelta,
    officeMaxDelta: entry.officeMaxDelta,
    maxDiffRatio: entry.maxDiffRatio,
  }));
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
  const runUrl = process.env.PPTX_VISUAL_RUN_URL;
  const pagesBaseUrl = process.env.PPTX_VISUAL_PAGES_BASE_URL;
  const prNumber = process.env.PPTX_VISUAL_PR_NUMBER;
  const reportBaseUrl =
    pagesBaseUrl && prNumber
      ? `${pagesBaseUrl.replace(/\/$/, "")}/pr-previews/${prNumber}`
      : undefined;

  if (summary.comparedCases === 0) {
    lines.push("No comparable visual baseline from `main` was found for this PR run.");
    if (runUrl) {
      lines.push("");
      lines.push(`[Open workflow run](${runUrl})`);
    }
    return lines.join("\n");
  }

  if (summary.changedCases === 0) {
    lines.push(
      `Compared \`${summary.comparedCases}\` visual cases against the latest successful \`main\` visual artifact and detected no renderer-output changes.`,
    );
  } else {
    lines.push(
      `Compared \`${summary.comparedCases}\` visual cases against the latest successful \`main\` visual artifact and found \`${summary.changedCases}\` changed cases.`,
    );
  }
  lines.push("");

  if (runUrl) {
    lines.push(`- [Workflow run](${runUrl})`);
  }
  if (reportBaseUrl) {
    lines.push(`- [PR vs main report](${reportBaseUrl}/visual-pr-comparison/index.html)`);
    lines.push(`- [Office reference vs current report](${reportBaseUrl}/visual-regression/index.html)`);
  }
  lines.push("");

  lines.push("### Renderer Change Ranking");
  lines.push("");
  lines.push("| Case | Max Diff | Mean Diff | Top Changed Slides |");
  lines.push("| --- | ---: | ---: | --- |");

  for (const entry of summary.worstCases) {
    const caseSummary = summary.cases.find((item) => item.caseId === entry.caseId);
    const topSlides = caseSummary.topSlides
      .map((slide) => `${slide.slide} (${slide.branchDiffRatio.toFixed(4)})`)
      .join(", ");
    const caseLink = reportBaseUrl
      ? `[${entry.caseId}](${reportBaseUrl}/visual-pr-comparison/index.html#case-${entry.caseId})`
      : entry.caseId;
    lines.push(
      `| ${caseLink} | ${entry.maxDiffRatio.toFixed(4)} | ${entry.meanDiffRatio.toFixed(4)} | ${topSlides || "none"} |`,
    );
  }

  const missing = summary.cases.filter((entry) => !entry.compared).map((entry) => `${entry.id} (${entry.reason})`);
  if (missing.length > 0) {
    lines.push("");
    lines.push(`Missing baseline coverage: ${missing.join(", ")}`);
  }

  lines.push("");

  if (summary.regressedCases.length > 0) {
    lines.push("### Most Regressed Vs Office Reference");
    lines.push("");
    lines.push("| Case | Mean Delta | Max Delta | Worst Slide Delta |");
    lines.push("| --- | ---: | ---: | --- |");

    for (const entry of summary.regressedCases.slice(0, 5)) {
      const caseSummary = summary.cases.find((item) => item.caseId === entry.caseId);
      const worstSlide = caseSummary.topRegressedSlides[0];
      const caseLink = reportBaseUrl
        ? `[${entry.caseId}](${reportBaseUrl}/visual-regression/index.html#case-${entry.caseId})`
        : entry.caseId;
      lines.push(
        `| ${caseLink} | +${entry.officeMeanDelta.toFixed(4)} | +${entry.officeMaxDelta.toFixed(4)} | ${
          worstSlide ? `${worstSlide.slide} (+${worstSlide.officeDeltaDiffRatio.toFixed(4)})` : "none"
        } |`,
      );
    }

    lines.push("");
  }

  if (summary.improvedCases.length > 0) {
    lines.push("### Most Improved Vs Office Reference");
    lines.push("");
    lines.push("| Case | Mean Delta | Max Delta | Best Slide Delta |");
    lines.push("| --- | ---: | ---: | --- |");

    for (const entry of summary.improvedCases.slice(0, 5)) {
      const caseSummary = summary.cases.find((item) => item.caseId === entry.caseId);
      const bestSlide = caseSummary.topImprovedSlides[0];
      const caseLink = reportBaseUrl
        ? `[${entry.caseId}](${reportBaseUrl}/visual-regression/index.html#case-${entry.caseId})`
        : entry.caseId;
      lines.push(
        `| ${caseLink} | ${entry.officeMeanDelta.toFixed(4)} | ${entry.officeMaxDelta.toFixed(4)} | ${
          bestSlide ? `${bestSlide.slide} (${bestSlide.officeDeltaDiffRatio.toFixed(4)})` : "none"
        } |`,
      );
    }

    lines.push("");
  }

  const previewCases = summary.worstCases.filter((entry) => entry.maxDiffRatio > 0).slice(0, 3);
  if (previewCases.length > 0 && reportBaseUrl) {
    lines.push("### Preview Slides");
    lines.push("");

    for (const entry of previewCases) {
      const caseSummary = summary.cases.find((item) => item.caseId === entry.caseId);
      const topSlide = caseSummary.topSlides[0];
      const slideLabel = `slide-${String(topSlide.slide).padStart(2, "0")}`;
      const caseUrl = `${reportBaseUrl}/visual-pr-comparison/${entry.caseId}`;
      lines.push(
        `**[${entry.caseId}](${reportBaseUrl}/visual-pr-comparison/index.html#case-${entry.caseId})** slide ${topSlide.slide}`,
      );
      lines.push("");
      lines.push("<table>");
      lines.push("<tr><th>Main</th><th>PR</th><th>Diff</th></tr>");
      lines.push(
        `<tr><td><img src="${caseUrl}/${slideLabel}.baseline.png" width="260" alt="Main baseline ${entry.caseId} slide ${topSlide.slide}"></td><td><img src="${caseUrl}/${slideLabel}.current.png" width="260" alt="PR current ${entry.caseId} slide ${topSlide.slide}"></td><td><img src="${caseUrl}/${slideLabel}.diff.png" width="260" alt="Diff ${entry.caseId} slide ${topSlide.slide}"></td></tr>`,
      );
      lines.push("</table>");
      lines.push("");
    }
  } else if (reportBaseUrl) {
    lines.push("No renderer deltas were detected in this PR. Open the published reports for full slide-by-slide inspection.");
    lines.push("");
  }

  lines.push("Artifacts include HTML side-by-side reports under `visual-regression-artifacts` and `visual-pr-comparison`.");
  return lines.join("\n");
}

function buildHtml(summary) {
  const navCards = summary.cases
    .filter((entry) => entry.compared)
    .map(
      (entry) => `
        <a class="nav-card" href="#case-${entry.caseId}">
          <strong>${escapeHtml(entry.caseId)}</strong>
          <span>Renderer diff ${entry.maxDiffRatio.toFixed(4)}</span>
          <span>Office delta ${entry.officeMeanDelta >= 0 ? "+" : ""}${entry.officeMeanDelta.toFixed(4)}</span>
        </a>
      `,
    )
    .join("\n");

  const cards = summary.cases
    .map((entry) => {
      if (!entry.compared) {
        return `
          <section class="case" id="case-${entry.id}">
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
                <p>Renderer diff ${slide.branchDiffRatio.toFixed(4)}</p>
                <p>Office delta ${slide.officeDeltaDiffRatio >= 0 ? "+" : ""}${slide.officeDeltaDiffRatio.toFixed(4)}</p>
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
        <section class="case" id="case-${entry.caseId}">
          <header class="case-header">
            <div>
              <h2>${escapeHtml(entry.caseId)}</h2>
              <p>Compared slides: ${entry.selectedSlides.join(", ")}</p>
            </div>
            <div class="stats">
              <span>Renderer max diff: ${entry.maxDiffRatio.toFixed(4)}</span>
              <span>Renderer mean diff: ${entry.meanDiffRatio.toFixed(4)}</span>
              <span>Office mean delta: ${entry.officeMeanDelta >= 0 ? "+" : ""}${entry.officeMeanDelta.toFixed(4)}</span>
              <span>Office max delta: ${entry.officeMaxDelta >= 0 ? "+" : ""}${entry.officeMaxDelta.toFixed(4)}</span>
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
      .nav-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(220px, 1fr)); gap: 14px; }
      .nav-card { display: grid; gap: 4px; padding: 16px; border-radius: 16px; border: 1px solid #dbe3ef; background: #f7f9fc; color: inherit; text-decoration: none; }
      .nav-card strong { font-size: 16px; }
      .nav-card span { color: #42526b; font-size: 14px; }
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
      <section class="case">
        <h2>Cases</h2>
        <div class="nav-grid">
          ${navCards}
        </div>
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
