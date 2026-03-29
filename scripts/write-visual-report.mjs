import { readdir, readFile, writeFile } from "node:fs/promises";
import path from "node:path";

const baseDir = path.resolve(process.env.PPTX_VISUAL_RESULTS_DIR || path.join("test-results", "visual-regression"));

try {
  const caseDirs = (await readdir(baseDir, { withFileTypes: true }))
    .filter((entry) => entry.isDirectory())
    .map((entry) => entry.name)
    .sort();

  if (caseDirs.length === 0) {
    process.exit(0);
  }

  const sections = [];
  const navCards = [];

  for (const caseId of caseDirs) {
    const metricsPath = path.join(baseDir, caseId, "metrics.json");
    const summary = JSON.parse(await readFile(metricsPath, "utf8"));
    navCards.push(`
      <a class="nav-card" href="#case-${caseId}">
        <strong>${escapeHtml(summary.fixture)}</strong>
        <span>Slides: ${summary.selectedSlides.join(", ")}</span>
        <span>Max diff: ${Number(summary.maxObservedDiffRatio).toFixed(4)}</span>
      </a>
    `);
    const cards = summary.metrics
      .map((metric) => {
        const slideLabel = `slide-${String(metric.slide).padStart(2, "0")}`;

        return `
          <section class="slide-card">
            <header>
              <h3>Slide ${metric.slide}</h3>
              <p>Diff ratio ${Number(metric.diffRatio).toFixed(4)} • ${metric.mismatchPixels.toLocaleString()} mismatched pixels</p>
              <p>Compared size ${metric.comparedWidth}×${metric.comparedHeight}</p>
            </header>
            <div class="images">
              <figure>
                <img loading="lazy" src="./${caseId}/${slideLabel}.reference.png" alt="Reference slide ${metric.slide}" />
                <figcaption>Reference</figcaption>
              </figure>
              <figure>
                <img loading="lazy" src="./${caseId}/${slideLabel}.actual.png" alt="Actual slide ${metric.slide}" />
                <figcaption>Renderer output</figcaption>
              </figure>
              <figure>
                <img loading="lazy" src="./${caseId}/${slideLabel}.diff.png" alt="Diff slide ${metric.slide}" />
                <figcaption>Diff</figcaption>
              </figure>
            </div>
          </section>
        `;
      })
      .join("\n");

    sections.push(`
      <section class="case" id="case-${caseId}">
        <header class="case-header">
          <div>
            <h2>${escapeHtml(summary.fixture)}</h2>
            <p>Slides compared: ${summary.selectedSlides.join(", ")}</p>
          </div>
          <div class="stats">
            <span>Enforced: ${summary.enforceThresholds}</span>
            <span>Max diff: ${Number(summary.maxObservedDiffRatio).toFixed(4)}</span>
            <span>Mean diff: ${Number(summary.meanObservedDiffRatio).toFixed(4)}</span>
            <span>Thresholds: ${Number(summary.thresholds.maxDiffRatio).toFixed(4)} / ${Number(summary.thresholds.meanDiffRatio).toFixed(4)}</span>
          </div>
        </header>
        ${cards}
      </section>
    `);
  }

  const html = `<!DOCTYPE html>
<html lang="en">
  <head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <title>Visual Regression Report</title>
    <style>
      :root {
        color-scheme: light;
        font-family: Inter, ui-sans-serif, system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
        background: #f5f7fb;
        color: #142033;
      }

      * {
        box-sizing: border-box;
      }

      body {
        margin: 0;
        padding: 32px;
      }

      main {
        display: grid;
        gap: 32px;
      }

      .case {
        background: white;
        border-radius: 20px;
        padding: 24px;
        box-shadow: 0 24px 60px rgba(20, 32, 51, 0.08);
      }

      .nav-grid {
        display: grid;
        grid-template-columns: repeat(auto-fit, minmax(220px, 1fr));
        gap: 14px;
      }

      .nav-card {
        display: grid;
        gap: 4px;
        padding: 16px;
        border-radius: 16px;
        border: 1px solid #dbe3ef;
        background: #f7f9fc;
        color: inherit;
        text-decoration: none;
      }

      .nav-card strong {
        font-size: 16px;
      }

      .nav-card span {
        color: #42526b;
        font-size: 14px;
      }

      .case-header {
        display: flex;
        justify-content: space-between;
        gap: 24px;
        align-items: flex-start;
        margin-bottom: 24px;
      }

      .case-header h2 {
        margin: 0 0 8px;
      }

      .case-header p,
      .case-header span {
        margin: 0;
        color: #42526b;
      }

      .stats {
        display: grid;
        gap: 6px;
        text-align: right;
      }

      .slide-card {
        border-top: 1px solid #e3e8f0;
        padding-top: 20px;
        margin-top: 20px;
      }

      .slide-card h3,
      .slide-card p {
        margin: 0 0 6px;
      }

      .images {
        display: grid;
        grid-template-columns: repeat(3, minmax(0, 1fr));
        gap: 16px;
        margin-top: 16px;
      }

      figure {
        margin: 0;
        background: #f7f9fc;
        border: 1px solid #dbe3ef;
        border-radius: 16px;
        padding: 12px;
      }

      img {
        display: block;
        width: 100%;
        height: auto;
        border-radius: 10px;
        background: white;
      }

      figcaption {
        margin-top: 10px;
        font-size: 14px;
        color: #42526b;
        text-align: center;
      }

      @media (max-width: 1100px) {
        .images {
          grid-template-columns: 1fr;
        }

        .case-header {
          flex-direction: column;
        }

        .stats {
          text-align: left;
        }
      }
    </style>
  </head>
  <body>
    <main>
      <section class="case">
        <header class="case-header">
          <div>
            <h1>Visual Regression Report</h1>
            <p>Reference images are rendered from PPTX with LibreOffice. Renderer images come from the current branch in Chromium.</p>
          </div>
        </header>
      </section>
      <section class="case">
        <h2>Cases</h2>
        <div class="nav-grid">
          ${navCards.join("\n")}
        </div>
      </section>
      ${sections.join("\n")}
    </main>
  </body>
</html>`;

  await writeFile(path.join(baseDir, "index.html"), html);
} catch (error) {
  if (process.env.CI !== "true") {
    console.error(error);
  }
  process.exit(0);
}

function escapeHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}
