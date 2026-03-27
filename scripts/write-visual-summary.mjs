import { readdir, readFile } from "node:fs/promises";
import path from "node:path";

const baseDir = path.resolve(process.env.PPTX_VISUAL_RESULTS_DIR || path.join("test-results", "visual-regression"));

try {
  const caseDirs = (await readdir(baseDir, { withFileTypes: true }))
    .filter((entry) => entry.isDirectory())
    .map((entry) => entry.name)
    .sort();

  if (caseDirs.length === 0) {
    process.stdout.write(`## Visual Regression Summary\n\nNo visual regression metrics were found in \`${baseDir}\`.\n`);
    process.exit(0);
  }

  const lines = ["## Visual Regression Summary", ""];

  for (const caseId of caseDirs) {
    const metricsPath = path.join(baseDir, caseId, "metrics.json");
    const summary = JSON.parse(await readFile(metricsPath, "utf8"));

    lines.push(`### ${summary.fixture}`);
    lines.push("");
    lines.push(`- Slides compared: \`${summary.selectedSlides.join(", ")}\``);
    lines.push(`- Enforced: \`${summary.enforceThresholds}\``);
    lines.push(`- Max diff ratio: \`${Number(summary.maxObservedDiffRatio).toFixed(4)}\``);
    lines.push(`- Mean diff ratio: \`${Number(summary.meanObservedDiffRatio).toFixed(4)}\``);
    lines.push(
      `- Thresholds: max \`${Number(summary.thresholds.maxDiffRatio).toFixed(4)}\`, mean \`${Number(summary.thresholds.meanDiffRatio).toFixed(4)}\``,
    );
    lines.push("");
    lines.push("| Slide | Diff Ratio | Mismatch Pixels | Size |");
    lines.push("| --- | ---: | ---: | --- |");
    lines.push(
      ...summary.metrics.map(
        (metric) =>
          `| ${metric.slide} | ${Number(metric.diffRatio).toFixed(4)} | ${metric.mismatchPixels} | ${metric.comparedWidth}x${metric.comparedHeight} |`,
      ),
    );
    lines.push("");
  }

  process.stdout.write(lines.join("\n"));
} catch (error) {
  process.stdout.write(`## Visual Regression Summary\n\nUnable to read visual regression metrics in \`${baseDir}\`.\n`);
  if (process.env.CI !== "true") {
    console.error(error);
  }
}
