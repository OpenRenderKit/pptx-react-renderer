import { readFile } from "node:fs/promises";
import path from "node:path";

const metricsPath = path.resolve(
  process.env.PPTX_VISUAL_METRICS_PATH ||
    path.join("test-results", "visual-regression", "complex-sanitized", "metrics.json"),
);

try {
  const summary = JSON.parse(await readFile(metricsPath, "utf8"));

  const lines = [
    "## Visual Regression Summary",
    "",
    `- Fixture: \`${summary.fixture}\``,
    `- Slides compared: \`${summary.selectedSlides.join(", ")}\``,
    `- Enforced: \`${summary.enforceThresholds}\``,
    `- Max diff ratio: \`${Number(summary.maxObservedDiffRatio).toFixed(4)}\``,
    `- Mean diff ratio: \`${Number(summary.meanObservedDiffRatio).toFixed(4)}\``,
    `- Thresholds: max \`${Number(summary.thresholds.maxDiffRatio).toFixed(4)}\`, mean \`${Number(
      summary.thresholds.meanDiffRatio,
    ).toFixed(4)}\``,
    "",
    "| Slide | Diff Ratio | Mismatch Pixels | Size |",
    "| --- | ---: | ---: | --- |",
    ...summary.metrics.map(
      (metric) =>
        `| ${metric.slide} | ${Number(metric.diffRatio).toFixed(4)} | ${metric.mismatchPixels} | ${metric.width}x${metric.height} |`,
    ),
    "",
  ];

  process.stdout.write(lines.join("\n"));
} catch (error) {
  process.stdout.write(`## Visual Regression Summary\n\nUnable to read metrics at \`${metricsPath}\`.\n`);
  if (process.env.CI !== "true") {
    console.error(error);
  }
}
