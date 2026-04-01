function range(start, end) {
  return Array.from({ length: end - start + 1 }, (_, index) => start + index);
}

export const visualCases = [
  {
    id: "slop-sample",
    sourceFile: "test/fixtures/real/slop-sample.pptx",
    browserFixturePath: "/test/fixtures/real/slop-sample.pptx",
    slides: range(1, 12),
    thresholds: {
      maxDiffRatio: 0.9,
      meanDiffRatio: 0.6,
    },
    enforce: false,
  },
  {
    id: "complex-sanitized",
    sourceFile: "test/fixtures/real/complex-sanitized.pptx",
    browserFixturePath: "/test/fixtures/real/complex-sanitized.pptx",
    slides: [1, 4, 5, 8],
    thresholds: {
      maxDiffRatio: 0.45,
      meanDiffRatio: 0.25,
    },
    enforce: false,
  },
  {
    id: "text-layout",
    sourceFile: ".tmp/visual-fixtures/text-layout.pptx",
    browserFixturePath: "/.tmp/visual-fixtures/text-layout.pptx",
    slides: [1],
    thresholds: {
      maxDiffRatio: 0.35,
      meanDiffRatio: 0.2,
    },
    enforce: true,
  },
  {
    id: "bullets-numbering",
    sourceFile: ".tmp/visual-fixtures/bullets-numbering.pptx",
    browserFixturePath: "/.tmp/visual-fixtures/bullets-numbering.pptx",
    slides: [1],
    thresholds: {
      maxDiffRatio: 0.38,
      meanDiffRatio: 0.22,
    },
    enforce: true,
  },
  {
    id: "text-insets-anchor",
    sourceFile: ".tmp/visual-fixtures/text-insets-anchor.pptx",
    browserFixturePath: "/.tmp/visual-fixtures/text-insets-anchor.pptx",
    slides: [1],
    thresholds: {
      maxDiffRatio: 0.35,
      meanDiffRatio: 0.2,
    },
    enforce: true,
  },
  {
    id: "paragraph-spacing",
    sourceFile: ".tmp/visual-fixtures/paragraph-spacing.pptx",
    browserFixturePath: "/.tmp/visual-fixtures/paragraph-spacing.pptx",
    slides: [1],
    thresholds: {
      maxDiffRatio: 0.38,
      meanDiffRatio: 0.22,
    },
    enforce: true,
  },
  {
    id: "theme-colors",
    sourceFile: ".tmp/visual-fixtures/theme-colors.pptx",
    browserFixturePath: "/.tmp/visual-fixtures/theme-colors.pptx",
    slides: [1],
    thresholds: {
      maxDiffRatio: 0.35,
      meanDiffRatio: 0.2,
    },
    enforce: true,
  },
  {
    id: "theme-luminance",
    sourceFile: ".tmp/visual-fixtures/theme-luminance.pptx",
    browserFixturePath: "/.tmp/visual-fixtures/theme-luminance.pptx",
    slides: [1],
    thresholds: {
      maxDiffRatio: 0.35,
      meanDiffRatio: 0.2,
    },
    enforce: true,
  },
  {
    id: "theme-style-refs",
    sourceFile: ".tmp/visual-fixtures/theme-style-refs.pptx",
    browserFixturePath: "/.tmp/visual-fixtures/theme-style-refs.pptx",
    slides: [1],
    thresholds: {
      maxDiffRatio: 0.35,
      meanDiffRatio: 0.2,
    },
    enforce: true,
  },
  {
    id: "group-transform",
    sourceFile: ".tmp/visual-fixtures/group-transform.pptx",
    browserFixturePath: "/.tmp/visual-fixtures/group-transform.pptx",
    slides: [1],
    thresholds: {
      maxDiffRatio: 0.35,
      meanDiffRatio: 0.2,
    },
    enforce: true,
  },
  {
    id: "nested-groups",
    sourceFile: ".tmp/visual-fixtures/nested-groups.pptx",
    browserFixturePath: "/.tmp/visual-fixtures/nested-groups.pptx",
    slides: [1],
    thresholds: {
      maxDiffRatio: 0.42,
      meanDiffRatio: 0.24,
    },
    enforce: false,
  },
  {
    id: "table-layout",
    sourceFile: ".tmp/visual-fixtures/table-layout.pptx",
    browserFixturePath: "/.tmp/visual-fixtures/table-layout.pptx",
    slides: [1],
    thresholds: {
      maxDiffRatio: 0.4,
      meanDiffRatio: 0.22,
    },
    enforce: true,
  },
  {
    id: "image-placement",
    sourceFile: ".tmp/visual-fixtures/image-placement.pptx",
    browserFixturePath: "/.tmp/visual-fixtures/image-placement.pptx",
    slides: [1],
    thresholds: {
      maxDiffRatio: 0.35,
      meanDiffRatio: 0.2,
    },
    enforce: true,
  },
  {
    id: "image-cropping",
    sourceFile: ".tmp/visual-fixtures/image-cropping.pptx",
    browserFixturePath: "/.tmp/visual-fixtures/image-cropping.pptx",
    slides: [1],
    thresholds: {
      maxDiffRatio: 0.75,
      meanDiffRatio: 0.45,
    },
    enforce: false,
  },
  {
    id: "diagram-fallback",
    sourceFile: "test/fixtures/real/complex-sanitized.pptx",
    browserFixturePath: "/test/fixtures/real/complex-sanitized.pptx",
    slides: [2, 10],
    thresholds: {
      maxDiffRatio: 0.8,
      meanDiffRatio: 0.5,
    },
    enforce: false,
  },
  {
    id: "unsupported-chart",
    sourceFile: ".tmp/visual-fixtures/unsupported-chart.pptx",
    browserFixturePath: "/.tmp/visual-fixtures/unsupported-chart.pptx",
    slides: [1],
    thresholds: {
      maxDiffRatio: 0.95,
      meanDiffRatio: 0.7,
    },
    enforce: false,
  },
];

export function getSelectedVisualCases() {
  const requested = process.env.PPTX_VISUAL_CASES
    ?.split(",")
    .map((value) => value.trim())
    .filter(Boolean);

  if (!requested || requested.length === 0) {
    return visualCases;
  }

  return visualCases.filter((visualCase) => requested.includes(visualCase.id));
}
