export const visualCases = [
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
    enforce: false,
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
    enforce: false,
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
    enforce: false,
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
