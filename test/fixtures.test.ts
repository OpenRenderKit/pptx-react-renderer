import { readFile } from "node:fs/promises";
import { resolve } from "node:path";
import { describe, expect, it } from "vitest";
import { parsePptx, renderSlides } from "../src/index";

const basicFixturePath = resolve("test/fixtures/real/basic-preview.pptx");
const complexFixturePath = resolve("test/fixtures/real/complex-sanitized.pptx");

function summarizeSlides(result: Awaited<ReturnType<typeof parsePptx>>) {
  return result.slides.map((slide) => ({
    number: slide.number,
    background: slide.background,
    elements: slide.elements.map((element) => {
      switch (element.type) {
        case "text":
          return {
            type: element.type,
            runs: element.runs.map((run) => run.text),
            paragraphs:
              element.paragraphs?.map((paragraph) => ({
                bullet: paragraph.bullet ?? false,
                texts: paragraph.runs.map((run) => ("text" in run ? run.text : run.type)),
              })) || [],
          };
        case "image":
          return {
            type: element.type,
            mimeType: element.mimeType,
            hasSrc: Boolean(element.src),
          };
        case "graphicFrame":
          return {
            type: element.type,
            graphicType: element.graphicType,
          };
        case "shape":
          return {
            type: element.type,
            shapeType: element.shapeType,
          };
        default:
          return {
            type: element.type,
          };
      }
    }),
  }));
}

function normalizeRenderedHtml(html: string): string {
  return html
    .replace(/data:[^"]+;base64,[^"]+/g, "data:<inline>")
    .replace(/\s+/g, " ")
    .trim();
}

describe("real PPTX fixtures", () => {
  it("parses the committed PPTX fixture into a stable model summary", async () => {
    const buffer = await readFile(basicFixturePath);
    const result = await parsePptx(buffer);

    expect(summarizeSlides(result)).toMatchInlineSnapshot(`
      [
        {
          "background": {
            "color": "#F8FAFC",
            "type": "solid",
          },
          "elements": [
            {
              "paragraphs": [
                {
                  "bullet": false,
                  "texts": [
                    "Quarterly Review",
                  ],
                },
                {
                  "bullet": true,
                  "texts": [
                    "Revenue is up 18%",
                    "break",
                    "1",
                  ],
                },
              ],
              "runs": [
                "Quarterly Review",
                "Revenue is up 18%",
              ],
              "type": "text",
            },
            {
              "hasSrc": true,
              "mimeType": "image/png",
              "type": "image",
            },
          ],
          "number": 1,
        },
        {
          "background": undefined,
          "elements": [
            {
              "shapeType": "rect",
              "type": "shape",
            },
            {
              "graphicType": "chart",
              "type": "graphicFrame",
            },
          ],
          "number": 2,
        },
      ]
    `);
  });

  it("renders the committed PPTX fixture into stable DOM output", async () => {
    const buffer = await readFile(basicFixturePath);
    const result = await parsePptx(buffer);
    const container = document.createElement("div");

    renderSlides(result.slides, {
      container,
      scale: 0.5,
      showSlideNumbers: true,
      theme: "light",
    });

    expect(normalizeRenderedHtml(container.innerHTML)).toMatchInlineSnapshot(
      `"<div class="pptx-renderer-wrapper"><div class="pptx-slide" style="position: relative; width: 480px; height: 360px; margin: 0px auto 40px; box-shadow: 0 4px 20px rgba(0,0,0,0.1); overflow: hidden; background-color: rgb(248, 250, 252);"><div class="pptx-slide-number">Slide 1</div><div class="pptx-text-element" style="left: 24px; top: 18px; width: 384px; height: 96px; position: absolute; box-sizing: border-box; font-family: Arial, sans-serif; text-align: left; color: rgb(17, 24, 39); display: flex; flex-direction: column; justify-content: center; word-wrap: break-word; overflow-wrap: break-word; white-space: pre-wrap; overflow: hidden; padding: 0px 1.333px 0px 0.6665px;"><div style="width: 100%; height: 100%; overflow: hidden;"><p style="margin: 0px; padding: 0px;"><span style="font-size: 18.666662000000002px; font-family: Arial, sans-serif; font-weight: 700; color: rgb(17, 24, 39);">Quarterly Review</span></p><p style="margin: 0px 0px 0px 23.999994px; text-indent: -11.999997px; display: flex; align-items: flex-start; gap: 0.125em; padding: 0px 0px 0px 0em; width: 100%; box-sizing: border-box;"><span style="flex-shrink: 0; width: 0.5em; text-align: center; user-select: none; font-family: Wingdings;">•</span><span style="flex: 1 1 0%; overflow: hidden; word-wrap: break-word;"><span style="font-size: 14.666663000000002px; font-family: Arial, sans-serif; font-weight: 400; color: rgb(55, 65, 81);">Revenue is up 18%</span><br><span style="font-size: 11.999997px; font-family: Arial, sans-serif; font-weight: 400; color: rgb(17, 24, 39);">1</span></span></p></div></div><div class="pptx-image-element" style="left: 360px; top: 36px; width: 72px; height: 72px; position: absolute;"><img src="data:<inline>" alt="" style="width: 100%; height: 100%; object-fit: contain;"></div></div><div class="pptx-slide" style="position: relative; width: 480px; height: 360px; margin: 0px auto 40px; box-shadow: 0 4px 20px rgba(0,0,0,0.1); overflow: hidden; background-color: rgb(255, 255, 255);"><div class="pptx-slide-number">Slide 2</div><div class="pptx-shape-element" style="left: 36px; top: 36px; width: 96px; height: 48px; position: absolute; box-sizing: border-box; background-color: transparent; border-radius: 0;"></div><div class="pptx-graphic-frame" style="left: 180px; top: 36px; width: 192px; height: 120px; position: absolute; background-color: rgba(200, 230, 255, 0.3); border: 2px dashed rgb(74, 144, 226); display: flex; align-items: center; justify-content: center;"><span style="color: #4a90e2; font-size: 14px;">📊 Chart</span></div></div></div>"`
    );
  });

  it("parses the sanitized complex fixture into a stable structural summary", async () => {
    const buffer = await readFile(complexFixturePath);
    const result = await parsePptx(buffer);

    const summary = {
      slides: result.slides.length,
      totalElements: result.slides.reduce((sum, slide) => sum + slide.elements.length, 0),
      perSlide: result.slides.map((slide) => ({
        number: slide.number,
        text: slide.elements.filter((el) => el.type === "text").length,
        image: slide.elements.filter((el) => el.type === "image").length,
        shape: slide.elements.filter((el) => el.type === "shape").length,
        table: slide.elements.filter((el) => el.type === "table").length,
        graphicFrame: slide.elements.filter((el) => el.type === "graphicFrame").length,
        graphicTypes: slide.elements
          .filter((el) => el.type === "graphicFrame")
          .map((el) => el.graphicType),
        textPreview: slide.elements
          .filter((el) => el.type === "text")
          .flatMap((el) => el.runs.slice(0, 2).map((run) => run.text))
          .filter((text) => text.trim().length > 0)
          .slice(0, 4),
      })),
    };

    expect(summary).toMatchInlineSnapshot(`
      {
        "perSlide": [
          {
            "graphicFrame": 0,
            "graphicTypes": [],
            "image": 1,
            "number": 1,
            "shape": 0,
            "table": 0,
            "text": 6,
            "textPreview": [
              "OpenBridge",
              ":",
              "Team:",
              "Alex Morgan",
            ],
          },
          {
            "graphicFrame": 1,
            "graphicTypes": [
              "diagram",
            ],
            "image": 1,
            "number": 2,
            "shape": 0,
            "table": 0,
            "text": 3,
            "textPreview": [
              "The Challenge:",
              "The Care Continuity Gap",
            ],
          },
          {
            "graphicFrame": 0,
            "graphicTypes": [],
            "image": 2,
            "number": 3,
            "shape": 1,
            "table": 0,
            "text": 2,
            "textPreview": [
              "Our Approach:",
              "The OpenBridge Network",
              "A connected system for ongoing care:",
              "Companion App:",
            ],
          },
          {
            "graphicFrame": 1,
            "graphicTypes": [
              "table",
            ],
            "image": 0,
            "number": 4,
            "shape": 0,
            "table": 0,
            "text": 5,
            "textPreview": [
              "Evaluation Matrix",
              "A single app or gadget only solves part of the experience. The",
              "OpenBridge Network™",
              "We reviewed multiple solution directions against the key project constraints:",
            ],
          },
          {
            "graphicFrame": 1,
            "graphicTypes": [
              "table",
            ],
            "image": 1,
            "number": 5,
            "shape": 0,
            "table": 0,
            "text": 6,
            "textPreview": [
              "Landscape Review:",
              "Most products solve one slice of the problem.",
              "OpenBridge combines the missing pieces.",
            ],
          },
          {
            "graphicFrame": 0,
            "graphicTypes": [],
            "image": 1,
            "number": 6,
            "shape": 0,
            "table": 0,
            "text": 4,
            "textPreview": [
              "Why OpenBridge is the Stronger Choice.",
              "The",
              "OpenBridge Network™",
            ],
          },
          {
            "graphicFrame": 0,
            "graphicTypes": [],
            "image": 6,
            "number": 7,
            "shape": 0,
            "table": 0,
            "text": 13,
            "textPreview": [
              "Road to Prototype:",
            ],
          },
          {
            "graphicFrame": 1,
            "graphicTypes": [
              "table",
            ],
            "image": 1,
            "number": 8,
            "shape": 0,
            "table": 0,
            "text": 3,
            "textPreview": [
              "Road to Prototype:",
              "Skills & Resources",
            ],
          },
          {
            "graphicFrame": 1,
            "graphicTypes": [
              "diagram",
            ],
            "image": 1,
            "number": 9,
            "shape": 1,
            "table": 0,
            "text": 1,
            "textPreview": [
              "The Key Risk:",
              "Unhealthy Reliance",
            ],
          },
          {
            "graphicFrame": 1,
            "graphicTypes": [
              "diagram",
            ],
            "image": 1,
            "number": 10,
            "shape": 0,
            "table": 0,
            "text": 3,
            "textPreview": [
              "Who benefits?",
            ],
          },
          {
            "graphicFrame": 0,
            "graphicTypes": [],
            "image": 1,
            "number": 11,
            "shape": 2,
            "table": 0,
            "text": 11,
            "textPreview": [
              "Build the Network.",
              "OpenBridge",
              "is more than a feature set; it is a coordination layer for a more responsive, more human model of care.",
            ],
          },
          {
            "graphicFrame": 0,
            "graphicTypes": [],
            "image": 0,
            "number": 12,
            "shape": 0,
            "table": 0,
            "text": 0,
            "textPreview": [],
          },
        ],
        "slides": 12,
        "totalElements": 83,
      }
    `);
  });

  it("renders the sanitized complex fixture into a stable DOM footprint", async () => {
    const buffer = await readFile(complexFixturePath);
    const result = await parsePptx(buffer);
    const container = document.createElement("div");

    renderSlides(result.slides, {
      container,
      scale: 0.3,
      showSlideNumbers: true,
      theme: "light",
    });

    expect({
      slides: container.querySelectorAll(".pptx-slide").length,
      textElements: container.querySelectorAll(".pptx-text-element").length,
      imageElements: container.querySelectorAll(".pptx-image-element").length,
      graphicFrames: container.querySelectorAll(".pptx-graphic-frame").length,
      hasOpenBridge: container.textContent?.includes("OpenBridge") ?? false,
      hasEvaluationMatrix: container.textContent?.includes("Evaluation Matrix") ?? false,
    }).toMatchInlineSnapshot(`
      {
        "graphicFrames": 6,
        "hasEvaluationMatrix": true,
        "hasOpenBridge": true,
        "imageElements": 16,
        "slides": 12,
        "textElements": 57,
      }
    `);
  });
});
