import { describe, expect, it } from "vitest";
import { parsePptx, renderPptx, renderSlides } from "../src/index";
import { createTestPptx } from "./helpers";

describe("pptx-react-renderer", () => {
  it("parses slide text and dimensions from a minimal deck", async () => {
    const buffer = await createTestPptx([
      { text: ["Quarterly Review", "Revenue is up 18%"], backgroundColor: "#F8FAFC" },
    ]);

    const result = await parsePptx(buffer);

    expect(result.slideWidth).toBe(960);
    expect(result.slideHeight).toBe(720);
    expect(result.slides).toHaveLength(1);
    expect(result.slides[0].background?.color).toBe("#F8FAFC");
    expect(result.slides[0].elements[0]).toMatchObject({
      type: "text",
      runs: [
        { text: "Quarterly Review" },
        { text: "Revenue is up 18%" },
      ],
    });
  });

  it("renders parsed slides into a target container", async () => {
    const buffer = await createTestPptx([{ text: ["Deck Title", "Second line"] }]);
    const { slides } = await parsePptx(buffer);
    const container = document.createElement("div");

    renderSlides(slides, {
      container,
      scale: 0.5,
      showSlideNumbers: true,
      theme: "light",
    });

    expect(container.querySelector(".pptx-renderer-wrapper")).not.toBeNull();
    expect(container.querySelector(".pptx-slide-number")?.textContent).toBe("Slide 1");
    expect(container.textContent).toContain("Deck Title");
    expect(document.head.querySelector("#pptx-renderer-styles")).not.toBeNull();
  });

  it("supports the one-shot parse and render API", async () => {
    const buffer = await createTestPptx([{ text: ["One Shot Render"] }]);
    const container = document.createElement("div");

    await renderPptx(buffer, {
      container,
      scale: 0.75,
      showSlideNumbers: false,
      theme: "dark",
    });

    expect(container.querySelector(".pptx-slide")).not.toBeNull();
    expect(container.textContent).toContain("One Shot Render");
  });
});
