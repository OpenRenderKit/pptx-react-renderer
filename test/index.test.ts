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

  it("parses paragraph bullets, fields, and line breaks from rich text", async () => {
    const buffer = await createTestPptx([
      {
        rawXml: `<?xml version="1.0" encoding="UTF-8"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:sp>
        <p:spPr>
          <a:xfrm>
            <a:off x="457200" y="457200"/>
            <a:ext cx="7315200" cy="1828800"/>
          </a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
        </p:spPr>
        <p:txBody>
          <a:bodyPr wrap="square" anchor="ctr" lIns="12700" rIns="25400"/>
          <a:lstStyle/>
          <a:p>
            <a:pPr algn="ctr" marL="457200" indent="-228600">
              <a:buChar char="•"/>
              <a:buFont typeface="Wingdings"/>
            </a:pPr>
            <a:r>
              <a:rPr sz="2400" b="1"/>
              <a:t>Bullet item</a:t>
            </a:r>
            <a:br/>
            <a:fld type="slidenum">
              <a:rPr sz="2000"/>
              <a:t>7</a:t>
            </a:fld>
          </a:p>
          <a:p>
            <a:r>
              <a:t>Second paragraph</a:t>
            </a:r>
          </a:p>
        </p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
</p:sld>`,
      },
    ]);

    const { slides } = await parsePptx(buffer);
    const textElement = slides[0].elements[0];

    expect(textElement).toMatchObject({
      type: "text",
      align: "center",
      verticalAlign: "middle",
      leftInset: 1,
      rightInset: 2,
    });

    if (textElement.type !== "text" || !textElement.paragraphs) {
      throw new Error("Expected a parsed text element with paragraphs");
    }

    expect(textElement.paragraphs[0]).toMatchObject({
      bullet: true,
      bulletChar: "•",
      bulletFont: "Wingdings",
      align: "center",
      leftMargin: 457200,
      indent: -228600,
    });
    expect(textElement.paragraphs[0].runs).toEqual([
      expect.objectContaining({ type: "run", text: "Bullet item", bold: true }),
      { type: "break" },
      expect.objectContaining({ type: "field", text: "7", fieldType: "slidenum" }),
    ]);
  });

  it("falls back to a placeholder for unresolved images", async () => {
    const buffer = await createTestPptx([
      {
        extraSpTreeChildren: [
          `<p:pic>
            <p:blipFill>
              <a:blip r:embed="rId1" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>
            </p:blipFill>
            <p:spPr>
              <a:xfrm>
                <a:off x="914400" y="914400"/>
                <a:ext cx="1828800" cy="1828800"/>
              </a:xfrm>
              <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
            </p:spPr>
          </p:pic>`,
        ],
      },
    ]);

    const { slides } = await parsePptx(buffer);
    const container = document.createElement("div");

    renderSlides(slides, { container });

    expect(slides[0].elements).toEqual(
      expect.arrayContaining([
        expect.objectContaining({ type: "image", src: "", mimeType: "image/png" }),
      ]),
    );
    expect(container.textContent).toContain("📷");
  });

  it("renders chart graphic frames as explicit placeholders", async () => {
    const buffer = await createTestPptx([
      {
        extraSpTreeChildren: [
          `<p:graphicFrame>
            <p:xfrm>
              <a:off x="914400" y="914400"/>
              <a:ext cx="2743200" cy="1828800"/>
            </p:xfrm>
            <a:graphic xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
              <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
                <c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rIdChart1"/>
              </a:graphicData>
            </a:graphic>
          </p:graphicFrame>`,
        ],
      },
    ]);

    const { slides } = await parsePptx(buffer);
    const container = document.createElement("div");

    renderSlides(slides, { container });

    expect(slides[0].elements).toEqual(
      expect.arrayContaining([expect.objectContaining({ type: "graphicFrame", graphicType: "chart" })]),
    );
    expect(container.textContent).toContain("Chart");
  });

  it("throws a clear error when browser parsing APIs are unavailable", async () => {
    const buffer = await createTestPptx([{ text: ["Guard"] }]);
    const originalDomParser = globalThis.DOMParser;
    const originalFileReader = globalThis.FileReader;
    const originalBlob = globalThis.Blob;

    Object.defineProperty(globalThis, "DOMParser", { value: undefined, configurable: true });
    Object.defineProperty(globalThis, "FileReader", { value: undefined, configurable: true });
    Object.defineProperty(globalThis, "Blob", { value: undefined, configurable: true });

    try {
      await expect(parsePptx(buffer)).rejects.toThrow(
        "pptx-react-renderer requires browser parsing APIs",
      );
    } finally {
      Object.defineProperty(globalThis, "DOMParser", {
        value: originalDomParser,
        configurable: true,
      });
      Object.defineProperty(globalThis, "FileReader", {
        value: originalFileReader,
        configurable: true,
      });
      Object.defineProperty(globalThis, "Blob", { value: originalBlob, configurable: true });
    }
  });

  it("throws a clear error when document is unavailable for rendering", () => {
    const originalDocument = globalThis.document;

    Object.defineProperty(globalThis, "document", { value: undefined, configurable: true });

    try {
      expect(() =>
        renderSlides(
          [
            {
              number: 1,
              width: 960,
              height: 720,
              elements: [],
            },
          ],
          {
            container: {} as HTMLElement,
          },
        ),
      ).toThrow("pptx-react-renderer requires a browser document to render slides.");
    } finally {
      Object.defineProperty(globalThis, "document", {
        value: originalDocument,
        configurable: true,
      });
    }
  });
});
