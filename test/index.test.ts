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

  it("applies OOXML default text insets when bodyPr omits inset attributes", async () => {
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
            <a:ext cx="3657600" cy="1371600"/>
          </a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
        </p:spPr>
        <p:txBody>
          <a:bodyPr anchor="t"/>
          <a:lstStyle/>
          <a:p><a:r><a:t>Default inset text</a:t></a:r></a:p>
        </p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
</p:sld>`,
      },
    ]);

    const { slides } = await parsePptx(buffer);
    const textElement = slides[0].elements[0];

    if (textElement.type !== "text") {
      throw new Error("Expected a parsed text element");
    }

    expect(textElement.leftInset).toBeCloseTo(7.2, 4);
    expect(textElement.rightInset).toBeCloseTo(7.2, 4);
    expect(textElement.topInset).toBeCloseTo(3.6, 4);
    expect(textElement.bottomInset).toBeCloseTo(3.6, 4);
  });

  it("parses empty txBody shape placeholders as visual shapes", async () => {
    const buffer = await createTestPptx([
      {
        rawXml: `<?xml version="1.0" encoding="UTF-8"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:sp>
        <p:spPr>
          <a:xfrm><a:off x="0" y="0"/><a:ext cx="9144000" cy="6858000"/></a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
          <a:solidFill><a:schemeClr val="tx1"/></a:solidFill>
        </p:spPr>
        <p:txBody>
          <a:bodyPr anchor="ctr"/>
          <a:lstStyle/>
          <a:p><a:pPr algn="ctr"/></a:p>
        </p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
</p:sld>`,
      },
    ]);

    const { slides } = await parsePptx(buffer);
    expect(slides[0].elements[0]).toMatchObject({
      type: "shape",
      shapeType: "rect",
      fillColor: "#000000",
    });
  });

  it("keeps empty txBody placeholders as text when no explicit visual style exists", async () => {
    const buffer = await createTestPptx([
      {
        rawXml: `<?xml version="1.0" encoding="UTF-8"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:sp>
        <p:spPr>
          <a:xfrm><a:off x="0" y="0"/><a:ext cx="1828800" cy="914400"/></a:xfrm>
        </p:spPr>
        <p:txBody>
          <a:bodyPr anchor="ctr"/>
          <a:lstStyle/>
          <a:p><a:pPr algn="ctr"/></a:p>
        </p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
</p:sld>`,
      },
    ]);

    const { slides } = await parsePptx(buffer);
    expect(slides[0].elements[0]).toMatchObject({ type: "text" });
  });

  it("parses renderable text shapes with visual frame styling", async () => {
    const buffer = await createTestPptx([
      {
        rawXml: `<?xml version="1.0" encoding="UTF-8"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:sp>
        <p:spPr>
          <a:xfrm><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="914400"/></a:xfrm>
          <a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom>
          <a:solidFill><a:srgbClr val="4472C4"/></a:solidFill>
          <a:ln w="25400">
            <a:solidFill><a:srgbClr val="112233"/></a:solidFill>
            <a:prstDash val="dash"/>
          </a:ln>
        </p:spPr>
        <p:txBody>
          <a:bodyPr anchor="ctr" lIns="38100" rIns="38100"/>
          <a:lstStyle/>
          <a:p>
            <a:pPr algn="ctr"><a:buNone/></a:pPr>
            <a:r><a:rPr sz="1800"/><a:t>Styled text box</a:t></a:r>
          </a:p>
        </p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
</p:sld>`,
      },
    ]);

    const { slides } = await parsePptx(buffer);
    expect(slides[0].elements[0]).toMatchObject({
      type: "text",
      x: 96,
      y: 96,
      width: 192,
      height: 96,
      frame: {
        shapeType: "roundRect",
        fillColor: "#4472c4",
        strokeColor: "#112233",
        strokeWidth: 2,
        strokeDash: "dashed",
        roundedCorners: 0.15,
      },
    });
  });

  it("renders text shape frame styling without dropping text content", async () => {
    const buffer = await createTestPptx([
      {
        rawXml: `<?xml version="1.0" encoding="UTF-8"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:sp>
        <p:spPr>
          <a:xfrm><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="914400"/></a:xfrm>
          <a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom>
          <a:solidFill><a:srgbClr val="4472C4"/></a:solidFill>
          <a:ln w="25400">
            <a:solidFill><a:srgbClr val="112233"/></a:solidFill>
            <a:prstDash val="dash"/>
          </a:ln>
        </p:spPr>
        <p:txBody>
          <a:bodyPr anchor="ctr" lIns="38100" rIns="38100"/>
          <a:lstStyle/>
          <a:p>
            <a:pPr algn="ctr"><a:buNone/></a:pPr>
            <a:r><a:rPr sz="1800"/><a:t>Styled text box</a:t></a:r>
          </a:p>
        </p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
</p:sld>`,
      },
    ]);

    const { slides } = await parsePptx(buffer);
    const container = document.createElement("div");

    renderSlides(slides, { container });

    const textElement = container.querySelector(".pptx-text-element");
    if (!(textElement instanceof HTMLElement)) {
      throw new Error("Expected a rendered text element");
    }

    expect(textElement.textContent).toContain("Styled text box");
    expect(textElement.style.backgroundColor).toBe("rgb(68, 114, 196)");
    expect(textElement.style.borderWidth).toBe("2px");
    expect(textElement.style.borderStyle).toBe("dashed");
    expect(textElement.style.borderColor).toBe("rgb(17, 34, 51)");
    expect(parseFloat(textElement.style.borderRadius)).toBeCloseTo(14.4, 4);
  });

  it("falls back to theme style refs for text shape frame and font defaults", async () => {
    const themeXml = `<?xml version="1.0" encoding="UTF-8"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <a:themeElements>
    <a:clrScheme name="Office">
      <a:dk1><a:srgbClr val="000000"/></a:dk1>
      <a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
      <a:accent1><a:srgbClr val="4472C4"/></a:accent1>
      <a:accent2><a:srgbClr val="ED7D31"/></a:accent2>
    </a:clrScheme>
    <a:fontScheme name="Office">
      <a:majorFont><a:latin typeface="Aptos Display"/></a:majorFont>
      <a:minorFont><a:latin typeface="Aptos"/></a:minorFont>
    </a:fontScheme>
    <a:fmtScheme name="Office">
      <a:fillStyleLst>
        <a:solidFill><a:srgbClr val="111111"/></a:solidFill>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
      </a:fillStyleLst>
      <a:lnStyleLst>
        <a:ln w="12700">
          <a:solidFill><a:srgbClr val="111111"/></a:solidFill>
          <a:prstDash val="dot"/>
        </a:ln>
        <a:ln w="25400">
          <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
          <a:prstDash val="dash"/>
        </a:ln>
      </a:lnStyleLst>
      <a:bgFillStyleLst>
        <a:solidFill><a:schemeClr val="lt1"/></a:solidFill>
      </a:bgFillStyleLst>
    </a:fmtScheme>
  </a:themeElements>
</a:theme>`;

    const buffer = await createTestPptx(
      [
        {
          rawXml: `<?xml version="1.0" encoding="UTF-8"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:sp>
        <p:spPr>
          <a:xfrm><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="914400"/></a:xfrm>
          <a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom>
        </p:spPr>
        <p:style>
          <a:lnRef idx="2"><a:schemeClr val="accent2"/></a:lnRef>
          <a:fillRef idx="2"><a:schemeClr val="accent1"/></a:fillRef>
          <a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef>
          <a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef>
        </p:style>
        <p:txBody>
          <a:bodyPr anchor="ctr"/>
          <a:lstStyle/>
          <a:p>
            <a:pPr algn="ctr"><a:buNone/></a:pPr>
            <a:r><a:t>Theme styled box</a:t></a:r>
          </a:p>
        </p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
</p:sld>`,
        },
      ],
      { themeXml },
    );

    const { slides } = await parsePptx(buffer);
    expect(slides[0].elements[0]).toMatchObject({
      type: "text",
      frame: {
        shapeType: "roundRect",
        fillColor: "#4472c4",
        strokeColor: "#ed7d31",
        strokeWidth: 2,
        strokeDash: "dashed",
        roundedCorners: 0.15,
      },
      defaultFontFamily: "Aptos",
      defaultColor: "#ffffff",
    });
  });

  it("does not apply theme style fills when a shape uses background fill", async () => {
    const themeXml = `<?xml version="1.0" encoding="UTF-8"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <a:themeElements>
    <a:clrScheme name="Office">
      <a:accent1><a:srgbClr val="156082"/></a:accent1>
    </a:clrScheme>
    <a:fontScheme name="Office">
      <a:majorFont><a:latin typeface="Aptos Display"/></a:majorFont>
      <a:minorFont><a:latin typeface="Aptos"/></a:minorFont>
    </a:fontScheme>
    <a:fmtScheme name="Office">
      <a:fillStyleLst>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
      </a:fillStyleLst>
      <a:lnStyleLst>
        <a:ln w="19050">
          <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        </a:ln>
      </a:lnStyleLst>
      <a:bgFillStyleLst>
        <a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>
      </a:bgFillStyleLst>
    </a:fmtScheme>
  </a:themeElements>
</a:theme>`;

    const buffer = await createTestPptx(
      [
        {
          rawXml: `<?xml version="1.0" encoding="UTF-8"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:sp useBgFill="1">
        <p:spPr>
          <a:xfrm><a:off x="0" y="0"/><a:ext cx="9144000" cy="6858000"/></a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
          <a:ln><a:noFill/></a:ln>
        </p:spPr>
        <p:style>
          <a:lnRef idx="1"><a:schemeClr val="accent1"/></a:lnRef>
          <a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef>
          <a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef>
          <a:fontRef idx="minor"><a:schemeClr val="accent1"/></a:fontRef>
        </p:style>
        <p:txBody>
          <a:bodyPr anchor="ctr"/>
          <a:lstStyle/>
          <a:p/>
        </p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
</p:sld>`,
        },
      ],
      { themeXml },
    );

    const { slides } = await parsePptx(buffer);
    expect(slides[0].elements[0]).toMatchObject({
      type: "shape",
      fillColor: undefined,
      fillType: undefined,
      strokeColor: undefined,
      strokeWidth: undefined,
      strokeDash: undefined,
    });
  });

  it("renders theme style ref defaults for text shapes", async () => {
    const themeXml = `<?xml version="1.0" encoding="UTF-8"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <a:themeElements>
    <a:clrScheme name="Office">
      <a:dk1><a:srgbClr val="000000"/></a:dk1>
      <a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
      <a:accent1><a:srgbClr val="4472C4"/></a:accent1>
      <a:accent2><a:srgbClr val="ED7D31"/></a:accent2>
    </a:clrScheme>
    <a:fontScheme name="Office">
      <a:majorFont><a:latin typeface="Aptos Display"/></a:majorFont>
      <a:minorFont><a:latin typeface="Aptos"/></a:minorFont>
    </a:fontScheme>
    <a:fmtScheme name="Office">
      <a:fillStyleLst>
        <a:solidFill><a:srgbClr val="111111"/></a:solidFill>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
      </a:fillStyleLst>
      <a:lnStyleLst>
        <a:ln w="12700">
          <a:solidFill><a:srgbClr val="111111"/></a:solidFill>
          <a:prstDash val="dot"/>
        </a:ln>
        <a:ln w="25400">
          <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
          <a:prstDash val="dash"/>
        </a:ln>
      </a:lnStyleLst>
      <a:bgFillStyleLst>
        <a:solidFill><a:schemeClr val="lt1"/></a:solidFill>
      </a:bgFillStyleLst>
    </a:fmtScheme>
  </a:themeElements>
</a:theme>`;

    const buffer = await createTestPptx(
      [
        {
          rawXml: `<?xml version="1.0" encoding="UTF-8"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:sp>
        <p:spPr>
          <a:xfrm><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="914400"/></a:xfrm>
          <a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom>
        </p:spPr>
        <p:style>
          <a:lnRef idx="2"><a:schemeClr val="accent2"/></a:lnRef>
          <a:fillRef idx="2"><a:schemeClr val="accent1"/></a:fillRef>
          <a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef>
          <a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef>
        </p:style>
        <p:txBody>
          <a:bodyPr anchor="ctr"/>
          <a:lstStyle/>
          <a:p>
            <a:pPr algn="ctr"><a:buNone/></a:pPr>
            <a:r><a:t>Theme styled box</a:t></a:r>
          </a:p>
        </p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
</p:sld>`,
        },
      ],
      { themeXml },
    );

    const { slides } = await parsePptx(buffer);
    const container = document.createElement("div");

    renderSlides(slides, { container });

    const textElement = container.querySelector(".pptx-text-element");
    const textSpan = container.querySelector(".pptx-text-element span");
    if (!(textElement instanceof HTMLElement) || !(textSpan instanceof HTMLElement)) {
      throw new Error("Expected rendered theme-styled text shape");
    }

    expect(textElement.style.backgroundColor).toBe("rgb(68, 114, 196)");
    expect(textElement.style.borderColor).toBe("rgb(237, 125, 49)");
    expect(textElement.style.borderStyle).toBe("dashed");
    expect(textSpan.style.fontFamily).toBe("Aptos");
    expect(textSpan.style.color).toBe("rgb(255, 255, 255)");
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

  it("parses graphicFrame tables into styled table elements", async () => {
    const buffer = await createTestPptx([
      {
        extraSpTreeChildren: [
          `<p:graphicFrame>
            <p:xfrm>
              <a:off x="914400" y="914400"/>
              <a:ext cx="3657600" cy="1828800"/>
            </p:xfrm>
            <a:graphic>
              <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">
                <a:tbl>
                  <a:tblPr firstRow="1">
                    <a:solidFill><a:schemeClr val="lt1"><a:lumMod val="95000"/></a:schemeClr></a:solidFill>
                  </a:tblPr>
                  <a:tblGrid>
                    <a:gridCol w="1828800"/>
                    <a:gridCol w="1828800"/>
                  </a:tblGrid>
                  <a:tr h="457200">
                    <a:tc>
                      <a:txBody>
                        <a:bodyPr/>
                        <a:lstStyle/>
                        <a:p>
                          <a:pPr algn="ctr"><a:buNone/></a:pPr>
                          <a:r>
                            <a:rPr sz="1800" b="1"><a:solidFill><a:schemeClr val="lt1"/></a:solidFill></a:rPr>
                            <a:t>Heading</a:t>
                          </a:r>
                        </a:p>
                      </a:txBody>
                      <a:tcPr anchor="ctr" marL="12700" marR="12700" marT="12700" marB="12700">
                        <a:solidFill><a:schemeClr val="accent1"/></a:solidFill>
                        <a:lnL w="12700"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:lnL>
                        <a:lnR w="12700"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:lnR>
                        <a:lnT w="12700"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:lnT>
                        <a:lnB w="12700"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:lnB>
                      </a:tcPr>
                    </a:tc>
                    <a:tc>
                      <a:txBody>
                        <a:bodyPr/>
                        <a:lstStyle/>
                        <a:p>
                          <a:r>
                            <a:rPr sz="1600"><a:solidFill><a:schemeClr val="dk1"/></a:solidFill></a:rPr>
                            <a:t>Body</a:t>
                          </a:r>
                        </a:p>
                      </a:txBody>
                      <a:tcPr anchor="b" gridSpan="1">
                        <a:solidFill><a:schemeClr val="lt1"><a:lumMod val="95000"/></a:schemeClr></a:solidFill>
                      </a:tcPr>
                    </a:tc>
                  </a:tr>
                </a:tbl>
              </a:graphicData>
            </a:graphic>
          </p:graphicFrame>`,
        ],
      },
    ]);

    const { slides } = await parsePptx(buffer);
    const table = slides[0].elements[0];

    expect(table).toMatchObject({
      type: "table",
      x: 96,
      y: 96,
      width: 384,
      height: 192,
      columnWidths: [192, 192],
      rows: [
        {
          height: 48,
          cells: [
            expect.objectContaining({
              content: "Heading",
              backgroundColor: "#4472c4",
              verticalAlign: "middle",
              defaultColor: "#ffffff",
            }),
            expect.objectContaining({
              content: "Body",
              verticalAlign: "bottom",
              backgroundColor: "#f2f2f2",
            }),
          ],
        },
      ],
    });

    const container = document.createElement("div");
    renderSlides(slides, { container });

    expect(container.querySelector(".pptx-table-element")).not.toBeNull();
    expect(container.textContent).toContain("Heading");
    expect(container.textContent).toContain("Body");
  });

  it("parses grouped children using group transform coordinates and renders them", async () => {
    const buffer = await createTestPptx([
      {
        rawXml: `<?xml version="1.0" encoding="UTF-8"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:grpSp>
        <p:grpSpPr>
          <a:xfrm>
            <a:off x="914400" y="914400"/>
            <a:ext cx="1828800" cy="1828800"/>
            <a:chOff x="457200" y="457200"/>
            <a:chExt cx="914400" cy="914400"/>
          </a:xfrm>
        </p:grpSpPr>
        <p:sp>
          <p:spPr>
            <a:xfrm>
              <a:off x="457200" y="457200"/>
              <a:ext cx="457200" cy="228600"/>
            </a:xfrm>
            <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
          </p:spPr>
          <p:txBody>
            <a:bodyPr wrap="square"/>
            <a:lstStyle/>
            <a:p>
              <a:r>
                <a:t>Grouped Title</a:t>
              </a:r>
            </a:p>
          </p:txBody>
        </p:sp>
        <p:sp>
          <p:spPr>
            <a:xfrm>
              <a:off x="685800" y="685800"/>
              <a:ext cx="228600" cy="228600"/>
            </a:xfrm>
            <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
            <a:solidFill><a:srgbClr val="4472C4"/></a:solidFill>
          </p:spPr>
        </p:sp>
      </p:grpSp>
    </p:spTree>
  </p:cSld>
</p:sld>`,
      },
    ]);

    const { slides } = await parsePptx(buffer);
    const group = slides[0].elements[0];

    expect(group).toMatchObject({
      type: "group",
      x: 96,
      y: 96,
      width: 192,
      height: 192,
    });

    if (group.type !== "group") {
      throw new Error("Expected a parsed group element");
    }

    expect(group.children).toEqual([
      expect.objectContaining({
        type: "text",
        x: 96,
        y: 96,
        width: 96,
        height: 48,
      }),
      expect.objectContaining({
        type: "shape",
        x: 144,
        y: 144,
        width: 48,
        height: 48,
      }),
    ]);

    const container = document.createElement("div");
    renderSlides(slides, { container });

    expect(container.textContent).toContain("Grouped Title");
    expect(container.querySelector(".pptx-group-element .pptx-text-element")).not.toBeNull();
  });

  it("applies scheme color luminance transforms using HSL lightness", async () => {
    const baseAccent = "#4472c4";
    const expectedLumMod = applyExpectedLightnessTransform(baseAccent, { lumMod: 70000 });
    const expectedLumOff = applyExpectedLightnessTransform(baseAccent, { lumOff: 20000 });
    const expectedTint = applyExpectedLightnessTransform(baseAccent, { tint: 25000 });
    const expectedShade = applyExpectedLightnessTransform(baseAccent, { shade: 60000 });

    const buffer = await createTestPptx([
      {
        rawXml: `<?xml version="1.0" encoding="UTF-8"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:sp>
        <p:spPr>
          <a:xfrm><a:off x="457200" y="457200"/><a:ext cx="914400" cy="914400"/></a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
          <a:solidFill><a:schemeClr val="accent1"><a:lumMod val="70000"/></a:schemeClr></a:solidFill>
        </p:spPr>
      </p:sp>
      <p:sp>
        <p:spPr>
          <a:xfrm><a:off x="1600200" y="457200"/><a:ext cx="914400" cy="914400"/></a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
          <a:solidFill><a:schemeClr val="accent1"><a:lumOff val="20000"/></a:schemeClr></a:solidFill>
        </p:spPr>
      </p:sp>
      <p:sp>
        <p:spPr>
          <a:xfrm><a:off x="2743200" y="457200"/><a:ext cx="914400" cy="914400"/></a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
          <a:solidFill><a:schemeClr val="accent1"><a:tint val="25000"/></a:schemeClr></a:solidFill>
        </p:spPr>
      </p:sp>
      <p:sp>
        <p:spPr>
          <a:xfrm><a:off x="3886200" y="457200"/><a:ext cx="914400" cy="914400"/></a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
          <a:solidFill><a:schemeClr val="accent1"><a:shade val="60000"/></a:schemeClr></a:solidFill>
        </p:spPr>
      </p:sp>
    </p:spTree>
  </p:cSld>
</p:sld>`,
      },
    ]);

    const { slides } = await parsePptx(buffer);
    const shapes = slides[0].elements.filter((element) => element.type === "shape");

    expect(shapes).toHaveLength(4);
    expect(shapes[0]).toMatchObject({ fillColor: expectedLumMod });
    expect(shapes[1]).toMatchObject({ fillColor: expectedLumOff });
    expect(shapes[2]).toMatchObject({ fillColor: expectedTint });
    expect(shapes[3]).toMatchObject({ fillColor: expectedShade });
  });

  it("parses custom geometry shapes with path data and alpha-aware fills", async () => {
    const buffer = await createTestPptx([
      {
        rawXml: `<?xml version="1.0" encoding="UTF-8"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:sp>
        <p:spPr>
          <a:xfrm><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="914400"/></a:xfrm>
          <a:custGeom>
            <a:avLst/>
            <a:pathLst>
              <a:path w="1000" h="500">
                <a:moveTo><a:pt x="0" y="0"/></a:moveTo>
                <a:lnTo><a:pt x="1000" y="0"/></a:lnTo>
                <a:lnTo><a:pt x="1000" y="500"/></a:lnTo>
                <a:lnTo><a:pt x="0" y="500"/></a:lnTo>
                <a:close/>
              </a:path>
            </a:pathLst>
          </a:custGeom>
          <a:solidFill><a:schemeClr val="accent1"><a:alpha val="30000"/></a:schemeClr></a:solidFill>
          <a:ln w="12700"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:ln>
        </p:spPr>
      </p:sp>
    </p:spTree>
  </p:cSld>
</p:sld>`,
      },
    ]);

    const { slides } = await parsePptx(buffer);
    const customShape = slides[0].elements[0];

    expect(customShape).toMatchObject({
      type: "shape",
      shapeType: "custom",
      fillColor: "rgba(68, 114, 196, 0.3)",
      strokeColor: "#ffffff",
      strokeWidth: 1,
      viewBox: { x: 0, y: 0, width: 1000, height: 500 },
    });

    if (customShape.type !== "shape") {
      throw new Error("Expected a parsed shape element");
    }

    expect(customShape.pathData).toHaveLength(1);
    expect(customShape.pathData?.[0].commands).toEqual([
      { type: "M", x: 0, y: 0 },
      { type: "L", x: 1000, y: 0 },
      { type: "L", x: 1000, y: 500 },
      { type: "L", x: 0, y: 500 },
      { type: "Z" },
    ]);

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

function applyExpectedLightnessTransform(
  baseColor: string,
  transform: {
    lumMod?: number;
    lumOff?: number;
    tint?: number;
    shade?: number;
  },
): string {
  const rgb = hexToRgb(baseColor);
  if (!rgb) {
    throw new Error(`Invalid expected base color: ${baseColor}`);
  }

  const hsl = rgbToHsl(rgb[0], rgb[1], rgb[2]);
  if (!hsl) {
    throw new Error(`Failed to convert expected color to HSL: ${baseColor}`);
  }

  const [h, s, initialL] = hsl;
  let l = initialL;

  if (transform.lumMod !== undefined || transform.lumOff !== undefined) {
    const mod = (transform.lumMod ?? 100000) / 100000;
    const off = (transform.lumOff ?? 0) / 100000;
    l = clampUnit(l * mod + off);
  }

  if (transform.tint !== undefined) {
    const tintFactor = transform.tint / 100000;
    l = clampUnit(l + (1 - l) * tintFactor);
  }

  if (transform.shade !== undefined) {
    const shadeFactor = transform.shade / 100000;
    l = clampUnit(l * shadeFactor);
  }

  const [r, g, b] = hslToRgb(h, s, l);
  return rgbToHex(r, g, b);
}

function hexToRgb(color: string): [number, number, number] | undefined {
  const hex = color.replace("#", "");
  if (hex.length !== 6) return undefined;
  return [
    parseInt(hex.slice(0, 2), 16),
    parseInt(hex.slice(2, 4), 16),
    parseInt(hex.slice(4, 6), 16),
  ];
}

function rgbToHex(r: number, g: number, b: number): string {
  return `#${[r, g, b]
    .map((channel) => clampChannel(channel).toString(16).padStart(2, "0"))
    .join("")}`;
}

function rgbToHsl(r: number, g: number, b: number): [number, number, number] | undefined {
  const rn = r / 255;
  const gn = g / 255;
  const bn = b / 255;
  const max = Math.max(rn, gn, bn);
  const min = Math.min(rn, gn, bn);
  const delta = max - min;
  const l = (max + min) / 2;

  let h = 0;
  let s = 0;

  if (delta !== 0) {
    s = l > 0.5 ? delta / (2 - max - min) : delta / (max + min);

    if (max === rn) {
      h = (gn - bn) / delta + (gn < bn ? 6 : 0);
    } else if (max === gn) {
      h = (bn - rn) / delta + 2;
    } else {
      h = (rn - gn) / delta + 4;
    }

    h /= 6;
  }

  if (!Number.isFinite(h) || !Number.isFinite(s) || !Number.isFinite(l)) {
    return undefined;
  }

  return [h, s, l];
}

function hueToRgb(p: number, q: number, t: number): number {
  let adjusted = t;
  if (adjusted < 0) adjusted += 1;
  if (adjusted > 1) adjusted -= 1;
  if (adjusted < 1 / 6) return p + (q - p) * 6 * adjusted;
  if (adjusted < 1 / 2) return q;
  if (adjusted < 2 / 3) return p + (q - p) * (2 / 3 - adjusted) * 6;
  return p;
}

function hslToRgb(h: number, s: number, l: number): [number, number, number] {
  if (s === 0) {
    const gray = Math.round(l * 255);
    return [gray, gray, gray];
  }

  const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
  const p = 2 * l - q;
  const r = hueToRgb(p, q, h + 1 / 3);
  const g = hueToRgb(p, q, h);
  const b = hueToRgb(p, q, h - 1 / 3);

  return [Math.round(r * 255), Math.round(g * 255), Math.round(b * 255)];
}

function clampUnit(value: number): number {
  return Math.max(0, Math.min(1, value));
}

function clampChannel(value: number): number {
  return Math.max(0, Math.min(255, Math.round(value)));
}
