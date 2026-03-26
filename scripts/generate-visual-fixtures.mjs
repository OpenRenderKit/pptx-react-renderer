import { mkdir, readFile, writeFile } from "node:fs/promises";
import path from "node:path";
import JSZip from "jszip";
import { PNG } from "pngjs";

const templatePath = path.resolve("test/fixtures/real/basic-preview.pptx");
const outputDir = path.resolve(".tmp", "visual-fixtures");

await mkdir(outputDir, { recursive: true });

await Promise.all([
  writeVisualFixture("text-layout", createTextLayoutSlide()),
  writeVisualFixture("theme-colors", createThemeColorsSlide()),
  writeVisualFixture("group-transform", createGroupTransformSlide()),
  writeVisualFixture("table-layout", createTableLayoutSlide()),
  writeVisualFixture("image-placement", createImagePlacementSlide(), {
    relsXml: createImagePlacementRelationships(),
    media: createImageMediaFiles(),
  }),
]);

console.log(`Generated visual fixtures in ${outputDir}`);

async function writeVisualFixture(name, slideXml, options = {}) {
  const zip = await JSZip.loadAsync(await readFile(templatePath));
  zip.file("ppt/slides/slide1.xml", slideXml);
  zip.file("ppt/slides/slide2.xml", createBlankSlide());
  zip.file("ppt/slides/_rels/slide1.xml.rels", options.relsXml || createEmptyRelationships());
  zip.file("ppt/slides/_rels/slide2.xml.rels", createEmptyRelationships());

  for (const [fileName, buffer] of Object.entries(options.media || {})) {
    zip.file(fileName, buffer);
  }

  await writeFile(path.join(outputDir, `${name}.pptx`), await zip.generateAsync({ type: "nodebuffer" }));
}

function createTextLayoutSlide() {
  return slideXml(`
    <p:sp>
      <p:spPr>
        <a:xfrm><a:off x="457200" y="342900"/><a:ext cx="8229600" cy="731520"/></a:xfrm>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
      </p:spPr>
      <p:txBody>
        <a:bodyPr anchor="ctr" lIns="38100" rIns="38100"/>
        <a:lstStyle/>
        <a:p>
          <a:pPr algn="ctr"><a:buNone/></a:pPr>
          <a:r><a:rPr sz="2800"><a:latin typeface="Liberation Sans"/></a:rPr><a:t>Text Layout Visual Case</a:t></a:r>
        </a:p>
      </p:txBody>
    </p:sp>
    <p:sp>
      <p:spPr>
        <a:xfrm><a:off x="457200" y="1219200"/><a:ext cx="3657600" cy="2743200"/></a:xfrm>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
      </p:spPr>
      <p:txBody>
        <a:bodyPr wrap="square" lIns="38100" tIns="38100" rIns="38100" bIns="38100"/>
        <a:lstStyle/>
        <a:p>
          <a:pPr marL="342900" indent="-171450"><a:buChar char="•"/><a:buFont typeface="Liberation Sans"/></a:pPr>
          <a:r><a:rPr sz="2000"><a:latin typeface="Liberation Sans"/></a:rPr><a:t>Bulleted body copy with wrap.</a:t></a:r>
          <a:br/>
          <a:r><a:rPr sz="1800"><a:latin typeface="Liberation Sans"/></a:rPr><a:t>Second line for layout drift.</a:t></a:r>
        </a:p>
      </p:txBody>
    </p:sp>
    <p:sp>
      <p:spPr>
        <a:xfrm><a:off x="4572000" y="1219200"/><a:ext cx="3657600" cy="2743200"/></a:xfrm>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
      </p:spPr>
      <p:txBody>
        <a:bodyPr anchor="b" lIns="38100" tIns="38100" rIns="38100" bIns="38100"/>
        <a:lstStyle/>
        <a:p>
          <a:pPr algn="r"><a:buNone/></a:pPr>
          <a:r><a:rPr sz="2200" b="1"><a:latin typeface="Liberation Sans"/></a:rPr><a:t>Bottom Right</a:t></a:r>
        </a:p>
        <a:p>
          <a:pPr algn="r"><a:buNone/></a:pPr>
          <a:r><a:rPr sz="1800"><a:latin typeface="Liberation Sans"/></a:rPr><a:t>Insets, anchor, and alignment.</a:t></a:r>
        </a:p>
      </p:txBody>
    </p:sp>
  `);
}

function createThemeColorsSlide() {
  return slideXml(`
    <p:sp>
      <p:spPr>
        <a:xfrm><a:off x="457200" y="685800"/><a:ext cx="2286000" cy="1371600"/></a:xfrm>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
        <a:solidFill><a:schemeClr val="accent1"/></a:solidFill>
      </p:spPr>
    </p:sp>
    <p:sp>
      <p:spPr>
        <a:xfrm><a:off x="3200400" y="685800"/><a:ext cx="2286000" cy="1371600"/></a:xfrm>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
        <a:solidFill><a:schemeClr val="accent2"><a:lumMod val="85000"/></a:schemeClr></a:solidFill>
      </p:spPr>
    </p:sp>
    <p:sp>
      <p:spPr>
        <a:xfrm><a:off x="5943600" y="685800"/><a:ext cx="2286000" cy="1371600"/></a:xfrm>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
        <a:solidFill><a:schemeClr val="accent3"><a:tint val="20000"/></a:schemeClr></a:solidFill>
      </p:spPr>
    </p:sp>
    <p:sp>
      <p:spPr>
        <a:xfrm><a:off x="457200" y="2514600"/><a:ext cx="7772400" cy="1143000"/></a:xfrm>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
      </p:spPr>
      <p:txBody>
        <a:bodyPr anchor="ctr"/>
        <a:lstStyle/>
        <a:p>
          <a:pPr algn="ctr"><a:buNone/></a:pPr>
          <a:r>
            <a:rPr sz="2400">
              <a:solidFill><a:schemeClr val="tx1"/></a:solidFill>
              <a:latin typeface="Liberation Sans"/>
            </a:rPr>
            <a:t>Theme Colors And Color Transforms</a:t>
          </a:r>
        </a:p>
      </p:txBody>
    </p:sp>
  `);
}

function createGroupTransformSlide() {
  return slideXml(`
    <p:grpSp>
      <p:grpSpPr>
        <a:xfrm>
          <a:off x="1371600" y="1143000"/>
          <a:ext cx="5486400" cy="3200400"/>
          <a:chOff x="457200" y="457200"/>
          <a:chExt cx="2743200" cy="1828800"/>
        </a:xfrm>
      </p:grpSpPr>
      <p:sp>
        <p:spPr>
          <a:xfrm><a:off x="457200" y="457200"/><a:ext cx="1828800" cy="548640"/></a:xfrm>
          <a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom>
          <a:solidFill><a:schemeClr val="accent1"/></a:solidFill>
        </p:spPr>
        <p:txBody>
          <a:bodyPr anchor="ctr"/>
          <a:lstStyle/>
          <a:p>
            <a:pPr algn="ctr"><a:buNone/></a:pPr>
            <a:r>
              <a:rPr sz="2200"><a:solidFill><a:schemeClr val="lt1"/></a:solidFill><a:latin typeface="Liberation Sans"/></a:rPr>
              <a:t>Grouped Title</a:t>
            </a:r>
          </a:p>
        </p:txBody>
      </p:sp>
      <p:sp>
        <p:spPr>
          <a:xfrm><a:off x="685800" y="1143000"/><a:ext cx="914400" cy="685800"/></a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
          <a:solidFill><a:schemeClr val="accent2"/></a:solidFill>
        </p:spPr>
      </p:sp>
      <p:sp>
        <p:spPr>
          <a:xfrm><a:off x="1714500" y="1143000"/><a:ext cx="1143000" cy="685800"/></a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
          <a:solidFill><a:schemeClr val="accent3"/></a:solidFill>
        </p:spPr>
      </p:sp>
    </p:grpSp>
  `);
}

function createTableLayoutSlide() {
  return slideXml(`
    <p:graphicFrame>
      <p:xfrm><a:off x="685800" y="914400"/><a:ext cx="7772400" cy="3657600"/></p:xfrm>
      <a:graphic>
        <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">
          <a:tbl>
            <a:tblPr firstRow="1">
              <a:solidFill><a:schemeClr val="bg1"><a:lumMod val="97000"/></a:schemeClr></a:solidFill>
            </a:tblPr>
            <a:tblGrid>
              <a:gridCol w="2286000"/>
              <a:gridCol w="2743200"/>
              <a:gridCol w="2743200"/>
            </a:tblGrid>
            <a:tr h="548640">
              ${tableCell("Capability", { header: true })}
              ${tableCell("Current", { header: true })}
              ${tableCell("Target", { header: true })}
            </a:tr>
            <a:tr h="822960">
              ${tableCell("Visual Regression", { align: "l" })}
              ${tableCell("One complex deck", { align: "l" })}
              ${tableCell("Case-based coverage", { align: "l", fill: "accent3" })}
            </a:tr>
            <a:tr h="822960">
              ${tableCell("Artifacts", { align: "l" })}
              ${tableCell("Smoke only", { align: "l" })}
              ${tableCell("Reference, actual, diff", { align: "l", fill: "accent1", text: "lt1" })}
            </a:tr>
          </a:tbl>
        </a:graphicData>
      </a:graphic>
    </p:graphicFrame>
  `);
}

function createImagePlacementSlide() {
  return slideXml(`
    <p:sp>
      <p:spPr>
        <a:xfrm><a:off x="457200" y="342900"/><a:ext cx="8229600" cy="731520"/></a:xfrm>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
      </p:spPr>
      <p:txBody>
        <a:bodyPr anchor="ctr"/>
        <a:lstStyle/>
        <a:p>
          <a:pPr algn="ctr"><a:buNone/></a:pPr>
          <a:r><a:rPr sz="2400"><a:latin typeface="Liberation Sans"/></a:rPr><a:t>Image Placement Visual Case</a:t></a:r>
        </a:p>
      </p:txBody>
    </p:sp>
    <p:pic>
      <p:blipFill><a:blip r:embed="rId1"/></p:blipFill>
      <p:spPr>
        <a:xfrm><a:off x="1371600" y="1371600"/><a:ext cx="1828800" cy="1828800"/></a:xfrm>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
      </p:spPr>
    </p:pic>
    <p:pic>
      <p:blipFill><a:blip r:embed="rId2"/></p:blipFill>
      <p:spPr>
        <a:xfrm><a:off x="5029200" y="1600200"/><a:ext cx="1371600" cy="2286000"/></a:xfrm>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
      </p:spPr>
    </p:pic>
  `);
}

function createImagePlacementRelationships() {
  return `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image2.png"/>
</Relationships>`;
}

function createImageMediaFiles() {
  return {
    "ppt/media/image1.png": createSolidPng("#2563eb"),
    "ppt/media/image2.png": createSolidPng("#f97316"),
  };
}

function createSolidPng(hexColor) {
  const png = new PNG({ width: 64, height: 64 });
  const [r, g, b] = hexToRgb(hexColor);

  for (let index = 0; index < png.data.length; index += 4) {
    png.data[index] = r;
    png.data[index + 1] = g;
    png.data[index + 2] = b;
    png.data[index + 3] = 255;
  }

  return PNG.sync.write(png);
}

function tableCell(text, options = {}) {
  const fill = options.fill || (options.header ? "accent1" : "bg1");
  const textColor = options.text || (options.header ? "lt1" : "tx1");
  const align = options.align || "ctr";

  return `
    <a:tc>
      <a:txBody>
        <a:bodyPr/>
        <a:lstStyle/>
        <a:p>
          <a:pPr algn="${align}"><a:buNone/></a:pPr>
          <a:r>
            <a:rPr sz="${options.header ? "1800" : "1600"}" ${options.header ? 'b="1"' : ""}>
              <a:solidFill><a:schemeClr val="${textColor}"/></a:solidFill>
              <a:latin typeface="Liberation Sans"/>
            </a:rPr>
            <a:t>${escapeXml(text)}</a:t>
          </a:r>
        </a:p>
      </a:txBody>
      <a:tcPr marL="38100" marR="38100" marT="38100" marB="38100" anchor="ctr">
        <a:lnL w="12700"><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></a:lnL>
        <a:lnR w="12700"><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></a:lnR>
        <a:lnT w="12700"><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></a:lnT>
        <a:lnB w="12700"><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></a:lnB>
        <a:solidFill><a:schemeClr val="${fill}"/></a:solidFill>
      </a:tcPr>
    </a:tc>
  `;
}

function createBlankSlide() {
  return slideXml("");
}

function slideXml(spTreeChildren) {
  return `<?xml version="1.0" encoding="UTF-8"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:cSld>
    <p:spTree>
      ${spTreeChildren}
    </p:spTree>
  </p:cSld>
</p:sld>`;
}

function createEmptyRelationships() {
  return `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`;
}

function hexToRgb(hex) {
  const clean = hex.replace("#", "");
  return [
    Number.parseInt(clean.slice(0, 2), 16),
    Number.parseInt(clean.slice(2, 4), 16),
    Number.parseInt(clean.slice(4, 6), 16),
  ];
}

function escapeXml(value) {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}
