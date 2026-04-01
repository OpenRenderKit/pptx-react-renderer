import { mkdir, readFile, writeFile } from "node:fs/promises";
import path from "node:path";
import JSZip from "jszip";
import { PNG } from "pngjs";

const templatePath = path.resolve("test/fixtures/real/complex-sanitized.pptx");
const outputDir = path.resolve(".tmp", "visual-fixtures");
const slideCount = 12;

await mkdir(outputDir, { recursive: true });

await Promise.all([
  writeVisualFixture("text-layout", createTextLayoutSlide()),
  writeVisualFixture("bullets-numbering", createBulletsNumberingSlide()),
  writeVisualFixture("text-insets-anchor", createTextInsetsAnchorSlide()),
  writeVisualFixture("paragraph-spacing", createParagraphSpacingSlide()),
  writeVisualFixture("theme-colors", createThemeColorsSlide(), {
    relsXml: createThemeFixtureRelationships(),
  }),
  writeVisualFixture("theme-luminance", createThemeLuminanceSlide(), {
    relsXml: createThemeFixtureRelationships(),
  }),
  writeVisualFixture("theme-style-refs", createThemeStyleRefsSlide(), {
    relsXml: createThemeFixtureRelationships(),
    files: {
      "ppt/theme/theme1.xml": createThemeStyleRefsTheme(),
    },
  }),
  writeVisualFixture("group-transform", createGroupTransformSlide()),
  writeVisualFixture("nested-groups", createNestedGroupsSlide()),
  writeVisualFixture("table-layout", createTableLayoutSlide()),
  writeVisualFixture("image-placement", createImagePlacementSlide(), {
    relsXml: createImagePlacementRelationships(),
    media: createImageMediaFiles(),
  }),
  writeVisualFixture("image-cropping", createImageCroppingSlide(), {
    relsXml: createImageCroppingRelationships(),
    media: createImageCroppingMediaFiles(),
  }),
  writeVisualFixture("unsupported-chart", createUnsupportedChartSlide(), {
    relsXml: createUnsupportedChartRelationships(),
    files: {
      "ppt/charts/chart1.xml": createChartXml(),
    },
  }),
]);

console.log(`Generated visual fixtures in ${outputDir}`);

async function writeVisualFixture(name, slideXml, options = {}) {
  const zip = await JSZip.loadAsync(await readFile(templatePath));
  zip.file("ppt/slides/slide1.xml", slideXml);
  zip.file("ppt/slides/_rels/slide1.xml.rels", options.relsXml || createEmptyRelationships());

  for (let slideNumber = 2; slideNumber <= slideCount; slideNumber += 1) {
    zip.file(`ppt/slides/slide${slideNumber}.xml`, createBlankSlide());
    zip.file(`ppt/slides/_rels/slide${slideNumber}.xml.rels`, createEmptyRelationships());
  }

  for (const [fileName, buffer] of Object.entries(options.media || {})) {
    zip.file(fileName, buffer);
  }
  for (const [fileName, content] of Object.entries(options.files || {})) {
    zip.file(fileName, content);
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

function createBulletsNumberingSlide() {
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
          <a:r><a:rPr sz="2600"><a:latin typeface="Liberation Sans"/></a:rPr><a:t>Bullets And Auto-numbering</a:t></a:r>
        </a:p>
      </p:txBody>
    </p:sp>
    <p:sp>
      <p:spPr>
        <a:xfrm><a:off x="685800" y="1371600"/><a:ext cx="3657600" cy="2743200"/></a:xfrm>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
      </p:spPr>
      <p:txBody>
        <a:bodyPr lIns="38100" rIns="38100"/>
        <a:lstStyle/>
        <a:p>
          <a:pPr marL="342900" indent="-171450"><a:buChar char="•"/><a:buFont typeface="Liberation Sans"/></a:pPr>
          <a:r><a:rPr sz="1900"><a:latin typeface="Liberation Sans"/></a:rPr><a:t>Primary bullet</a:t></a:r>
        </a:p>
        <a:p>
          <a:pPr marL="685800" indent="-171450" lvl="1"><a:buChar char="◦"/><a:buFont typeface="Liberation Sans"/></a:pPr>
          <a:r><a:rPr sz="1800"><a:latin typeface="Liberation Sans"/></a:rPr><a:t>Nested bullet</a:t></a:r>
        </a:p>
      </p:txBody>
    </p:sp>
    <p:sp>
      <p:spPr>
        <a:xfrm><a:off x="4572000" y="1371600"/><a:ext cx="3200400" cy="2743200"/></a:xfrm>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
      </p:spPr>
      <p:txBody>
        <a:bodyPr lIns="38100" rIns="38100"/>
        <a:lstStyle/>
        <a:p>
          <a:pPr marL="342900" indent="-171450"><a:buAutoNum type="arabic" startAt="3"/></a:pPr>
          <a:r><a:rPr sz="1900"><a:latin typeface="Liberation Sans"/></a:rPr><a:t>Third step</a:t></a:r>
        </a:p>
        <a:p>
          <a:pPr marL="342900" indent="-171450"><a:buAutoNum type="arabic" startAt="4"/></a:pPr>
          <a:r><a:rPr sz="1900"><a:latin typeface="Liberation Sans"/></a:rPr><a:t>Fourth step</a:t></a:r>
        </a:p>
      </p:txBody>
    </p:sp>
  `);
}

function createTextInsetsAnchorSlide() {
  return slideXml(`
    <p:sp>
      <p:spPr>
        <a:xfrm><a:off x="685800" y="914400"/><a:ext cx="2286000" cy="2286000"/></a:xfrm>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
        <a:solidFill><a:schemeClr val="accent1"><a:lumOff val="70000"/></a:schemeClr></a:solidFill>
      </p:spPr>
      <p:txBody>
        <a:bodyPr anchor="t" lIns="228600" tIns="228600" rIns="38100" bIns="38100"/>
        <a:lstStyle/>
        <a:p><a:pPr algn="l"><a:buNone/></a:pPr><a:r><a:rPr sz="1800"><a:latin typeface="Liberation Sans"/></a:rPr><a:t>Top anchored</a:t></a:r></a:p>
      </p:txBody>
    </p:sp>
    <p:sp>
      <p:spPr>
        <a:xfrm><a:off x="3429000" y="914400"/><a:ext cx="2286000" cy="2286000"/></a:xfrm>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
        <a:solidFill><a:schemeClr val="accent2"><a:lumOff val="70000"/></a:schemeClr></a:solidFill>
      </p:spPr>
      <p:txBody>
        <a:bodyPr anchor="ctr" lIns="38100" tIns="38100" rIns="228600" bIns="38100"/>
        <a:lstStyle/>
        <a:p><a:pPr algn="ctr"><a:buNone/></a:pPr><a:r><a:rPr sz="1800"><a:latin typeface="Liberation Sans"/></a:rPr><a:t>Middle anchored</a:t></a:r></a:p>
      </p:txBody>
    </p:sp>
    <p:sp>
      <p:spPr>
        <a:xfrm><a:off x="6172200" y="914400"/><a:ext cx="2286000" cy="2286000"/></a:xfrm>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
        <a:solidFill><a:schemeClr val="accent3"><a:lumOff val="70000"/></a:schemeClr></a:solidFill>
      </p:spPr>
      <p:txBody>
        <a:bodyPr anchor="b" lIns="38100" tIns="38100" rIns="38100" bIns="228600"/>
        <a:lstStyle/>
        <a:p><a:pPr algn="r"><a:buNone/></a:pPr><a:r><a:rPr sz="1800"><a:latin typeface="Liberation Sans"/></a:rPr><a:t>Bottom anchored</a:t></a:r></a:p>
      </p:txBody>
    </p:sp>
  `);
}

function createParagraphSpacingSlide() {
  return slideXml(`
    <p:sp>
      <p:spPr>
        <a:xfrm><a:off x="685800" y="685800"/><a:ext cx="7772400" cy="3886200"/></a:xfrm>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
      </p:spPr>
      <p:txBody>
        <a:bodyPr lIns="38100" tIns="38100" rIns="38100" bIns="38100"/>
        <a:lstStyle/>
        <a:p>
          <a:pPr algn="l">
            <a:spcBef><a:spcPts val="1200"/></a:spcBef>
            <a:spcAft><a:spcPts val="800"/></a:spcAft>
            <a:lnSpc><a:spcPct val="140000"/></a:lnSpc>
            <a:buNone/>
          </a:pPr>
          <a:r><a:rPr sz="1800"><a:latin typeface="Liberation Sans"/></a:rPr><a:t>Paragraph one uses increased line spacing and paragraph spacing.</a:t></a:r>
        </a:p>
        <a:p>
          <a:pPr algn="l">
            <a:spcBef><a:spcPts val="600"/></a:spcBef>
            <a:spcAft><a:spcPts val="1800"/></a:spcAft>
            <a:lnSpc><a:spcPct val="90000"/></a:lnSpc>
            <a:buNone/>
          </a:pPr>
          <a:r><a:rPr sz="1800"><a:latin typeface="Liberation Sans"/></a:rPr><a:t>Paragraph two should sit tighter internally but farther from the next block.</a:t></a:r>
        </a:p>
        <a:p>
          <a:pPr algn="l"><a:buNone/></a:pPr>
          <a:r><a:rPr sz="1800"><a:latin typeface="Liberation Sans"/></a:rPr><a:t>Paragraph three is the baseline.</a:t></a:r>
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

function createThemeLuminanceSlide() {
  return slideXml(`
    <p:sp>
      <p:spPr>
        <a:xfrm><a:off x="685800" y="685800"/><a:ext cx="1600200" cy="2286000"/></a:xfrm>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
        <a:solidFill><a:schemeClr val="accent1"/></a:solidFill>
      </p:spPr>
    </p:sp>
    <p:sp>
      <p:spPr>
        <a:xfrm><a:off x="2514600" y="685800"/><a:ext cx="1600200" cy="2286000"/></a:xfrm>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
        <a:solidFill><a:schemeClr val="accent1"><a:lumMod val="70000"/></a:schemeClr></a:solidFill>
      </p:spPr>
    </p:sp>
    <p:sp>
      <p:spPr>
        <a:xfrm><a:off x="4343400" y="685800"/><a:ext cx="1600200" cy="2286000"/></a:xfrm>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
        <a:solidFill><a:schemeClr val="accent1"><a:lumOff val="20000"/></a:schemeClr></a:solidFill>
      </p:spPr>
    </p:sp>
    <p:sp>
      <p:spPr>
        <a:xfrm><a:off x="6172200" y="685800"/><a:ext cx="1600200" cy="2286000"/></a:xfrm>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
        <a:solidFill><a:schemeClr val="accent1"><a:tint val="25000"/></a:schemeClr></a:solidFill>
      </p:spPr>
    </p:sp>
    <p:sp>
      <p:spPr>
        <a:xfrm><a:off x="685800" y="3200400"/><a:ext cx="7772400" cy="1143000"/></a:xfrm>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
      </p:spPr>
      <p:txBody>
        <a:bodyPr anchor="ctr"/>
        <a:lstStyle/>
        <a:p><a:pPr algn="ctr"><a:buNone/></a:pPr><a:r><a:rPr sz="2200"><a:latin typeface="Liberation Sans"/></a:rPr><a:t>Luminance And Tint Transform Coverage</a:t></a:r></a:p>
      </p:txBody>
    </p:sp>
  `);
}

function createThemeStyleRefsSlide() {
  return slideXml(`
    <p:sp>
      <p:spPr>
        <a:xfrm><a:off x="685800" y="685800"/><a:ext cx="3657600" cy="1828800"/></a:xfrm>
        <a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom>
      </p:spPr>
      <p:style>
        <a:lnRef idx="2"><a:schemeClr val="accent2"/></a:lnRef>
        <a:fillRef idx="2"><a:schemeClr val="accent1"/></a:fillRef>
        <a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef>
        <a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef>
      </p:style>
      <p:txBody>
        <a:bodyPr anchor="ctr" lIns="152400" tIns="76200" rIns="152400" bIns="76200"/>
        <a:lstStyle/>
        <a:p>
          <a:pPr algn="ctr"><a:buNone/></a:pPr>
          <a:r><a:rPr sz="2600" b="1"/><a:t>Theme Style Refs</a:t></a:r>
        </a:p>
        <a:p>
          <a:pPr algn="ctr"><a:buNone/></a:pPr>
          <a:r><a:rPr sz="1800"/><a:t>Fill, line, and font defaults come from p:style.</a:t></a:r>
        </a:p>
      </p:txBody>
    </p:sp>
    <p:sp>
      <p:spPr>
        <a:xfrm><a:off x="4790400" y="1828800"/><a:ext cx="2971800" cy="1600200"/></a:xfrm>
        <a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom>
      </p:spPr>
      <p:style>
        <a:lnRef idx="2"><a:schemeClr val="accent1"/></a:lnRef>
        <a:fillRef idx="2"><a:schemeClr val="accent2"/></a:fillRef>
        <a:effectRef idx="0"><a:schemeClr val="accent2"/></a:effectRef>
        <a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef>
      </p:style>
      <p:txBody>
        <a:bodyPr anchor="ctr" lIns="114300" tIns="76200" rIns="114300" bIns="76200"/>
        <a:lstStyle/>
        <a:p>
          <a:pPr algn="ctr"><a:buNone/></a:pPr>
          <a:r><a:rPr sz="2200" b="1"/><a:t>Secondary Box</a:t></a:r>
        </a:p>
      </p:txBody>
    </p:sp>
  `);
}

function createThemeStyleRefsTheme() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Theme Style Refs">
  <a:themeElements>
    <a:clrScheme name="Custom">
      <a:dk1><a:srgbClr val="111111"/></a:dk1>
      <a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
      <a:dk2><a:srgbClr val="1F1F1F"/></a:dk2>
      <a:lt2><a:srgbClr val="F5F5F5"/></a:lt2>
      <a:accent1><a:srgbClr val="4472C4"/></a:accent1>
      <a:accent2><a:srgbClr val="ED7D31"/></a:accent2>
      <a:accent3><a:srgbClr val="70AD47"/></a:accent3>
      <a:accent4><a:srgbClr val="FFC000"/></a:accent4>
      <a:accent5><a:srgbClr val="5B9BD5"/></a:accent5>
      <a:accent6><a:srgbClr val="A5A5A5"/></a:accent6>
      <a:hlink><a:srgbClr val="0563C1"/></a:hlink>
      <a:folHlink><a:srgbClr val="954F72"/></a:folHlink>
    </a:clrScheme>
    <a:fontScheme name="Custom">
      <a:majorFont><a:latin typeface="Liberation Sans"/></a:majorFont>
      <a:minorFont><a:latin typeface="Liberation Sans"/></a:minorFont>
    </a:fontScheme>
    <a:fmtScheme name="Custom">
      <a:fillStyleLst>
        <a:solidFill><a:srgbClr val="DDDDDD"/></a:solidFill>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
      </a:fillStyleLst>
      <a:lnStyleLst>
        <a:ln w="12700">
          <a:solidFill><a:srgbClr val="999999"/></a:solidFill>
        </a:ln>
        <a:ln w="25400">
          <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        </a:ln>
      </a:lnStyleLst>
      <a:bgFillStyleLst>
        <a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>
      </a:bgFillStyleLst>
    </a:fmtScheme>
  </a:themeElements>
</a:theme>`;
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

function createNestedGroupsSlide() {
  return slideXml(`
    <p:grpSp>
      <p:grpSpPr>
        <a:xfrm>
          <a:off x="1143000" y="914400"/>
          <a:ext cx="5943600" cy="3657600"/>
          <a:chOff x="0" y="0"/>
          <a:chExt cx="5943600" cy="3657600"/>
        </a:xfrm>
      </p:grpSpPr>
      <p:sp>
        <p:spPr>
          <a:xfrm><a:off x="0" y="0"/><a:ext cx="5943600" cy="3657600"/></a:xfrm>
          <a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom>
          <a:solidFill><a:schemeClr val="accent1"><a:lumOff val="82000"/></a:schemeClr></a:solidFill>
        </p:spPr>
      </p:sp>
      <p:grpSp>
        <p:grpSpPr>
          <a:xfrm>
            <a:off x="685800" y="685800"/>
            <a:ext cx="4572000" cy="2286000"/>
            <a:chOff x="0" y="0"/>
            <a:chExt cx="4572000" cy="2286000"/>
          </a:xfrm>
        </p:grpSpPr>
        <p:sp>
          <p:spPr>
            <a:xfrm><a:off x="0" y="0"/><a:ext cx="4572000" cy="548640"/></a:xfrm>
            <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
            <a:solidFill><a:schemeClr val="accent2"/></a:solidFill>
          </p:spPr>
          <p:txBody>
            <a:bodyPr anchor="ctr"/>
            <a:lstStyle/>
            <a:p><a:pPr algn="ctr"><a:buNone/></a:pPr><a:r><a:rPr sz="2200"><a:solidFill><a:schemeClr val="lt1"/></a:solidFill><a:latin typeface="Liberation Sans"/></a:rPr><a:t>Nested Group Banner</a:t></a:r></a:p>
          </p:txBody>
        </p:sp>
        <p:sp>
          <p:spPr>
            <a:xfrm><a:off x="457200" y="914400"/><a:ext cx="1371600" cy="914400"/></a:xfrm>
            <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
            <a:solidFill><a:schemeClr val="accent3"/></a:solidFill>
          </p:spPr>
        </p:sp>
        <p:grpSp>
          <p:grpSpPr>
            <a:xfrm>
              <a:off x="2286000" y="914400"/>
              <a:ext cx="1828800" cy="914400"/>
              <a:chOff x="0" y="0"/>
              <a:chExt cx="1828800" cy="914400"/>
            </a:xfrm>
          </p:grpSpPr>
          <p:sp>
            <p:spPr>
              <a:xfrm><a:off x="0" y="0"/><a:ext cx="1828800" cy="914400"/></a:xfrm>
              <a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom>
              <a:solidFill><a:schemeClr val="accent4"/></a:solidFill>
            </p:spPr>
            <p:txBody>
              <a:bodyPr anchor="ctr"/>
              <a:lstStyle/>
              <a:p><a:pPr algn="ctr"><a:buNone/></a:pPr><a:r><a:rPr sz="1700"><a:solidFill><a:schemeClr val="lt1"/></a:solidFill><a:latin typeface="Liberation Sans"/></a:rPr><a:t>Inner Card</a:t></a:r></a:p>
            </p:txBody>
          </p:sp>
        </p:grpSp>
      </p:grpSp>
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

function createImageCroppingSlide() {
  return slideXml(`
    <p:sp>
      <p:spPr>
        <a:xfrm><a:off x="457200" y="342900"/><a:ext cx="8229600" cy="731520"/></a:xfrm>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
      </p:spPr>
      <p:txBody>
        <a:bodyPr anchor="ctr"/>
        <a:lstStyle/>
        <a:p><a:pPr algn="ctr"><a:buNone/></a:pPr><a:r><a:rPr sz="2400"><a:latin typeface="Liberation Sans"/></a:rPr><a:t>Image Cropping Gap Case</a:t></a:r></a:p>
      </p:txBody>
    </p:sp>
    <p:pic>
      <p:blipFill>
        <a:blip r:embed="rId1"/>
        <a:srcRect l="15000" r="15000"/>
        <a:stretch><a:fillRect/></a:stretch>
      </p:blipFill>
      <p:spPr>
        <a:xfrm><a:off x="1600200" y="1600200"/><a:ext cx="2743200" cy="1828800"/></a:xfrm>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
      </p:spPr>
    </p:pic>
  `);
}

function createUnsupportedChartSlide() {
  return slideXml(`
    <p:sp>
      <p:spPr>
        <a:xfrm><a:off x="457200" y="342900"/><a:ext cx="8229600" cy="731520"/></a:xfrm>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
      </p:spPr>
      <p:txBody>
        <a:bodyPr anchor="ctr"/>
        <a:lstStyle/>
        <a:p><a:pPr algn="ctr"><a:buNone/></a:pPr><a:r><a:rPr sz="2400"><a:latin typeface="Liberation Sans"/></a:rPr><a:t>Unsupported Chart Placeholder Case</a:t></a:r></a:p>
      </p:txBody>
    </p:sp>
    <p:graphicFrame>
      <p:xfrm><a:off x="1371600" y="1371600"/><a:ext cx="5486400" cy="3200400"/></p:xfrm>
      <a:graphic>
        <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
          <c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rIdChart1"/>
        </a:graphicData>
      </a:graphic>
    </p:graphicFrame>
  `);
}

function createImagePlacementRelationships() {
  return `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image2.png"/>
</Relationships>`;
}

function createImageCroppingRelationships() {
  return `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
</Relationships>`;
}

function createUnsupportedChartRelationships() {
  return `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdChart1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>
</Relationships>`;
}

function createChartXml() {
  return `<?xml version="1.0" encoding="UTF-8"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart/>
</c:chartSpace>`;
}

function createImageMediaFiles() {
  return {
    "ppt/media/image1.png": createSolidPng("#2563eb"),
    "ppt/media/image2.png": createSolidPng("#f97316"),
  };
}

function createImageCroppingMediaFiles() {
  return {
    "ppt/media/image1.png": createSplitPng("#2563eb", "#f97316"),
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

function createSplitPng(leftHexColor, rightHexColor) {
  const png = new PNG({ width: 80, height: 80 });
  const [lr, lg, lb] = hexToRgb(leftHexColor);
  const [rr, rg, rb] = hexToRgb(rightHexColor);

  for (let y = 0; y < png.height; y += 1) {
    for (let x = 0; x < png.width; x += 1) {
      const index = (png.width * y + x) * 4;
      const [r, g, b] = x < png.width / 2 ? [lr, lg, lb] : [rr, rg, rb];
      png.data[index] = r;
      png.data[index + 1] = g;
      png.data[index + 2] = b;
      png.data[index + 3] = 255;
    }
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

function createThemeFixtureRelationships() {
  return `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
</Relationships>`;
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
