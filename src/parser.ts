import JSZip from "jszip";
import type { GradientFill } from "./types";
import type {
  PptxSlide,
  PptxParseResult,
  SlideElement,
  TextElement,
  Paragraph,
  ImageElement,
  ShapeElement,
  GroupElement,
  TableElement,
  TableRow,
  TableCell,
  SlideBackground,
  GraphicFrameElement,
  TextShapeFrame,
  DiagramData,
  DiagramNode,
  DiagramShape,
} from "./types";

/**
 * Parse a PPTX file and extract slide content
 */
export async function parsePptx(arrayBuffer: ArrayBuffer | Uint8Array): Promise<PptxParseResult> {
  assertParseApisAvailable();
  const zip = await JSZip.loadAsync(arrayBuffer);

  // Get presentation dimensions
  const presentationXml = await getFileContent(zip, "ppt/presentation.xml");
  const dimensions = parsePresentationDimensions(presentationXml);

  // Get slide list
  const slideList = await getSlideList(zip);

  // Extract all images from the PPTX
  const imageMap = await extractImages(zip);

  // Load theme colors for resolving schemeClr references
  const theme = await loadThemeData(zip);

  // Parse each slide
  const slides: PptxSlide[] = [];
  for (let i = 0; i < slideList.length; i++) {
    const slidePath = slideList[i];
    const slideXml = await getFileContent(zip, slidePath);

    // Get slide relationships for this slide
    const slideNumber = i + 1;
    const relsPath = `ppt/slides/_rels/slide${slideNumber}.xml.rels`;
    const relsXml = await getFileContent(zip, relsPath).catch(() => "");
    const relationships = parseRelationships(relsXml);

    const slide = await parseSlide(
      slideXml,
      slideNumber,
      dimensions,
      relationships,
      imageMap,
      zip,
      theme,
    );
    slides.push(slide);
  }

  return {
    slides,
    slideWidth: dimensions.width,
    slideHeight: dimensions.height,
  };
}

function assertParseApisAvailable(): void {
  if (
    typeof DOMParser === "undefined" ||
    typeof FileReader === "undefined" ||
    typeof Blob === "undefined"
  ) {
    throw new Error(
      "pptx-react-renderer requires browser parsing APIs (DOMParser, FileReader, Blob).",
    );
  }
}

/**
 * Load theme colors from theme1.xml
 */
type ThemeData = {
  colors: Map<string, string>;
  fillStyles: Element[];
  lineStyles: Element[];
  backgroundFillStyles: Element[];
  fonts: {
    major?: string;
    minor?: string;
  };
};

async function loadThemeData(zip: JSZip): Promise<ThemeData> {
  const themeColors = new Map<string, string>();
  const themeData: ThemeData = {
    colors: themeColors,
    fillStyles: [],
    lineStyles: [],
    backgroundFillStyles: [],
    fonts: {},
  };

  try {
    const themeXml = await getFileContent(zip, "ppt/theme/theme1.xml");
    const parser = new DOMParser();
    const doc = parser.parseFromString(themeXml, "application/xml");

    // Parse color scheme - clrScheme contains dk1, lt1, dk2, lt2, accent1-6, hlink, folHlink
    const clrScheme = doc.querySelector("clrScheme");
    if (clrScheme) {
      // Map of color names to their RGB values
      const colorElements = [
        "dk1",
        "lt1",
        "dk2",
        "lt2",
        "accent1",
        "accent2",
        "accent3",
        "accent4",
        "accent5",
        "accent6",
        "hlink",
        "folHlink",
      ];

      for (const colorName of colorElements) {
        const colorEl = clrScheme.querySelector(colorName);
        if (colorEl) {
          // Check for srgbClr (direct RGB)
          const srgbClr = colorEl.querySelector("srgbClr");
          if (srgbClr) {
            const val = srgbClr.getAttribute("val");
            if (val) {
              themeColors.set(colorName, `#${val}`);
            }
          }
          // Check for sysClr (system color with lastClr fallback)
          const sysClr = colorEl.querySelector("sysClr");
          if (sysClr) {
            const lastClr = sysClr.getAttribute("lastClr");
            if (lastClr) {
              themeColors.set(colorName, `#${lastClr}`);
            }
          }
        }
      }

      if (themeColors.has("lt1")) themeColors.set("bg1", themeColors.get("lt1")!);
      if (themeColors.has("dk1")) themeColors.set("tx1", themeColors.get("dk1")!);
      if (themeColors.has("lt2")) themeColors.set("bg2", themeColors.get("lt2")!);
      if (themeColors.has("dk2")) themeColors.set("tx2", themeColors.get("dk2")!);
    }

    themeData.fillStyles = Array.from(doc.querySelectorAll("fmtScheme fillStyleLst > *"));
    themeData.lineStyles = Array.from(doc.querySelectorAll("fmtScheme lnStyleLst > ln"));
    themeData.backgroundFillStyles = Array.from(doc.querySelectorAll("fmtScheme bgFillStyleLst > *"));

    themeData.fonts.major =
      doc.querySelector("fontScheme > majorFont > latin")?.getAttribute("typeface") || undefined;
    themeData.fonts.minor =
      doc.querySelector("fontScheme > minorFont > latin")?.getAttribute("typeface") || undefined;
  } catch (error) {
    console.warn("Failed to load theme colors:", error);
  }

  return themeData;
}

/**
 * Extract all images from PPTX and convert to base64
 */
async function extractImages(zip: JSZip): Promise<Map<string, { src: string; mimeType: string }>> {
  const imageMap = new Map<string, { src: string; mimeType: string }>();

  // Find all image files in ppt/media/
  const mediaFiles: string[] = [];
  zip.forEach((relativePath, file) => {
    if (relativePath.startsWith("ppt/media/") && !file.dir) {
      mediaFiles.push(relativePath);
    }
  });

  // Extract each image and convert to base64
  for (const mediaPath of mediaFiles) {
    try {
      const file = zip.file(mediaPath);
      if (!file) continue;

      const blob = await file.async("blob");
      const base64 = await blobToBase64(blob);
      const mimeType = getMimeType(mediaPath);

      // Store with the filename as key (e.g., "image1.png")
      const filename = mediaPath.replace("ppt/media/", "");
      imageMap.set(filename, { src: base64, mimeType });
    } catch (e) {
      console.warn(`Failed to extract image: ${mediaPath}`, e);
    }
  }

  return imageMap;
}

/**
 * Convert Blob to base64 data URL
 */
function blobToBase64(blob: Blob): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onloadend = () => resolve(reader.result as string);
    reader.onerror = reject;
    reader.readAsDataURL(blob);
  });
}

/**
 * Get MIME type from file extension
 */
function getMimeType(filename: string): string {
  const ext = filename.split(".").pop()?.toLowerCase();
  const mimeTypes: Record<string, string> = {
    png: "image/png",
    jpg: "image/jpeg",
    jpeg: "image/jpeg",
    gif: "image/gif",
    bmp: "image/bmp",
    svg: "image/svg+xml",
    emf: "image/emf",
    wmf: "image/wmf",
  };
  return mimeTypes[ext || ""] || "image/png";
}

/**
 * Parse slide relationships XML to map rId to target files
 */
function parseRelationships(relsXml: string): Map<string, string> {
  const relationships = new Map<string, string>();

  if (!relsXml) return relationships;

  const parser = new DOMParser();
  const doc = parser.parseFromString(relsXml, "application/xml");

  doc.querySelectorAll("Relationship").forEach((rel) => {
    const id = rel.getAttribute("Id");
    const target = rel.getAttribute("Target");
    if (id && target) {
      // Store the target filename
      const filename = target.replace(/^\.\.\/media\//, "");
      relationships.set(id, filename);
    }
  });

  return relationships;
}

/**
 * Extract slide dimensions from presentation.xml
 */
function parsePresentationDimensions(xml: string): { width: number; height: number } {
  const parser = new DOMParser();
  const doc = parser.parseFromString(xml, "application/xml");

  const sldSz = doc.querySelector("sldSz");
  if (!sldSz) return { width: 9144000, height: 6858000 }; // Default 16:9

  const cx = parseInt(sldSz.getAttribute("cx") || "9144000", 10);
  const cy = parseInt(sldSz.getAttribute("cy") || "6858000", 10);

  // Convert EMUs to pixels (1 inch = 914400 EMUs, 96 DPI)
  return {
    width: Math.round((cx / 914400) * 96),
    height: Math.round((cy / 914400) * 96),
  };
}

/**
 * Get list of slide files in order
 */
async function getSlideList(zip: JSZip): Promise<string[]> {
  const contentXml = await getFileContent(zip, "[Content_Types].xml");
  const parser = new DOMParser();
  const doc = parser.parseFromString(contentXml, "application/xml");

  const overrides = doc.querySelectorAll("Override");
  const slides: string[] = [];

  overrides.forEach((override) => {
    const partName = override.getAttribute("PartName");
    const contentType = override.getAttribute("ContentType");
    if (
      partName &&
      contentType?.includes("slide") &&
      !contentType.includes("slideLayout") &&
      !contentType.includes("slideMaster")
    ) {
      slides.push(partName.replace(/^\//, ""));
    }
  });

  // Sort by slide number
  return slides.sort((a, b) => {
    const numA = parseInt(a.match(/slide(\d+)/)?.[1] || "0", 10);
    const numB = parseInt(b.match(/slide(\d+)/)?.[1] || "0", 10);
    return numA - numB;
  });
}

/**
 * Parse a single slide XML
 */
async function parseSlide(
  xml: string,
  slideNumber: number,
  dimensions: { width: number; height: number },
  relationships: Map<string, string>,
  imageMap: Map<string, { src: string; mimeType: string }>,
  zip: JSZip,
  theme: ThemeData,
): Promise<PptxSlide> {
  const parser = new DOMParser();
  const doc = parser.parseFromString(xml, "application/xml");

  const elements: SlideElement[] = [];
  const spTree = doc.querySelector("spTree");

  if (spTree) {
    const zIndex = { value: 0 };
    elements.push(
      ...(await parseSpTreeChildren(
        spTree,
        zIndex,
        { relationships, imageMap, zip, theme },
        undefined,
      )),
    );
  }

  // Parse background
  const background = parseBackground(doc, relationships, imageMap);

  return {
    number: slideNumber,
    width: dimensions.width,
    height: dimensions.height,
    elements,
    background,
  };
}

type ElementTransform = {
  offsetX: number;
  offsetY: number;
  scaleX: number;
  scaleY: number;
};

const DEFAULT_TEXT_INSET_EMU = {
  left: 91440,
  right: 91440,
  top: 45720,
  bottom: 45720,
};

const EMPTY_TEXT_SHAPE_FALLBACK_MIN_AREA_PX = 650000;

type ElementContext = {
  relationships: Map<string, string>;
  imageMap: Map<string, { src: string; mimeType: string }>;
  zip: JSZip;
  theme: ThemeData;
};

async function parseSpTreeChildren(
  container: Element,
  zIndex: { value: number },
  context: ElementContext,
  transform?: ElementTransform,
): Promise<SlideElement[]> {
  const elements: SlideElement[] = [];

  for (const child of Array.from(container.children)) {
    switch (child.localName) {
      case "sp": {
        const element = parseShape(child, zIndex.value++, context.theme, transform);
        if (element) elements.push(element);
        break;
      }
      case "pic": {
        const element = parsePicture(
          child,
          zIndex.value++,
          context.relationships,
          context.imageMap,
          transform,
        );
        if (element) elements.push(element);
        break;
      }
      case "grpSp": {
        const element = await parseGroup(child, zIndex, context, transform);
        if (element) elements.push(element);
        break;
      }
      case "graphicFrame": {
        const element = await parseGraphicFrame(
          child,
          zIndex.value++,
          context.relationships,
          context.zip,
          context.theme.colors,
          transform,
        );
        if (element) elements.push(element);
        break;
      }
      default:
        break;
    }
  }

  return elements;
}

/**
 * Parse a shape element
 */
function parseShape(
  sp: Element,
  zIndex: number,
  theme: ThemeData,
  transform?: ElementTransform,
): SlideElement | null {
  const xfrm = sp.querySelector(":scope > spPr > xfrm");
  if (!xfrm) return null;

  const position = parsePosition(xfrm, transform);
  const styleRefs = parseShapeStyleRefs(sp, theme);

  // Check if it's a text element
  const txBody = sp.querySelector("txBody");
  if (txBody) {
    if (hasRenderableTextContent(txBody)) {
      return parseTextElement(
        txBody,
        position,
        zIndex,
        theme.colors,
        parseTextShapeFrame(sp, theme, styleRefs),
        styleRefs?.text,
      );
    }

    if (isPlaceholderShape(sp)) {
      return parseTextElement(txBody, position, zIndex, theme.colors, undefined, styleRefs?.text);
    }

    if (!hasExplicitShapeVisualStyle(sp)) {
      return parseTextElement(txBody, position, zIndex, theme.colors, undefined, styleRefs?.text);
    }

    if (!isLargeVisualShape(position)) {
      return parseTextElement(txBody, position, zIndex, theme.colors, undefined, styleRefs?.text);
    }
  }

  // Check if it's a table
  const tbl = sp.querySelector("tbl");
  if (tbl) {
    return parseTable(tbl, position, zIndex, theme.colors);
  }

  // Otherwise it's a shape
  const prstGeom = sp.querySelector(":scope > spPr > prstGeom");
  const custGeom = sp.querySelector(":scope > spPr > custGeom");
  const shapeType = prstGeom?.getAttribute("prst") || (custGeom ? "custom" : "rect");

  let pathData: import("./types").ShapePath[] | undefined;
  let viewBox: { x: number; y: number; width: number; height: number } | undefined;

  if (custGeom) {
    const pathLst = custGeom.querySelector(":scope > pathLst");
    if (pathLst) {
      pathData = parsePathList(pathLst, theme.colors);

      const firstPath = pathLst.querySelector(":scope > path");
      const viewWidth = firstPath ? parseFloat(firstPath.getAttribute("w") || "0") : 0;
      const viewHeight = firstPath ? parseFloat(firstPath.getAttribute("h") || "0") : 0;

      if (viewWidth > 0 && viewHeight > 0) {
        viewBox = { x: 0, y: 0, width: viewWidth, height: viewHeight };
      }
    }
  }

  const themeFill = shouldApplyThemeFillFallback(sp) ? styleRefs?.frame : undefined;
  const themeLine = shouldApplyThemeLineFallback(sp) ? styleRefs?.frame : undefined;

  const shapeElement: ShapeElement = {
    type: "shape",
    ...position,
    zIndex,
    shapeType,
    fillColor: parseFillColor(sp, theme.colors) || themeFill?.fillColor,
    fillType: themeFill?.fillType,
    gradient: themeFill?.gradient,
    strokeColor: parseStrokeColor(sp, theme.colors) || themeLine?.strokeColor,
    strokeWidth: parseStrokeWidth(sp) || themeLine?.strokeWidth,
    strokeDash: parseShapeStrokeDash(sp) || themeLine?.strokeDash,
    roundedCorners: parseRoundedCornersFromGeom(prstGeom) || styleRefs?.frame?.roundedCorners,
    pathData,
    viewBox,
  };

  return shapeElement;
}

function hasRenderableTextContent(txBody: Element): boolean {
  for (const node of Array.from(txBody.querySelectorAll("t"))) {
    const text = node.textContent || "";
    if (text.trim().length > 0) return true;
  }

  if (txBody.querySelector("fld") || txBody.querySelector("br")) {
    return true;
  }

  return false;
}

function hasExplicitShapeVisualStyle(sp: Element): boolean {
  const spPr = sp.querySelector(":scope > spPr");
  if (!spPr) return false;

  if (spPr.querySelector(":scope > solidFill, :scope > gradFill, :scope > pattFill, :scope > blipFill")) {
    return true;
  }

  const ln = spPr.querySelector(":scope > ln");
  if (ln?.querySelector(":scope > solidFill, :scope > gradFill, :scope > noFill")) {
    return true;
  }

  return false;
}

function shouldApplyThemeFillFallback(sp: Element): boolean {
  const spPr = sp.querySelector(":scope > spPr");
  if (!spPr) return false;

  if (sp.getAttribute("useBgFill") === "1") {
    return false;
  }

  if (spPr.querySelector(":scope > noFill")) {
    return false;
  }

  return !spPr.querySelector(":scope > solidFill, :scope > gradFill, :scope > pattFill, :scope > blipFill");
}

function shouldApplyThemeLineFallback(sp: Element): boolean {
  const ln = sp.querySelector(":scope > spPr > ln");
  if (!ln) return true;

  return !ln.querySelector(":scope > noFill");
}

type ShapeStyleRefs = {
  frame?: {
    fillType?: "solid" | "gradient";
    fillColor?: string;
    gradient?: GradientFill;
    strokeColor?: string;
    strokeWidth?: number;
    strokeDash?: string;
    roundedCorners?: number;
  };
  text?: {
    color?: string;
    fontFamily?: string;
  };
};

function parseShapeStyleRefs(sp: Element, theme: ThemeData): ShapeStyleRefs | undefined {
  const style = sp.querySelector(":scope > style");
  if (!style) return undefined;

  const frame: NonNullable<ShapeStyleRefs["frame"]> = {};
  const text: NonNullable<ShapeStyleRefs["text"]> = {};

  const fillRef = style.querySelector(":scope > fillRef");
  if (fillRef) {
    const fillStyle = getThemeFillStyle(fillRef, theme);
    if (fillStyle?.fillType === "gradient" && fillStyle.gradient) {
      frame.fillType = "gradient";
      frame.gradient = fillStyle.gradient;
    } else if (fillStyle?.fillColor) {
      frame.fillType = "solid";
      frame.fillColor = fillStyle.fillColor;
    }
  }

  const lnRef = style.querySelector(":scope > lnRef");
  if (lnRef) {
    const lineStyle = getThemeLineStyle(lnRef, theme);
    if (lineStyle) {
      frame.strokeColor = lineStyle.color;
      frame.strokeWidth = lineStyle.width;
      frame.strokeDash = lineStyle.dash;
    }
  }

  const fontRef = style.querySelector(":scope > fontRef");
  if (fontRef) {
    text.color = parseColorValue(fontRef, theme.colors);
    const idx = fontRef.getAttribute("idx");
    if (idx === "minor") text.fontFamily = theme.fonts.minor;
    if (idx === "major") text.fontFamily = theme.fonts.major;
  }

  if (sp.querySelector(":scope > spPr > prstGeom")) {
    frame.roundedCorners = parseRoundedCornersFromGeom(
      sp.querySelector(":scope > spPr > prstGeom"),
    );
  }

  return {
    frame: Object.values(frame).some((value) => value !== undefined) ? frame : undefined,
    text: Object.values(text).some((value) => value !== undefined) ? text : undefined,
  };
}

function getThemeFillStyle(
  fillRef: Element,
  theme: ThemeData,
): { fillType?: "solid" | "gradient"; fillColor?: string; gradient?: GradientFill } | undefined {
  const idx = parseOptionalInt(fillRef.getAttribute("idx"));
  if (idx === undefined) return undefined;

  const styleElement =
    idx >= 1001
      ? theme.backgroundFillStyles[idx - 1001]
      : idx >= 1
        ? theme.fillStyles[idx - 1]
        : undefined;
  if (!styleElement) return undefined;

  const colorOverrides = new Map<string, string>();
  const placeholderColor = parseColorValue(fillRef, theme.colors);
  if (placeholderColor) {
    colorOverrides.set("phClr", placeholderColor);
  }

  if (styleElement.localName === "solidFill") {
    const fillColor = parseColorValue(styleElement, theme.colors, { colorOverrides });
    return fillColor ? { fillType: "solid", fillColor } : undefined;
  }

  if (styleElement.localName === "gradFill") {
    const gradient = parseGradientFill(styleElement, theme.colors, colorOverrides);
    return gradient ? { fillType: "gradient", gradient } : undefined;
  }

  return undefined;
}

function getThemeLineStyle(
  lnRef: Element,
  theme: ThemeData,
): { color?: string; width?: number; dash?: string } | undefined {
  const idx = parseOptionalInt(lnRef.getAttribute("idx"));
  if (idx === undefined || idx < 1) return undefined;

  const lineElement = theme.lineStyles[idx - 1];
  if (!lineElement) return undefined;

  const colorOverrides = new Map<string, string>();
  const placeholderColor = parseColorValue(lnRef, theme.colors);
  if (placeholderColor) {
    colorOverrides.set("phClr", placeholderColor);
  }

  return {
    color: parseColorValue(lineElement, theme.colors, { colorOverrides, includeAlpha: true }),
    width: parseOptionalInt(lineElement.getAttribute("w"))
      ? parseInt(lineElement.getAttribute("w") || "0", 10) / 12700
      : undefined,
    dash: parseThemeDash(lineElement),
  };
}

function parseGradientFill(
  gradFill: Element,
  themeColors: Map<string, string>,
  colorOverrides: Map<string, string>,
): GradientFill | undefined {
  const stops = Array.from(gradFill.querySelectorAll(":scope > gsLst > gs"))
    .map((stopEl) => {
      const color = parseColorValue(stopEl, themeColors, { colorOverrides });
      if (!color) return undefined;
      return {
        position: (parseOptionalInt(stopEl.getAttribute("pos")) ?? 0) / 100000,
        color,
      };
    })
    .filter((stop): stop is NonNullable<typeof stop> => stop !== undefined);

  if (stops.length === 0) return undefined;

  const lin = gradFill.querySelector(":scope > lin");
  if (lin) {
    return {
      type: "linear",
      angle: ((parseOptionalInt(lin.getAttribute("ang")) ?? 0) / 60000) % 360,
      stops,
    };
  }

  return {
    type: "radial",
    stops,
  };
}

function parseTextShapeFrame(
  sp: Element,
  theme: ThemeData,
  styleRefs?: ShapeStyleRefs,
): TextShapeFrame | undefined {
  if (!hasExplicitShapeVisualStyle(sp) && !styleRefs?.frame) {
    return undefined;
  }

  const prstGeom = sp.querySelector(":scope > spPr > prstGeom");
  if (!prstGeom) {
    return undefined;
  }

  const themeFill = shouldApplyThemeFillFallback(sp) ? styleRefs?.frame : undefined;
  const themeLine = shouldApplyThemeLineFallback(sp) ? styleRefs?.frame : undefined;

  return {
    shapeType: prstGeom.getAttribute("prst") || "rect",
    fillColor: parseFillColor(sp, theme.colors) || themeFill?.fillColor,
    strokeColor: parseStrokeColor(sp, theme.colors) || themeLine?.strokeColor,
    strokeWidth: parseStrokeWidth(sp) || themeLine?.strokeWidth,
    strokeDash: parseShapeStrokeDash(sp) || themeLine?.strokeDash,
    roundedCorners: parseRoundedCornersFromGeom(prstGeom) || styleRefs?.frame?.roundedCorners,
  };
}

function isPlaceholderShape(sp: Element): boolean {
  return sp.querySelector(":scope > nvSpPr > nvPr > ph") !== null;
}

function isLargeVisualShape(position: {
  x: number;
  y: number;
  width: number;
  height: number;
}): boolean {
  return position.width * position.height >= EMPTY_TEXT_SHAPE_FALLBACK_MIN_AREA_PX;
}

/**
 * Parse text element with rich formatting
 */
function parseTextElement(
  txBody: Element,
  position: { x: number; y: number; width: number; height: number },
  zIndex: number,
  themeColors: Map<string, string>,
  frame?: TextShapeFrame,
  textStyle?: ShapeStyleRefs["text"],
): TextElement {
  // Get default properties
  const defaultFontSize = parseFontSize(txBody);
  const defaultFontFamily = parseFontFamily(txBody) || textStyle?.fontFamily;
  const defaultColor = parseTextColor(txBody, themeColors) || textStyle?.color;

  // Parse paragraphs with rich text runs
  const paragraphs: Paragraph[] = [];
  txBody.querySelectorAll("p").forEach((p) => {
    const paragraph = parseParagraph(p, themeColors);
    if (paragraph.runs.length > 0) {
      paragraphs.push(paragraph);
    }
  });

  // Flatten runs for simple text elements - only include actual text runs (not fields or breaks)
  const runs: import("./types").TextRun[] = paragraphs
    .flatMap((p) => p.runs)
    .filter((r): r is import("./types").TextRun => {
      // TextRun either has type === 'run' or no type property (legacy)
      // TextField has type === 'field'
      // LineBreak has type === 'break'
      return r.type === "run" || r.type === undefined;
    });

  // Parse body properties for overflow and insets
  const bodyPr = txBody.querySelector("bodyPr");
  const wrap = bodyPr?.getAttribute("wrap");
  const vertOverflow = bodyPr?.getAttribute("vertOverflow") as
    | "overflow"
    | "clip"
    | "ellipsis"
    | undefined;
  const horzOverflow = bodyPr?.getAttribute("horzOverflow") as "overflow" | "clip" | undefined;

  // Parse insets (margins) - convert EMUs to points (1 pt = 12700 EMUs)
  const lIns = bodyPr?.getAttribute("lIns");
  const tIns = bodyPr?.getAttribute("tIns");
  const rIns = bodyPr?.getAttribute("rIns");
  const bIns = bodyPr?.getAttribute("bIns");

  const leftInsetEmu = parseOptionalInt(lIns) ?? DEFAULT_TEXT_INSET_EMU.left;
  const topInsetEmu = parseOptionalInt(tIns) ?? DEFAULT_TEXT_INSET_EMU.top;
  const rightInsetEmu = parseOptionalInt(rIns) ?? DEFAULT_TEXT_INSET_EMU.right;
  const bottomInsetEmu = parseOptionalInt(bIns) ?? DEFAULT_TEXT_INSET_EMU.bottom;

  return {
    type: "text",
    ...position,
    zIndex,
    runs,
    paragraphs: paragraphs.length > 1 ? paragraphs : undefined,
    frame,
    defaultFontSize,
    defaultFontFamily,
    defaultColor,
    align: parseTextAlign(txBody),
    verticalAlign: parseVerticalAlign(txBody),
    lineSpacing: parseLineSpacing(txBody),
    spaceBefore: parseSpaceBefore(txBody),
    spaceAfter: parseSpaceAfter(txBody),
    indent: parseIndent(txBody),
    wrap: wrap !== "none",
    vertOverflow: vertOverflow || "overflow",
    horzOverflow: horzOverflow || "overflow",
    leftInset: leftInsetEmu / 12700,
    topInset: topInsetEmu / 12700,
    rightInset: rightInsetEmu / 12700,
    bottomInset: bottomInsetEmu / 12700,
  };
}

/**
 * Parse a paragraph with text runs
 */
function parseParagraph(p: Element, themeColors: Map<string, string>): Paragraph {
  const runs: import("./types").ContentElement[] = [];

  // Get all child nodes of the paragraph
  Array.from(p.childNodes).forEach((node) => {
    if (node.nodeType === Node.ELEMENT_NODE) {
      const el = node as Element;
      const tagName = el.tagName.toLowerCase();

      // Parse text runs (<a:r> or <r>)
      if (tagName === "r" || tagName.endsWith(":r")) {
        const run = parseTextRun(el, themeColors);
        if (run.text || run.text === "") {
          runs.push(run);
        }
      }
      // Parse text fields (<a:fld> or <fld>)
      else if (tagName === "fld" || tagName.endsWith(":fld")) {
        const field = parseTextField(el, themeColors);
        if (field.text) {
          runs.push(field);
        }
      }
      // Parse line breaks (<a:br> or <br>)
      else if (tagName === "br" || tagName.endsWith(":br")) {
        runs.push({ type: "break" });
      }
    }
  });

  // Get paragraph properties
  const pPr = p.querySelector("pPr");

  // Parse comprehensive bullet formatting
  const bulletInfo = parseBulletInfo(pPr, themeColors);

  return {
    runs,
    align: parseParagraphAlign(pPr),
    bullet: bulletInfo.hasBullet,
    bulletChar: bulletInfo.char,
    bulletFont: bulletInfo.font,
    bulletSize: bulletInfo.size,
    bulletColor: bulletInfo.color,
    autoNumbering: bulletInfo.autoNumbering,
    level: parseLevel(pPr),
    leftMargin: parseLeftMargin(pPr),
    rightMargin: parseRightMargin(pPr),
    indent: parseIndentation(pPr),
    spaceBefore: parseParagraphSpaceBefore(pPr),
    spaceAfter: parseParagraphSpaceAfter(pPr),
    lineSpacing: parseParagraphLineSpacing(pPr),
    tabStops: parseTabStops(pPr),
    defaultRunProps: parseDefaultRunProps(pPr, themeColors),
  };
}

/**
 * Parse a single text run with formatting
 */
function parseTextRun(r: Element, themeColors: Map<string, string>): import("./types").TextRun {
  const text = r.querySelector("t")?.textContent || "";

  // Get run properties
  const rPr = r.querySelector("rPr");

  return {
    type: "run",
    text,
    fontSize: parseRunFontSize(rPr),
    fontFamily: parseRunFontFamily(rPr),
    bold: hasRunFormatting(rPr, "b"),
    italic: hasRunFormatting(rPr, "i"),
    underline: parseRunUnderline(rPr),
    color: parseRunColor(rPr, themeColors),
    highlight: parseRunHighlight(rPr),
    baseline: parseRunBaseline(rPr),
  };
}

/**
 * Parse a text field element (for slide numbers, dates, etc.)
 */
function parseTextField(
  fld: Element,
  themeColors: Map<string, string>,
): import("./types").TextField {
  const text = fld.querySelector("t")?.textContent || "";
  const fieldType = fld.getAttribute("type") || undefined;

  // Get field properties (same structure as run properties)
  const rPr = fld.querySelector("rPr");

  return {
    type: "field",
    text,
    fieldType,
    fontSize: parseRunFontSize(rPr),
    fontFamily: parseRunFontFamily(rPr),
    bold: hasRunFormatting(rPr, "b"),
    italic: hasRunFormatting(rPr, "i"),
    underline: parseRunUnderline(rPr),
    color: parseRunColor(rPr, themeColors),
  };
}

/**
 * Comprehensive bullet information parsing
 */
function parseBulletInfo(
  pPr: Element | null,
  themeColors: Map<string, string>,
): {
  hasBullet: boolean;
  char?: string;
  font?: string;
  size?: number;
  color?: string;
  autoNumbering?: {
    type: "arabic" | "roman" | "alphaLc" | "alphaUc" | "ordText";
    startAt?: number;
  };
} {
  if (!pPr) {
    return { hasBullet: false };
  }

  // Check for buNone (explicit no bullet)
  if (pPr.querySelector("buNone")) {
    return { hasBullet: false };
  }

  // Check for character bullet
  const buChar = pPr.querySelector("buChar");
  if (buChar) {
    const char = buChar.getAttribute("char") || "•";

    // Get bullet font
    const buFont = pPr.querySelector("buFont");
    const font = buFont?.getAttribute("typeface") || undefined;

    // Get bullet size
    const buSzPct = pPr.querySelector("buSzPct");
    const size = buSzPct ? parseInt(buSzPct.getAttribute("val") || "100000", 10) / 1000 : undefined;

    // Get bullet color
    const buClr = pPr.querySelector("buClr srgbClr, buClr schemeClr");
    let color: string | undefined;
    if (buClr) color = parseColorValue(buClr, themeColors);

    return {
      hasBullet: true,
      char,
      font,
      size,
      color,
    };
  }

  // Check for auto-numbering
  const buAutoNum = pPr.querySelector("buAutoNum");
  if (buAutoNum) {
    const type = buAutoNum.getAttribute("type") as
      | "arabic"
      | "roman"
      | "alphaLc"
      | "alphaUc"
      | "ordText"
      | null;
    const startAt = buAutoNum.getAttribute("startAt");

    if (type) {
      return {
        hasBullet: true,
        autoNumbering: {
          type,
          startAt: startAt ? parseInt(startAt, 10) : 1,
        },
      };
    }
  }

  // Check for picture bullet (buBlip)
  const buBlip = pPr.querySelector("buBlip");
  if (buBlip) {
    return { hasBullet: true };
  }

  return { hasBullet: false };
}

/**
 * Parse left margin (marL attribute)
 */
function parseLeftMargin(pPr: Element | null): number | undefined {
  if (!pPr) return undefined;
  const marL = pPr.getAttribute("marL");
  return marL ? parseInt(marL, 10) : undefined;
}

/**
 * Parse right margin (marR attribute)
 */
function parseRightMargin(pPr: Element | null): number | undefined {
  if (!pPr) return undefined;
  const marR = pPr.getAttribute("marR");
  return marR ? parseInt(marR, 10) : undefined;
}

/**
 * Parse first line indent
 */
function parseIndentation(pPr: Element | null): number | undefined {
  if (!pPr) return undefined;
  const indent = pPr.getAttribute("indent");
  return indent ? parseInt(indent, 10) : undefined;
}

/**
 * Parse tab stops
 */
function parseTabStops(pPr: Element | null): import("./types").TabStop[] | undefined {
  if (!pPr) return undefined;

  const tabLst = pPr.querySelector("tabLst");
  if (!tabLst) return undefined;

  const tabStops: import("./types").TabStop[] = [];
  tabLst.querySelectorAll("tab").forEach((tab) => {
    const pos = tab.getAttribute("pos");
    const type = tab.getAttribute("algn") as import("./types").TabStop["type"] | null;
    if (pos && type) {
      tabStops.push({
        position: parseInt(pos, 10),
        type,
      });
    }
  });

  return tabStops.length > 0 ? tabStops : undefined;
}

/**
 * Parse default run properties from paragraph
 */
function parseDefaultRunProps(
  pPr: Element | null,
  themeColors: Map<string, string>,
): import("./types").TextRun | undefined {
  if (!pPr) return undefined;

  const defRPr = pPr.querySelector("defRPr");
  if (!defRPr) return undefined;

  return {
    type: "run",
    text: "",
    fontSize: parseRunFontSize(defRPr),
    fontFamily: parseRunFontFamily(defRPr),
    bold: hasRunFormatting(defRPr, "b"),
    italic: hasRunFormatting(defRPr, "i"),
    underline: parseRunUnderline(defRPr),
    color: parseRunColor(defRPr, themeColors),
  };
}

/**
 * Parse font size from run properties
 */
function parseRunFontSize(rPr: Element | null): number | undefined {
  if (!rPr) return undefined;

  const sz = rPr.getAttribute("sz");
  if (sz) {
    return parseInt(sz, 10) / 100; // Convert hundredths to points
  }

  return undefined;
}

/**
 * Parse font family from run properties
 */
function parseRunFontFamily(rPr: Element | null): string | undefined {
  if (!rPr) return undefined;

  // Check latin font first
  const latin = rPr.querySelector("latin");
  if (latin) {
    return latin.getAttribute("typeface") || undefined;
  }

  // Check default font
  const ea = rPr.querySelector("ea");
  if (ea) {
    return ea.getAttribute("typeface") || undefined;
  }

  return undefined;
}

/**
 * Check for run-level formatting
 */
function hasRunFormatting(rPr: Element | null, attribute: string): boolean {
  if (!rPr) return false;
  return rPr.getAttribute(attribute) === "1";
}

/**
 * Parse underline style
 */
function parseRunUnderline(rPr: Element | null): boolean {
  if (!rPr) return false;
  return rPr.getAttribute("u") !== "none" && rPr.getAttribute("u") !== null;
}

/**
 * Parse run color
 */
function parseRunColor(rPr: Element | null, themeColors: Map<string, string>): string | undefined {
  if (!rPr) return undefined;

  const solidFill = rPr.querySelector("solidFill");
  if (solidFill) {
    return parseColorValue(solidFill, themeColors);
  }

  return undefined;
}

/**
 * Parse highlight color
 */
function parseRunHighlight(rPr: Element | null): string | undefined {
  if (!rPr) return undefined;

  const highlight = rPr.querySelector("highlight");
  if (highlight) {
    const val = highlight.getAttribute("val");
    if (val) return val; // Usually a color name like 'yellow'
  }

  return undefined;
}

/**
 * Parse baseline (superscript/subscript)
 */
function parseRunBaseline(rPr: Element | null): number | undefined {
  if (!rPr) return undefined;

  const baseline = rPr.getAttribute("baseline");
  if (baseline) {
    return parseInt(baseline, 10) / 1000; // Convert to percentage
  }

  return undefined;
}

/**
 * Parse a group element
 */
function parsePicture(
  pic: Element,
  zIndex: number,
  relationships: Map<string, string>,
  imageMap: Map<string, { src: string; mimeType: string }>,
  transform?: ElementTransform,
): ImageElement | null {
  const xfrm = pic.querySelector(":scope > spPr > xfrm");
  if (!xfrm) return null;

  const position = parsePosition(xfrm, transform);

  // Get image reference
  const blip = pic.querySelector("blip");
  if (!blip) return null;

  const embed = blip.getAttribute("r:embed");
  if (!embed) return null;

  // Look up the image filename from relationships
  const imageFilename = relationships.get(embed);
  if (!imageFilename) {
    // Return placeholder if we can't find the image
    return {
      type: "image",
      ...position,
      zIndex,
      src: "",
      mimeType: "image/png",
    };
  }

  // Look up the actual image data
  const imageData = imageMap.get(imageFilename);

  const imageElement: ImageElement = {
    type: "image",
    ...position,
    zIndex,
    src: imageData?.src || "",
    mimeType: imageData?.mimeType || "image/png",
  };

  return imageElement;
}

/**
 * Parse a group element
 */
async function parseGroup(
  grpSp: Element,
  zIndex: { value: number },
  context: ElementContext,
  parentTransform?: ElementTransform,
): Promise<GroupElement | null> {
  const groupZIndex = zIndex.value++;
  const xfrm = grpSp.querySelector(":scope > grpSpPr > xfrm");
  if (!xfrm) return null;

  const position = parsePosition(xfrm, parentTransform);
  const childTransform = parseGroupChildTransform(xfrm, parentTransform);
  const children = await parseSpTreeChildren(grpSp, zIndex, context, childTransform);

  return {
    type: "group",
    ...position,
    zIndex: groupZIndex,
    children,
  };
}

/**
 * Parse paragraph alignment
 */
function parseParagraphAlign(
  pPr: Element | null,
): "left" | "center" | "right" | "justify" | undefined {
  if (!pPr) return undefined;
  const algn = pPr.getAttribute("algn");
  if (algn === "ctr") return "center";
  if (algn === "r") return "right";
  if (algn === "just") return "justify";
  if (algn === "l") return "left";
  return undefined;
}

/**
 * Parse paragraph level
 */
function parseLevel(pPr: Element | null): number | undefined {
  if (!pPr) return undefined;
  const lvl = pPr.getAttribute("lvl");
  return lvl ? parseInt(lvl, 10) : undefined;
}

/**
 * Parse paragraph space before
 */
function parseParagraphSpaceBefore(pPr: Element | null): number | undefined {
  if (!pPr) return undefined;
  const spcBef = pPr.querySelector("spcBef > spcPts");
  if (spcBef) {
    const val = spcBef.getAttribute("val");
    if (val) return parseInt(val, 10) / 100; // Convert to points
  }
  return undefined;
}

/**
 * Parse paragraph space after
 */
function parseParagraphSpaceAfter(pPr: Element | null): number | undefined {
  if (!pPr) return undefined;
  const spcAft = pPr.querySelector("spcAft > spcPts");
  if (spcAft) {
    const val = spcAft.getAttribute("val");
    if (val) return parseInt(val, 10) / 100; // Convert to points
  }
  return undefined;
}

/**
 * Parse paragraph line spacing
 */
function parseParagraphLineSpacing(pPr: Element | null): number | undefined {
  if (!pPr) return undefined;
  const lnSpc = pPr.querySelector("lnSpc > spcPct");
  if (lnSpc) {
    const val = lnSpc.getAttribute("val");
    if (val) return parseInt(val, 10) / 100000; // Convert to multiplier
  }
  return undefined;
}

/**
 * Parse text alignment from txBody
 */
function parseTextAlign(txBody: Element): "left" | "center" | "right" | "justify" | undefined {
  const p = txBody.querySelector("p");
  const pPr = p?.querySelector("pPr");
  return parseParagraphAlign(pPr);
}

/**
 * Parse vertical alignment from txBody
 */
function parseVerticalAlign(txBody: Element): "top" | "middle" | "bottom" | undefined {
  const bodyPr = txBody.querySelector("bodyPr");
  if (!bodyPr) return undefined;

  const anchor = bodyPr.getAttribute("anchor");
  if (anchor === "ctr") return "middle";
  if (anchor === "b") return "bottom";
  if (anchor === "t") return "top";
  return undefined;
}

/**
 * Parse line spacing from txBody
 */
function parseLineSpacing(txBody: Element): number | undefined {
  const defPPr = txBody.querySelector("defPPr");
  if (!defPPr) return undefined;

  const lnSpc = defPPr.querySelector("lnSpc > spcPct");
  if (lnSpc) {
    const val = lnSpc.getAttribute("val");
    if (val) return parseInt(val, 10) / 100000;
  }

  return undefined;
}

/**
 * Parse space before from txBody
 */
function parseSpaceBefore(txBody: Element): number | undefined {
  const defPPr = txBody.querySelector("defPPr");
  if (!defPPr) return undefined;

  const spcBef = defPPr.querySelector("spcBef > spcPts");
  if (spcBef) {
    const val = spcBef.getAttribute("val");
    if (val) return parseInt(val, 10) / 100;
  }

  return undefined;
}

/**
 * Parse space after from txBody
 */
function parseSpaceAfter(txBody: Element): number | undefined {
  const defPPr = txBody.querySelector("defPPr");
  if (!defPPr) return undefined;

  const spcAft = defPPr.querySelector("spcAft > spcPts");
  if (spcAft) {
    const val = spcAft.getAttribute("val");
    if (val) return parseInt(val, 10) / 100;
  }

  return undefined;
}

/**
 * Parse indent from txBody
 */
function parseIndent(txBody: Element): number | undefined {
  const defPPr = txBody.querySelector("defPPr");
  if (!defPPr) return undefined;

  const marL = defPPr.getAttribute("marL");
  if (marL) {
    return parseInt(marL, 10) / 12700; // Convert EMUs to points
  }

  return undefined;
}

/**
 * Parse slide background
 */
function parseBackground(
  doc: Document,
  relationships: Map<string, string>,
  imageMap: Map<string, { src: string; mimeType: string }>,
): SlideBackground | undefined {
  const cSld = doc.querySelector("cSld");
  if (!cSld) return undefined;

  const bg = cSld.querySelector("bg");
  if (!bg) return undefined;

  // Check for solid fill
  const solidFill = bg.querySelector("solidFill");
  if (solidFill) {
    const srgbClr = solidFill.querySelector("srgbClr");
    if (srgbClr) {
      const val = srgbClr.getAttribute("val");
      if (val) {
        return {
          type: "solid",
          color: `#${val}`,
        };
      }
    }
  }

  // Check for image background
  const blipFill = bg.querySelector("blipFill");
  if (blipFill) {
    const blip = blipFill.querySelector("blip");
    if (blip) {
      const embed = blip.getAttribute("r:embed");
      if (embed) {
        const imageFilename = relationships.get(embed);
        if (imageFilename) {
          const imageData = imageMap.get(imageFilename);
          if (imageData) {
            return {
              type: "image",
              imageSrc: imageData.src,
            };
          }
        }
      }
    }
  }

  return undefined;
}

/**
 * Parse table element
 */
function parseTable(
  tbl: Element,
  position: { x: number; y: number; width: number; height: number },
  zIndex: number,
  themeColors: Map<string, string>,
): TableElement {
  const columnWidths = Array.from(tbl.querySelectorAll(":scope > tblGrid > gridCol")).map((gridCol) =>
    emuToPx(parseInt(gridCol.getAttribute("w") || "0", 10)),
  );
  const rows: TableRow[] = [];

  tbl.querySelectorAll(":scope > tr").forEach((tr) => {
    const cells: TableCell[] = [];

    tr.querySelectorAll(":scope > tc").forEach((tc) => {
      const cell = parseTableCell(tc, themeColors);
      if (cell) cells.push(cell);
    });

    rows.push({
      cells,
      height: emuToPx(parseInt(tr.getAttribute("h") || "0", 10)),
    });
  });

  return {
    type: "table",
    ...position,
    zIndex,
    rows,
    columnWidths,
    backgroundColor: parseTableBackground(tbl, themeColors),
  };
}

function parseTableCell(tc: Element, themeColors: Map<string, string>): TableCell | null {
  const txBody = tc.querySelector(":scope > txBody");
  const tcPr = tc.querySelector(":scope > tcPr");
  const paragraphs = txBody
    ? Array.from(txBody.querySelectorAll(":scope > p"))
        .map((p) => parseParagraph(p, themeColors))
        .filter((p) => p.runs.length > 0)
    : [];

  return {
    content: txBody ? extractTextContent(txBody) : extractTextContent(tc),
    paragraphs: paragraphs.length > 0 ? paragraphs : undefined,
    defaultFontSize: txBody ? parseFontSize(txBody) : undefined,
    defaultFontFamily: txBody ? parseFontFamily(txBody) : undefined,
    defaultColor: txBody ? parseTextColor(txBody, themeColors) : undefined,
    align: txBody ? parseTextAlign(txBody) : undefined,
    verticalAlign: parseTableCellVerticalAlign(tcPr),
    padding: parseTableCellPadding(tcPr),
    rowSpan: parseOptionalInt(tcPr?.getAttribute("rowSpan")),
    colSpan: parseOptionalInt(tcPr?.getAttribute("gridSpan")),
    backgroundColor: tcPr ? parseTableCellFill(tcPr, themeColors) : undefined,
    borders: tcPr ? parseTableCellBorders(tcPr, themeColors) : undefined,
  };
}

function parseTableCellVerticalAlign(tcPr: Element | null): TableCell["verticalAlign"] {
  const anchor = tcPr?.getAttribute("anchor");
  if (anchor === "ctr") return "middle";
  if (anchor === "b") return "bottom";
  return "top";
}

function parseTableCellPadding(tcPr: Element | null): TableCell["padding"] | undefined {
  if (!tcPr) return undefined;

  const left = parseOptionalInt(tcPr.getAttribute("marL"));
  const right = parseOptionalInt(tcPr.getAttribute("marR"));
  const top = parseOptionalInt(tcPr.getAttribute("marT"));
  const bottom = parseOptionalInt(tcPr.getAttribute("marB"));

  if (left === undefined && right === undefined && top === undefined && bottom === undefined) {
    return undefined;
  }

  return {
    left: left !== undefined ? left / 12700 : undefined,
    right: right !== undefined ? right / 12700 : undefined,
    top: top !== undefined ? top / 12700 : undefined,
    bottom: bottom !== undefined ? bottom / 12700 : undefined,
  };
}

function parseTableCellFill(tcPr: Element, themeColors: Map<string, string>): string | undefined {
  const solidFill = tcPr.querySelector(":scope > solidFill");
  return solidFill ? parseColorValue(solidFill, themeColors) : undefined;
}

function parseTableBackground(tbl: Element, themeColors: Map<string, string>): string | undefined {
  const solidFill = tbl.querySelector(":scope > tblPr > solidFill");
  return solidFill ? parseColorValue(solidFill, themeColors) : undefined;
}

function parseTableCellBorders(
  tcPr: Element,
  themeColors: Map<string, string>,
): TableCell["borders"] | undefined {
  const borders: NonNullable<TableCell["borders"]> = {};

  const left = parseTableBorder(tcPr.querySelector(":scope > lnL"), themeColors);
  const right = parseTableBorder(tcPr.querySelector(":scope > lnR"), themeColors);
  const top = parseTableBorder(tcPr.querySelector(":scope > lnT"), themeColors);
  const bottom = parseTableBorder(tcPr.querySelector(":scope > lnB"), themeColors);

  if (left) borders.left = left;
  if (right) borders.right = right;
  if (top) borders.top = top;
  if (bottom) borders.bottom = bottom;

  return Object.keys(borders).length > 0 ? borders : undefined;
}

function parseTableBorder(ln: Element | null, themeColors: Map<string, string>) {
  if (!ln || ln.querySelector(":scope > noFill")) return undefined;

  return {
    color: parseColorValue(ln, themeColors),
    width: parseOptionalInt(ln.getAttribute("w")) ? parseInt(ln.getAttribute("w") || "0", 10) / 12700 : undefined,
    style: parseTableBorderStyle(ln),
  };
}

function parseTableBorderStyle(ln: Element): string | undefined {
  const dash = ln.querySelector(":scope > prstDash")?.getAttribute("val");
  if (dash === "sysDot" || dash === "dot") return "dotted";
  if (dash === "dash" || dash === "sysDash") return "dashed";
  return "solid";
}

function parseOptionalInt(value: string | null | undefined): number | undefined {
  if (!value) return undefined;
  return parseInt(value, 10);
}

/**
 * Parse position and size from xfrm element
 */
function parsePosition(
  xfrm: Element,
  transform?: ElementTransform,
): { x: number; y: number; width: number; height: number } {
  const bounds = parseEmuBounds(xfrm);
  const transformed = applyTransform(bounds, transform);

  return {
    x: emuToPx(transformed.x),
    y: emuToPx(transformed.y),
    width: emuToPx(transformed.cx),
    height: emuToPx(transformed.cy),
  };
}

function parseEmuBounds(xfrm: Element): { x: number; y: number; cx: number; cy: number } {
  const off = xfrm.querySelector("off");
  const ext = xfrm.querySelector("ext");

  const x = parseInt(off?.getAttribute("x") || "0", 10);
  const y = parseInt(off?.getAttribute("y") || "0", 10);
  const cx = parseInt(ext?.getAttribute("cx") || "1000000", 10);
  const cy = parseInt(ext?.getAttribute("cy") || "1000000", 10);

  return {
    x,
    y,
    cx,
    cy,
  };
}

function applyTransform(
  bounds: { x: number; y: number; cx: number; cy: number },
  transform?: ElementTransform,
): { x: number; y: number; cx: number; cy: number } {
  if (!transform) {
    return bounds;
  }

  return {
    x: transform.offsetX + bounds.x * transform.scaleX,
    y: transform.offsetY + bounds.y * transform.scaleY,
    cx: bounds.cx * transform.scaleX,
    cy: bounds.cy * transform.scaleY,
  };
}

function parseGroupChildTransform(
  xfrm: Element,
  parentTransform?: ElementTransform,
): ElementTransform {
  const bounds = parseEmuBounds(xfrm);
  const transformedBounds = applyTransform(bounds, parentTransform);

  const chOff = xfrm.querySelector("chOff");
  const chExt = xfrm.querySelector("chExt");

  const childOffsetX = parseInt(chOff?.getAttribute("x") || "0", 10);
  const childOffsetY = parseInt(chOff?.getAttribute("y") || "0", 10);
  const childExtentX = parseInt(chExt?.getAttribute("cx") || String(bounds.cx || 1), 10) || 1;
  const childExtentY = parseInt(chExt?.getAttribute("cy") || String(bounds.cy || 1), 10) || 1;

  const scaleX = transformedBounds.cx / childExtentX;
  const scaleY = transformedBounds.cy / childExtentY;

  return {
    offsetX: transformedBounds.x - childOffsetX * scaleX,
    offsetY: transformedBounds.y - childOffsetY * scaleY,
    scaleX,
    scaleY,
  };
}

function emuToPx(value: number): number {
  return Math.round((value / 914400) * 96);
}

/**
 * Extract text content from txBody
 */
function extractTextContent(txBody: Element): string {
  const texts: string[] = [];
  const paragraphs = txBody.querySelectorAll("p");

  paragraphs.forEach((p) => {
    const paragraphTexts: string[] = [];
    p.querySelectorAll("t").forEach((t) => {
      paragraphTexts.push(t.textContent || "");
    });
    if (paragraphTexts.length > 0) {
      texts.push(paragraphTexts.join(""));
    }
  });

  return texts.join("\n");
}

/**
 * Parse font size
 */
function parseFontSize(txBody: Element): number | undefined {
  const defRPr = txBody.querySelector("defRPr");
  if (!defRPr) return undefined;

  const sz = defRPr.getAttribute("sz");
  if (!sz) return undefined;

  // Convert hundredths of a point to points
  return parseInt(sz, 10) / 100;
}

/**
 * Parse font family
 */
function parseFontFamily(txBody: Element): string | undefined {
  const latin = txBody.querySelector("latin");
  return latin?.getAttribute("typeface") || undefined;
}

/**
 * Check for formatting in text runs
 */
function _hasFormatting(txBody: Element, tag: string): boolean {
  // Check in runs (r) which contain the actual text formatting
  const runs = txBody.querySelectorAll("r");
  for (const run of runs) {
    const rPr = run.querySelector("rPr");
    if (rPr) {
      if (rPr.querySelector(tag)) return true;
    }
  }

  // Also check default run properties
  const defRPr = txBody.querySelector("defRPr");
  if (defRPr && defRPr.querySelector(tag)) return true;

  return false;
}

/**
 * Parse text color
 */
function parseTextColor(txBody: Element, themeColors: Map<string, string>): string | undefined {
  const solidFill = txBody.querySelector("solidFill");
  if (!solidFill) return undefined;

  return parseColorValue(solidFill, themeColors);
}
/**
 * Parse fill color from shape
 */
function parseFillColor(sp: Element, themeColors: Map<string, string>): string | undefined {
  const solidFill = sp.querySelector("spPr > solidFill");
  if (!solidFill) return undefined;

  return parseColorValue(solidFill, themeColors, { includeAlpha: true });
}

/**
 * Parse stroke color from shape
 */
function parseStrokeColor(sp: Element, themeColors: Map<string, string>): string | undefined {
  const ln = sp.querySelector("spPr > ln");
  if (!ln) return undefined;

  const solidFill = ln.querySelector("solidFill");
  if (!solidFill) return undefined;

  return parseColorValue(solidFill, themeColors, { includeAlpha: true });
}

/**
 * Parse stroke width from shape
 */
function parseStrokeWidth(sp: Element): number | undefined {
  const ln = sp.querySelector("spPr > ln");
  if (!ln) return undefined;

  const w = ln.getAttribute("w");
  if (!w) return undefined;

  // Convert EMUs to points
  return parseInt(w, 10) / 12700;
}

function parseShapeStrokeDash(sp: Element): string | undefined {
  const ln = sp.querySelector("spPr > ln");
  if (!ln || ln.querySelector(":scope > noFill")) return undefined;

  return parseThemeDash(ln);
}

function parseThemeDash(ln: Element): string | undefined {
  const dash = ln.querySelector(":scope > prstDash")?.getAttribute("val");
  if (dash === "sysDot" || dash === "dot") return "dotted";
  if (dash === "dash" || dash === "sysDash") return "dashed";
  return dash ? "solid" : undefined;
}

/**
 * Parse graphic frame element (charts, tables, SmartArt)
 */
async function parseGraphicFrame(
  graphicFrame: Element,
  zIndex: number,
  relationships: Map<string, string>,
  zip: JSZip,
  themeColors: Map<string, string>,
  transform?: ElementTransform,
): Promise<SlideElement | null> {
  const xfrm = graphicFrame.querySelector(":scope > xfrm");
  if (!xfrm) return null;

  const position = parsePosition(xfrm, transform);

  // Determine graphic type from graphicData URI
  const graphicData = graphicFrame.querySelector("graphicData");
  let graphicType: GraphicFrameElement["graphicType"] = "unknown";
  let textContent: string | undefined;
  let diagramData: DiagramData | undefined;

  if (graphicData) {
    const uri = graphicData.getAttribute("uri");
    if (uri) {
      if (uri.includes("chart")) {
        graphicType = "chart";
      } else if (uri.includes("table")) {
        const table = graphicFrame.querySelector(":scope > graphic > graphicData > tbl");
        if (table) {
          return parseTable(table, position, zIndex, themeColors);
        }
        graphicType = "table";
      } else if (uri.includes("diagram")) {
        graphicType = "diagram";
        // Extract text content from nested txBody
        textContent = extractTextFromGraphicFrame(graphicFrame);
        // Load diagram XML files
        diagramData = await loadDiagramData(graphicFrame, relationships, zip, themeColors);
      }
    }
  }

  return {
    type: "graphicFrame",
    ...position,
    zIndex,
    graphicType,
    textContent,
    diagramData,
  };
}

/**
 * Load diagram data from XML files
 */
async function loadDiagramData(
  graphicFrame: Element,
  relationships: Map<string, string>,
  zip: JSZip,
  themeColors: Map<string, string>,
): Promise<DiagramData | undefined> {
  try {
    // Find diagram references in relationships
    let diagramIndex: string | undefined;

    for (const [_relId, target] of relationships.entries()) {
      if (target.includes("diagrams/data") && target.endsWith(".xml")) {
        const match = target.match(/data(\d+)\.xml$/);
        if (match) {
          diagramIndex = match[1];
          break;
        }
      }
    }

    if (!diagramIndex) return undefined;

    // Load colors.xml for diagram-specific colors
    const colorsPath = `ppt/diagrams/colors${diagramIndex}.xml`;
    const diagramColors = await loadDiagramColors(zip, colorsPath, themeColors);

    // Load data.xml
    const dataPath = `ppt/diagrams/data${diagramIndex}.xml`;
    const dataXml = await getFileContent(zip, dataPath).catch(() => "");
    const nodes = dataXml ? parseDiagramData(dataXml) : [];

    // Load drawing.xml
    const drawingPath = `ppt/diagrams/drawing${diagramIndex}.xml`;
    const drawingXml = await getFileContent(zip, drawingPath).catch(() => "");
    const shapes = drawingXml ? parseDiagramDrawing(drawingXml, diagramColors) : [];

    if (nodes.length === 0 && shapes.length === 0) return undefined;

    return {
      nodes,
      shapes,
      category: "smartart",
    };
  } catch (error) {
    console.warn("Failed to load diagram data:", error);
    return undefined;
  }
}

/**
 * Load diagram colors from colors.xml
 */
async function loadDiagramColors(
  zip: JSZip,
  colorsPath: string,
  themeColors: Map<string, string>,
): Promise<Map<string, string>> {
  const diagramColors = new Map(themeColors); // Start with theme colors

  try {
    const colorsXml = await getFileContent(zip, colorsPath);
    const parser = new DOMParser();
    const doc = parser.parseFromString(colorsXml, "application/xml");

    // Parse color scheme from colors.xml
    const clrScheme = doc.querySelector("colorsDef");
    if (clrScheme) {
      // Colors in diagram colors file override theme colors
      // The structure varies, but typically contains styleLbl elements with color definitions
      const styleLbls = clrScheme.querySelectorAll("styleLbl");
      styleLbls.forEach((styleLbl) => {
        const name = styleLbl.getAttribute("name");
        if (name) {
          // Extract colors from this style label
          const fillClr = styleLbl.querySelector("fillClr srgbClr, fillClr schemeClr");
          if (fillClr) {
            const color = parseColorValue(fillClr, themeColors);
            if (color) {
              diagramColors.set(name, color);
            }
          }
        }
      });
    }
  } catch (error) {
    // If colors.xml doesn't exist or fails to parse, just use theme colors
    console.warn("Failed to load diagram colors:", error);
  }

  return diagramColors;
}

/**
 * Parse color value from element (srgbClr or schemeClr)
 */
function parseColorValue(
  element: Element,
  themeColors: Map<string, string>,
  options?: { includeAlpha?: boolean; colorOverrides?: Map<string, string> },
): string | undefined {
  const includeAlpha = options?.includeAlpha === true;
  const colorOverrides = options?.colorOverrides;

  // Check for srgbClr (direct RGB)
  const srgbClr = element.querySelector("srgbClr");
  if (srgbClr) {
    const val = srgbClr.getAttribute("val");
    if (val) return applyColorTransforms(`#${val}`, srgbClr, includeAlpha);
  }

  // Check for schemeClr (theme color reference)
  const schemeClr = element.querySelector("schemeClr");
  if (schemeClr) {
    const val = schemeClr.getAttribute("val");
    if (val) {
      // Look up in theme colors
      const resolved = colorOverrides?.get(val) || themeColors.get(val) || val;
      return applyColorTransforms(
        resolved.startsWith("#") ? resolved : `#${resolved}`,
        schemeClr,
        includeAlpha,
      );
    }
  }

  // Check if element itself is srgbClr or schemeClr
  const tagName = element.tagName.toLowerCase();
  if (tagName === "srgbclr" || tagName.endsWith(":srgbClr")) {
    const val = element.getAttribute("val");
    if (val) return applyColorTransforms(`#${val}`, element, includeAlpha);
  }
  if (tagName === "schemeclr" || tagName.endsWith(":schemeClr")) {
    const val = element.getAttribute("val");
    if (val) {
      const resolved = colorOverrides?.get(val) || themeColors.get(val) || val;
      return applyColorTransforms(
        resolved.startsWith("#") ? resolved : `#${resolved}`,
        element,
        includeAlpha,
      );
    }
  }

  return undefined;
}

function applyColorTransforms(baseColor: string, colorElement: Element, includeAlpha = false): string {
  const rgb = hexToRgb(baseColor);
  if (!rgb) return baseColor;

  const hsl = rgbToHsl(rgb[0], rgb[1], rgb[2]);
  if (!hsl) return baseColor;

  let [h, s, l] = hsl;

  const lumMod = parseOptionalInt(colorElement.querySelector(":scope > lumMod")?.getAttribute("val"));
  const lumOff = parseOptionalInt(colorElement.querySelector(":scope > lumOff")?.getAttribute("val"));
  const tint = parseOptionalInt(colorElement.querySelector(":scope > tint")?.getAttribute("val"));
  const shade = parseOptionalInt(colorElement.querySelector(":scope > shade")?.getAttribute("val"));

  if (lumMod !== undefined || lumOff !== undefined) {
    const mod = (lumMod ?? 100000) / 100000;
    const off = (lumOff ?? 0) / 100000;
    l = clampUnit(l * mod + off);
  }

  if (tint !== undefined) {
    const factor = tint / 100000;
    l = clampUnit(l + (1 - l) * factor);
  }

  if (shade !== undefined) {
    const factor = shade / 100000;
    l = clampUnit(l * factor);
  }

  const [r, g, b] = hslToRgb(h, s, l);

  if (includeAlpha) {
    const alpha = parseAlphaTransform(colorElement);
    if (alpha < 1) {
      return `rgba(${r}, ${g}, ${b}, ${formatAlpha(alpha)})`;
    }
  }

  return rgbToHex(r, g, b);
}

function parseAlphaTransform(colorElement: Element): number {
  const alpha = parseOptionalInt(colorElement.querySelector(":scope > alpha")?.getAttribute("val"));
  const alphaMod = parseOptionalInt(
    colorElement.querySelector(":scope > alphaMod")?.getAttribute("val"),
  );
  const alphaOff = parseOptionalInt(
    colorElement.querySelector(":scope > alphaOff")?.getAttribute("val"),
  );

  const base = (alpha ?? 100000) / 100000;
  const mod = (alphaMod ?? 100000) / 100000;
  const off = (alphaOff ?? 0) / 100000;

  return clampUnit(base * mod + off);
}

function formatAlpha(alpha: number): string {
  return (Math.round(alpha * 1000) / 1000).toString();
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
    .map((channel) => clampColor(channel).toString(16).padStart(2, "0"))
    .join("")}`;
}

function rgbToHsl(r: number, g: number, b: number): [number, number, number] | undefined {
  const rn = r / 255;
  const gn = g / 255;
  const bn = b / 255;
  const max = Math.max(rn, gn, bn);
  const min = Math.min(rn, gn, bn);
  const delta = max - min;

  let h = 0;
  let s = 0;
  const l = (max + min) / 2;

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

function clampColor(value: number): number {
  return Math.max(0, Math.min(255, value));
}

/**
 * Parse diagram data.xml
 */
function parseDiagramData(xml: string): DiagramNode[] {
  const nodes: DiagramNode[] = [];
  const parser = new DOMParser();
  const doc = parser.parseFromString(xml, "application/xml");

  // Parse point list (ptLst)
  doc.querySelectorAll("ptLst pt").forEach((pt) => {
    const id = pt.getAttribute("modelId") || "";
    const text = pt.querySelector("t")?.textContent || "";
    const level = parseInt(pt.getAttribute("lvl") || "0", 10);

    // Get parent reference from cxnFor
    const cxnFor = pt.querySelector("cxnFor");
    const parentId = cxnFor?.getAttribute("destId") || undefined;

    nodes.push({
      id,
      text,
      level,
      parentId,
    });
  });

  return nodes;
}

/**
 * Parse diagram drawing.xml with proper styling and custom geometry support
 */
function parseDiagramDrawing(xml: string, colors: Map<string, string>): DiagramShape[] {
  const shapes: DiagramShape[] = [];
  const parser = new DOMParser();
  const doc = parser.parseFromString(xml, "application/xml");

  // Parse shapes from the drawing
  doc.querySelectorAll("sp").forEach((sp, index) => {
    const xfrm = sp.querySelector("xfrm");
    if (!xfrm) return;

    // Get position and size
    const off = xfrm.querySelector("off");
    const ext = xfrm.querySelector("ext");

    if (off && ext) {
      const x = (parseInt(off.getAttribute("x") || "0", 10) / 914400) * 96;
      const y = (parseInt(off.getAttribute("y") || "0", 10) / 914400) * 96;
      const width = (parseInt(ext.getAttribute("cx") || "0", 10) / 914400) * 96;
      const height = (parseInt(ext.getAttribute("cy") || "0", 10) / 914400) * 96;

      // Get shape geometry - check for preset geometry or custom geometry
      const prstGeom = sp.querySelector("prstGeom");
      const custGeom = sp.querySelector("custGeom");
      const shapeType = prstGeom?.getAttribute("prst") || "rect";

      // Check if this is a custom geometry (icon)
      const isCustomGeom = custGeom !== null;
      const isIcon = isCustomGeom || shapeType === "icon";

      // Get shape properties
      const spPr = sp.querySelector("spPr");

      // Parse fill color with theme resolution
      let fillColor: string | undefined;
      let fillType: "solid" | "gradient" | "none" = "solid";

      if (spPr) {
        // Check for no fill
        if (spPr.querySelector("noFill")) {
          fillType = "none";
        }
        // Check for solid fill
        const solidFill = spPr.querySelector("solidFill");
        if (solidFill) {
          fillColor = parseColorValue(solidFill, colors);
        }
      }

      // Parse stroke
      const ln = spPr?.querySelector("ln");
      let strokeColor: string | undefined;
      let strokeWidth: number | undefined;

      if (ln && !ln.querySelector("noFill")) {
        const w = ln.getAttribute("w");
        if (w) strokeWidth = parseInt(w, 10) / 12700;

        const solidFill = ln.querySelector("solidFill");
        if (solidFill) {
          strokeColor = parseColorValue(solidFill, colors);
        }
      }

      // Parse text and text formatting
      const txBody = sp.querySelector("txBody");
      let text: string | undefined;
      let textColor: string | undefined;
      let fontSize: number | undefined;
      let fontFamily: string | undefined;
      let bold: boolean | undefined;
      let italic: boolean | undefined;
      let textAlign: "left" | "center" | "right" = "center";
      let textValign: "top" | "middle" | "bottom" = "middle";

      if (txBody) {
        // Get text
        text = txBody.querySelector("t")?.textContent || undefined;

        // Get body properties for alignment
        const bodyPr = txBody.querySelector("bodyPr");
        if (bodyPr) {
          const anchor = bodyPr.getAttribute("anchor");
          if (anchor === "ctr") textValign = "middle";
          else if (anchor === "b") textValign = "bottom";
          else if (anchor === "t") textValign = "top";

          const anchorCtr = bodyPr.getAttribute("anchorCtr");
          if (anchorCtr === "1") {
            textAlign = "center";
          }
        }

        // Get run properties for text formatting
        const r = txBody.querySelector("r");
        if (r) {
          const rPr = r.querySelector("rPr");
          if (rPr) {
            // Font size (in hundredths of a point)
            const sz = rPr.getAttribute("sz");
            if (sz) fontSize = parseInt(sz, 10) / 100;

            // Font family
            const latin = rPr.querySelector("latin");
            if (latin) fontFamily = latin.getAttribute("typeface") || undefined;

            // Bold/italic
            bold = rPr.getAttribute("b") === "1";
            italic = rPr.getAttribute("i") === "1";

            // Text color
            const solidFill = rPr.querySelector("solidFill");
            if (solidFill) {
              textColor = parseColorValue(solidFill, colors);
            }
          }
        }
      }

      // Parse custom geometry path data for icons
      let pathData: import("./types").ShapePath[] | undefined;
      let viewBox: { x: number; y: number; width: number; height: number } | undefined;

      if (isCustomGeom && custGeom) {
        const pathLst = custGeom.querySelector("pathLst");
        if (pathLst) {
          const avLst = custGeom.querySelector("avLst");
          const geomRect = avLst?.querySelector('gd[name="rect"]');

          // Get view box from geometry
          if (geomRect) {
            const fmla = geomRect.getAttribute("fmla");
            if (fmla) {
              // Parse view box from formula like "val 0,0,100,100"
              const match = fmla.match(/val\s+([\d.,]+)/);
              if (match) {
                const coords = match[1].split(",").map((v) => parseFloat(v.trim()));
                if (coords.length >= 4) {
                  viewBox = {
                    x: coords[0],
                    y: coords[1],
                    width: coords[2],
                    height: coords[3],
                  };
                }
              }
            }
          }

          // Parse path commands
          pathData = parsePathList(pathLst, colors);
        }
      }

      shapes.push({
        id: `shape-${index}`,
        type: isIcon ? "custom" : mapShapeType(shapeType),
        x,
        y,
        width,
        height,
        text,
        textColor,
        fontSize,
        fontFamily,
        bold,
        italic,
        textAlign,
        textValign,
        fillColor,
        fillType,
        strokeColor,
        strokeWidth,
        roundedCorners: parseRoundedCornersFromGeom(prstGeom),
        pathData,
        viewBox,
        isIcon,
        iconType: isIcon ? "custom" : undefined,
      });
    }
  });

  return shapes;
}

/**
 * Parse path list from custom geometry
 */
function parsePathList(
  pathLst: Element,
  colors: Map<string, string>,
): import("./types").ShapePath[] {
  const paths: import("./types").ShapePath[] = [];

  pathLst.querySelectorAll("path").forEach((pathEl) => {
    const commands: import("./types").PathCommand[] = [];

    for (const child of Array.from(pathEl.children)) {
      switch (child.localName) {
        case "moveTo": {
          const pt = child.querySelector(":scope > pt");
          if (!pt) break;
          commands.push({
            type: "M",
            x: parseFloat(pt.getAttribute("x") || "0"),
            y: parseFloat(pt.getAttribute("y") || "0"),
          });
          break;
        }
        case "lnTo": {
          const pt = child.querySelector(":scope > pt");
          if (!pt) break;
          commands.push({
            type: "L",
            x: parseFloat(pt.getAttribute("x") || "0"),
            y: parseFloat(pt.getAttribute("y") || "0"),
          });
          break;
        }
        case "cubicBezTo": {
          const pts = Array.from(child.querySelectorAll(":scope > pt"));
          if (pts.length < 3) break;
          commands.push({
            type: "C",
            cp1x: parseFloat(pts[0].getAttribute("x") || "0"),
            cp1y: parseFloat(pts[0].getAttribute("y") || "0"),
            cp2x: parseFloat(pts[1].getAttribute("x") || "0"),
            cp2y: parseFloat(pts[1].getAttribute("y") || "0"),
            x: parseFloat(pts[2].getAttribute("x") || "0"),
            y: parseFloat(pts[2].getAttribute("y") || "0"),
          });
          break;
        }
        case "quadBezTo": {
          const pts = Array.from(child.querySelectorAll(":scope > pt"));
          if (pts.length < 2) break;
          commands.push({
            type: "Q",
            cpx: parseFloat(pts[0].getAttribute("x") || "0"),
            cpy: parseFloat(pts[0].getAttribute("y") || "0"),
            x: parseFloat(pts[1].getAttribute("x") || "0"),
            y: parseFloat(pts[1].getAttribute("y") || "0"),
          });
          break;
        }
        case "arcTo": {
          // TODO: Implement arcTo conversion to SVG path arcs.
          break;
        }
        case "close": {
          commands.push({ type: "Z" });
          break;
        }
        default:
          break;
      }
    }

    // Parse fill/stroke for this path
    const fillColor =
      pathEl.getAttribute("fill") !== "none" ? parseColorValue(pathEl, colors) : undefined;

    paths.push({
      commands,
      fillColor: fillColor || undefined,
      strokeColor: undefined,
      strokeWidth: undefined,
    });
  });

  return paths;
}

/**
 * Map OOXML shape type to our shape type
 */
function mapShapeType(prstType: string): DiagramShape["type"] {
  switch (prstType) {
    case "ellipse":
    case "circle":
      return "ellipse";
    case "roundRect":
      return "roundedRect";
    default:
      return "rect";
  }
}

/**
 * Parse rounded corners from geometry
 */
function parseRoundedCornersFromGeom(prstGeom: Element | null): number | undefined {
  if (!prstGeom) return undefined;

  const prst = prstGeom.getAttribute("prst");
  if (prst === "roundRect") {
    const avLst = prstGeom.querySelector("avLst");
    const gd = avLst?.querySelector("gd");
    const fmla = gd?.getAttribute("fmla");
    if (fmla) {
      const match = fmla.match(/val\s+(\d+)/);
      if (match) {
        return parseInt(match[1], 10) / 100000;
      }
    }
    return 0.15;
  }
  return undefined;
}

/**
 * Extract text from graphic frame
 */
function extractTextFromGraphicFrame(graphicFrame: Element): string | undefined {
  const texts: string[] = [];
  graphicFrame.querySelectorAll("txBody p").forEach((p) => {
    const text = p.textContent?.trim();
    if (text) texts.push(text);
  });
  return texts.length > 0 ? texts.join("\n") : undefined;
}

/**
 * Get file content from ZIP
 */
async function getFileContent(zip: JSZip, path: string): Promise<string> {
  const file = zip.file(path);
  if (!file) throw new Error(`File not found: ${path}`);
  return await file.async("string");
}
