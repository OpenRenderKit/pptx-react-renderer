import type {
  PptxSlide,
  SlideElement,
  TextElement,
  TextRun,
  TextField,
  ContentElement,
  Paragraph,
  ImageElement,
  ShapeElement,
  TableElement,
  GroupElement,
  GraphicFrameElement,
  DiagramData,
  DiagramShape,
  DiagramNode,
  RenderOptions,
  SlideBackground,
} from "./types";
import { getThemeStyles as _getThemeStyles } from "./styles";

/**
 * Render PPTX slides to HTML
 */
export function renderSlides(slides: PptxSlide[], options: RenderOptions): void {
  assertDocumentAvailable();
  const { container, scale = 1, showSlideNumbers = true, theme = "light" } = options;

  // Clear container
  container.innerHTML = "";

  // Add wrapper
  const wrapper = document.createElement("div");
  wrapper.className = "pptx-renderer-wrapper";
  wrapper.style.cssText = getWrapperStyles(theme);
  container.appendChild(wrapper);

  // Render each slide
  slides.forEach((slide) => {
    const slideEl = renderSlide(slide, scale, showSlideNumbers, theme);
    wrapper.appendChild(slideEl);
  });

  // Inject styles
  injectStyles(theme);
}

function assertDocumentAvailable(): void {
  if (typeof document === "undefined") {
    throw new Error("pptx-react-renderer requires a browser document to render slides.");
  }
}

/**
 * Render a single slide
 */
function renderSlide(
  slide: PptxSlide,
  scale: number,
  showSlideNumbers: boolean,
  theme: "light" | "dark",
): HTMLElement {
  const slideEl = document.createElement("div");
  slideEl.className = "pptx-slide";
  slideEl.style.cssText = getSlideStyles(slide, scale, theme, slide.background);

  // Add slide number badge
  if (showSlideNumbers) {
    const badge = document.createElement("div");
    badge.className = "pptx-slide-number";
    badge.textContent = `Slide ${slide.number}`;
    badge.style.cssText = getSlideNumberStyles(theme);
    slideEl.appendChild(badge);
  }

  // Render elements
  slide.elements
    .sort((a, b) => a.zIndex - b.zIndex)
    .forEach((element) => {
      const el = renderElement(element, scale);
      if (el) slideEl.appendChild(el);
    });

  return slideEl;
}

/**
 * Render a single element
 */
function renderElement(element: SlideElement, scale: number): HTMLElement | null {
  switch (element.type) {
    case "text":
      return renderTextElement(element, scale);
    case "image":
      return renderImageElement(element, scale);
    case "shape":
      return renderShapeElement(element, scale);
    case "table":
      return renderTableElement(element, scale);
    case "group":
      return renderGroupElement(element, scale);
    case "graphicFrame":
      return renderGraphicFrameElement(element, scale);
    default:
      return null;
  }
}

function renderGroupElement(element: GroupElement, scale: number): HTMLElement {
  const div = document.createElement("div");
  div.className = "pptx-group-element";
  div.style.position = "absolute";
  div.style.left = "0";
  div.style.top = "0";
  div.style.width = "0";
  div.style.height = "0";
  div.style.pointerEvents = "none";

  element.children
    .slice()
    .sort((a, b) => a.zIndex - b.zIndex)
    .forEach((child) => {
      const childEl = renderElement(child, scale);
      if (childEl) div.appendChild(childEl);
    });

  return div;
}

/**
 * Render text element with rich formatting
 */
function renderTextElement(element: TextElement, scale: number): HTMLElement {
  const div = document.createElement("div");
  div.className = "pptx-text-element";

  // Position and size
  div.style.left = `${element.x * scale}px`;
  div.style.top = `${element.y * scale}px`;
  div.style.width = `${element.width * scale}px`;
  div.style.height = `${element.height * scale}px`;
  div.style.position = "absolute";

  // Always use box-sizing for consistent sizing
  div.style.boxSizing = "border-box";

  // Container styling
  div.style.fontFamily = element.defaultFontFamily || "Arial, sans-serif";
  div.style.textAlign = element.align || "left";
  div.style.color = element.defaultColor || "inherit";

  // Vertical alignment - use flexbox for proper centering
  div.style.display = "flex";
  div.style.flexDirection = "column";
  div.style.justifyContent =
    element.verticalAlign === "middle"
      ? "center"
      : element.verticalAlign === "bottom"
        ? "flex-end"
        : "flex-start";

  // Line spacing
  if (element.lineSpacing) {
    div.style.lineHeight = `${element.lineSpacing}`;
  }

  // Text wrapping - always wrap by default for proper text flow
  div.style.wordWrap = "break-word";
  div.style.overflowWrap = "break-word";
  div.style.whiteSpace = element.wrap === false ? "nowrap" : "pre-wrap";

  // Overflow behavior - clip by default to prevent text from spilling out
  const vertOverflow = element.vertOverflow || "clip";
  const _horzOverflow = element.horzOverflow || "clip";

  div.style.overflow = "hidden";

  if (vertOverflow === "ellipsis") {
    div.style.textOverflow = "ellipsis";
  }

  // Text insets (padding) - convert EMUs to pixels (1 pt ≈ 1.333 px at 96 DPI)
  const leftInset = (element.leftInset || 0) * scale * 1.333;
  const topInset = (element.topInset || 0) * scale * 1.333;
  const rightInset = (element.rightInset || 0) * scale * 1.333;
  const bottomInset = (element.bottomInset || 0) * scale * 1.333;

  if (leftInset || topInset || rightInset || bottomInset) {
    div.style.padding = `${topInset}px ${rightInset}px ${bottomInset}px ${leftInset}px`;
  }

  // Render content
  const contentDiv = document.createElement("div");
  contentDiv.style.width = "100%";
  contentDiv.style.height = "100%";
  contentDiv.style.overflow = "hidden";

  if (element.paragraphs && element.paragraphs.length > 1) {
    // Multiple paragraphs
    element.paragraphs.forEach((p) => {
      const pEl = renderParagraph(p, element, scale);
      contentDiv.appendChild(pEl);
    });
  } else if (element.runs.length > 0) {
    // Single paragraph or flat runs
    const pEl = document.createElement("p");
    pEl.style.margin = "0";
    pEl.style.padding = "0";
    pEl.style.width = "100%";

    if (element.spaceBefore) {
      pEl.style.marginTop = `${ptToPx(element.spaceBefore * scale)}px`;
    }
    if (element.spaceAfter) {
      pEl.style.marginBottom = `${ptToPx(element.spaceAfter * scale)}px`;
    }

    element.runs.forEach((run) => {
      const node = renderContentElement(run as ContentElement, element, scale);
      if (node) pEl.appendChild(node);
    });

    contentDiv.appendChild(pEl);
  }

  div.appendChild(contentDiv);

  return div;
}

/**
 * Render a paragraph
 */
function renderParagraph(p: Paragraph, element: TextElement, scale: number): HTMLElement {
  const pEl = document.createElement("p");
  pEl.style.margin = "0";
  pEl.style.padding = "0";

  // Paragraph alignment
  if (p.align) {
    pEl.style.textAlign = p.align;
  }

  // Paragraph spacing - convert pt to px for consistency
  if (p.spaceBefore) {
    pEl.style.marginTop = `${ptToPx(p.spaceBefore * scale)}px`;
  }
  if (p.spaceAfter) {
    pEl.style.marginBottom = `${ptToPx(p.spaceAfter * scale)}px`;
  }
  if (p.lineSpacing) {
    pEl.style.lineHeight = `${p.lineSpacing}`;
  }

  // Margins and indentation - convert EMUs to pixels (1 pt = 12700 EMUs, 1 pt ≈ 1.333 px)
  if (p.leftMargin !== undefined) {
    const marginPt = p.leftMargin / 12700;
    pEl.style.marginLeft = `${ptToPx(marginPt * scale)}px`;
  }
  if (p.rightMargin !== undefined) {
    const marginPt = p.rightMargin / 12700;
    pEl.style.marginRight = `${ptToPx(marginPt * scale)}px`;
  }
  if (p.indent !== undefined) {
    // Positive indent = first line indent, negative = hanging indent
    const indentPt = p.indent / 12700;
    pEl.style.textIndent = `${ptToPx(indentPt * scale)}px`;
  }

  // Level-based indentation (if no explicit margins)
  if (p.level && p.leftMargin === undefined) {
    pEl.style.marginLeft = `${p.level * 20 * scale}px`;
  }

  // Tab stops
  if (p.tabStops && p.tabStops.length > 0) {
    const tabStops = p.tabStops
      .map((tab) => {
        const pos = (tab.position / 12700) * scale;
        return `${pos}pt ${tab.type}`;
      })
      .join(", ");
    pEl.style.tabSize = tabStops;
  }

  // Handle bullets
  if (p.bullet) {
    // Check for auto-numbering
    if (p.autoNumbering) {
      // Auto-numbering is handled via CSS counters
      pEl.style.display = "list-item";
      pEl.style.listStyleType = getNumberingStyle(p.autoNumbering.type);
      pEl.style.listStylePosition = "inside";
      pEl.style.paddingLeft = `${(p.level || 0) * 0.5}em`;

      // Render runs
      p.runs.forEach((run) => {
        const node = renderContentElement(run, element, scale);
        if (node) pEl.appendChild(node);
      });
      return pEl;
    } else {
      // Custom bullet character - use a table-like layout for alignment
      pEl.style.display = "flex";
      pEl.style.alignItems = "flex-start";
      pEl.style.gap = `${0.25 * scale}em`;
      pEl.style.paddingLeft = `${(p.level || 0) * 1}em`;
      pEl.style.width = "100%";
      pEl.style.boxSizing = "border-box";

      // Create bullet span with fixed width for alignment
      const bulletSpan = document.createElement("span");
      bulletSpan.textContent = p.bulletChar || "•";
      bulletSpan.style.flexShrink = "0";
      bulletSpan.style.width = `${1 * scale}em`;
      bulletSpan.style.textAlign = "center";
      bulletSpan.style.userSelect = "none";

      // Bullet font (e.g., Wingdings)
      if (p.bulletFont) {
        bulletSpan.style.fontFamily = p.bulletFont;
      } else if (p.bulletChar === "•") {
        // Use default serif font for bullet point
        bulletSpan.style.fontFamily = "Times New Roman, serif";
      }

      // Bullet size
      if (p.bulletSize !== undefined) {
        bulletSpan.style.fontSize = `${p.bulletSize}%`;
      }

      // Bullet color
      if (p.bulletColor) {
        bulletSpan.style.color = p.bulletColor;
      }

      // Create content container
      const contentSpan = document.createElement("span");
      contentSpan.style.flex = "1";
      contentSpan.style.overflow = "hidden";
      contentSpan.style.wordWrap = "break-word";

      // Render runs into content span
      p.runs.forEach((run) => {
        const node = renderContentElement(run, element, scale);
        if (node) contentSpan.appendChild(node);
      });

      pEl.appendChild(bulletSpan);
      pEl.appendChild(contentSpan);
      return pEl;
    }
  }

  // Render runs (if not handled by bullet logic above)
  p.runs.forEach((run) => {
    const node = renderContentElement(run, element, scale);
    if (node) pEl.appendChild(node);
  });

  return pEl;
}

/**
 * Get CSS list style type from auto-numbering type
 */
function getNumberingStyle(type: "arabic" | "roman" | "alphaLc" | "alphaUc" | "ordText"): string {
  switch (type) {
    case "arabic":
      return "decimal";
    case "roman":
      return "lower-roman";
    case "alphaLc":
      return "lower-alpha";
    case "alphaUc":
      return "upper-alpha";
    case "ordText":
      return "decimal";
    default:
      return "decimal";
  }
}

/**
 * Convert points to pixels (1pt = 1.333px at 96 DPI)
 */
function ptToPx(pt: number): number {
  return pt * 1.333333;
}

/**
 * Render a single text run
 */
function renderTextRun(run: TextRun, element: TextElement, scale: number): HTMLElement {
  const span = document.createElement("span");

  // Font size (run-specific or default) - convert pt to px for consistent rendering
  const fontSizePt = run.fontSize || element.defaultFontSize || 12;
  const fontSizePx = ptToPx(fontSizePt * scale);
  span.style.fontSize = `${fontSizePx}px`;

  // Font family (run-specific or default)
  span.style.fontFamily = run.fontFamily || element.defaultFontFamily || "Arial, sans-serif";

  // Font weight - use numeric value for consistency
  if (run.bold) {
    span.style.fontWeight = "700";
  } else {
    span.style.fontWeight = "400";
  }

  // Font style
  if (run.italic) {
    span.style.fontStyle = "italic";
  }

  // Text decoration
  if (run.underline) {
    span.style.textDecoration = "underline";
  }

  // Color (run-specific or default)
  const color = run.color || element.defaultColor;
  if (color) {
    span.style.color = color;
  }

  // Highlight
  if (run.highlight) {
    span.style.backgroundColor = run.highlight;
  }

  // Baseline (superscript/subscript)
  if (run.baseline) {
    span.style.verticalAlign = run.baseline > 0 ? "super" : "sub";
    const superscriptSizePx = ptToPx(fontSizePt * 0.7 * scale);
    span.style.fontSize = `${superscriptSizePx}px`;
  }

  // Content - preserve whitespace but normalize
  span.textContent = run.text;

  return span;
}

/**
 * Render content element (run, field, or line break)
 */
function renderContentElement(
  content: ContentElement,
  element: TextElement,
  scale: number,
): HTMLElement | null {
  if (content.type === "field") {
    return renderTextField(content as TextField, element, scale);
  } else if (content.type === "break") {
    return renderLineBreak();
  } else {
    return renderTextRun(content as TextRun, element, scale);
  }
}

/**
 * Render a text field (for slide numbers, dates, etc.)
 */
function renderTextField(field: TextField, element: TextElement, scale: number): HTMLElement {
  const span = document.createElement("span");

  // Font size (field-specific or default) - convert pt to px for consistency
  const fontSizePt = field.fontSize || element.defaultFontSize || 12;
  const fontSizePx = ptToPx(fontSizePt * scale);
  span.style.fontSize = `${fontSizePx}px`;

  // Font family (field-specific or default)
  span.style.fontFamily = field.fontFamily || element.defaultFontFamily || "Arial, sans-serif";

  // Font weight - use numeric value for consistency
  if (field.bold) {
    span.style.fontWeight = "700";
  } else {
    span.style.fontWeight = "400";
  }

  // Font style
  if (field.italic) {
    span.style.fontStyle = "italic";
  }

  // Text decoration
  if (field.underline) {
    span.style.textDecoration = "underline";
  }

  // Color (field-specific or default)
  const color = field.color || element.defaultColor;
  if (color) {
    span.style.color = color;
  }

  // Content
  span.textContent = field.text;

  return span;
}

/**
 * Render a line break
 */
function renderLineBreak(): HTMLElement {
  const br = document.createElement("br");
  return br;
}

/**
 * Render image element
 */
function renderImageElement(element: ImageElement, scale: number): HTMLElement {
  const container = document.createElement("div");
  container.className = "pptx-image-element";
  container.style.left = `${element.x * scale}px`;
  container.style.top = `${element.y * scale}px`;
  container.style.width = `${element.width * scale}px`;
  container.style.height = `${element.height * scale}px`;
  container.style.position = "absolute";

  if (element.src) {
    const img = document.createElement("img");
    img.src = element.src;
    img.alt = element.alt || "";
    img.style.width = "100%";
    img.style.height = "100%";
    img.style.objectFit = "contain";
    container.appendChild(img);
  } else {
    // Placeholder for images we couldn't extract
    container.style.backgroundColor = "#f0f0f0";
    container.style.display = "flex";
    container.style.alignItems = "center";
    container.style.justifyContent = "center";
    container.style.border = "2px dashed #ccc";
    container.textContent = "📷";
    container.style.fontSize = "24px";
  }

  return container;
}

/**
 * Render shape element
 */
function renderShapeElement(element: ShapeElement, scale: number): HTMLElement {
  const div = document.createElement("div");
  div.className = "pptx-shape-element";
  div.style.left = `${element.x * scale}px`;
  div.style.top = `${element.y * scale}px`;
  div.style.width = `${element.width * scale}px`;
  div.style.height = `${element.height * scale}px`;
  div.style.position = "absolute";
  div.style.boxSizing = "border-box";

  // Handle gradient fill
  if (element.fillType === "gradient" && element.gradient) {
    const gradientStops = element.gradient.stops
      .map((s) => `${s.color} ${s.position * 100}%`)
      .join(", ");

    if (element.gradient.type === "linear") {
      const angle = element.gradient.angle || 0;
      div.style.background = `linear-gradient(${angle}deg, ${gradientStops})`;
    } else {
      div.style.background = `radial-gradient(circle, ${gradientStops})`;
    }
  } else {
    // Solid fill
    div.style.backgroundColor = element.fillColor || "transparent";
  }

  // Border styling
  if (element.strokeWidth && element.strokeWidth > 0) {
    div.style.borderWidth = `${element.strokeWidth * scale}px`;
    div.style.borderStyle = element.strokeDash || "solid";
    div.style.borderColor = element.strokeColor || "#000";
  }

  // Border radius based on shape type
  if (element.roundedCorners) {
    div.style.borderRadius = `${element.roundedCorners * scale}px`;
  } else {
    switch (element.shapeType) {
      case "ellipse":
      case "circle":
        div.style.borderRadius = "50%";
        break;
      case "roundRect":
        div.style.borderRadius = `${Math.min(element.width, element.height) * 0.1 * scale}px`;
        break;
      default:
        div.style.borderRadius = "0";
    }
  }

  return div;
}

/**
 * Render table element
 */
function renderTableElement(element: TableElement, scale: number): HTMLElement {
  const table = document.createElement("table");
  table.className = "pptx-table-element";
  table.style.left = `${element.x * scale}px`;
  table.style.top = `${element.y * scale}px`;
  table.style.width = `${element.width * scale}px`;
  table.style.height = `${element.height * scale}px`;
  table.style.position = "absolute";
  table.style.borderCollapse = "collapse";
  table.style.tableLayout = "fixed";

  if (element.backgroundColor) {
    table.style.backgroundColor = element.backgroundColor;
  }

  if (element.columnWidths?.length) {
    const colgroup = document.createElement("colgroup");
    element.columnWidths.forEach((width) => {
      const col = document.createElement("col");
      col.style.width = `${width * scale}px`;
      colgroup.appendChild(col);
    });
    table.appendChild(colgroup);
  }

  element.rows.forEach((row) => {
    const tr = document.createElement("tr");
    if (row.height) {
      tr.style.height = `${row.height * scale}px`;
    }

    row.cells.forEach((cell) => {
      const td = document.createElement("td");
      td.style.boxSizing = "border-box";
      td.style.overflow = "hidden";
      td.style.wordBreak = "break-word";

      if (cell.padding) {
        td.style.padding = `${ptToPx((cell.padding.top || 0) * scale)}px ${ptToPx(
          (cell.padding.right || 0) * scale,
        )}px ${ptToPx((cell.padding.bottom || 0) * scale)}px ${ptToPx((cell.padding.left || 0) * scale)}px`;
      } else {
        td.style.padding = "8px";
      }

      if (cell.backgroundColor) {
        td.style.backgroundColor = cell.backgroundColor;
      }

      if (cell.align) {
        td.style.textAlign = cell.align;
      }
      if (cell.verticalAlign) {
        td.style.verticalAlign =
          cell.verticalAlign === "middle"
            ? "middle"
            : cell.verticalAlign === "bottom"
              ? "bottom"
              : "top";
      }

      applyTableCellBorderStyles(td, cell);

      if (cell.rowSpan) td.rowSpan = cell.rowSpan;
      if (cell.colSpan) td.colSpan = cell.colSpan;

      const textDefaults: TextElement = {
        type: "text",
        x: 0,
        y: 0,
        width: 0,
        height: 0,
        zIndex: 0,
        runs: [],
        defaultFontSize: cell.defaultFontSize,
        defaultFontFamily: cell.defaultFontFamily,
        defaultColor: cell.defaultColor,
        align: cell.align,
        verticalAlign: cell.verticalAlign,
      };

      if (cell.paragraphs?.length) {
        cell.paragraphs.forEach((paragraph) => {
          td.appendChild(renderParagraph(paragraph, textDefaults, scale));
        });
      } else {
        td.textContent = cell.content;
      }

      tr.appendChild(td);
    });

    table.appendChild(tr);
  });

  return table;
}

function applyTableCellBorderStyles(cellEl: HTMLTableCellElement, cell: TableElement["rows"][number]["cells"][number]) {
  const borders = cell.borders;
  if (!borders) {
    cellEl.style.border = "1px solid #ccc";
    return;
  }

  applySingleBorder(cellEl, "Left", borders.left);
  applySingleBorder(cellEl, "Right", borders.right);
  applySingleBorder(cellEl, "Top", borders.top);
  applySingleBorder(cellEl, "Bottom", borders.bottom);
}

function applySingleBorder(
  element: HTMLTableCellElement,
  side: "Left" | "Right" | "Top" | "Bottom",
  border: NonNullable<TableElement["rows"][number]["cells"][number]["borders"]>[keyof NonNullable<
    TableElement["rows"][number]["cells"][number]["borders"]
  >],
) {
  const prop = `border${side}` as const;
  if (!border || !border.color || !border.width) {
    element.style[prop] = "none";
    return;
  }

  element.style[prop] = `${border.width}px ${border.style || "solid"} ${border.color}`;
}

/**
 * Render graphic frame element (charts, tables, SmartArt)
 */
function renderGraphicFrameElement(element: GraphicFrameElement, scale: number): HTMLElement {
  const div = document.createElement("div");
  div.className = "pptx-graphic-frame";
  div.style.left = `${element.x * scale}px`;
  div.style.top = `${element.y * scale}px`;
  div.style.width = `${element.width * scale}px`;
  div.style.height = `${element.height * scale}px`;
  div.style.position = "absolute";

  // Style based on graphic type
  switch (element.graphicType) {
    case "chart":
      div.style.backgroundColor = "rgba(200, 230, 255, 0.3)";
      div.style.border = "2px dashed #4a90e2";
      div.style.display = "flex";
      div.style.alignItems = "center";
      div.style.justifyContent = "center";
      div.innerHTML = '<span style="color: #4a90e2; font-size: 14px;">📊 Chart</span>';
      break;
    case "table":
      div.style.backgroundColor = "rgba(255, 230, 200, 0.3)";
      div.style.border = "2px dashed #e2964a";
      div.style.display = "flex";
      div.style.alignItems = "center";
      div.style.justifyContent = "center";
      div.innerHTML = '<span style="color: #e2964a; font-size: 14px;">📋 Table</span>';
      break;
    case "diagram":
      // Check if we have actual diagram data with shapes
      if (element.diagramData?.shapes && element.diagramData.shapes.length > 0) {
        // Render SmartArt with actual shapes
        renderSmartArtDiagram(div, element.diagramData, scale);
      } else if (element.textContent) {
        // Fallback: render text content
        div.style.backgroundColor = "rgba(230, 255, 230, 0.1)";
        div.style.border = "1px solid rgba(92, 184, 92, 0.3)";
        div.style.display = "flex";
        div.style.flexDirection = "column";
        div.style.alignItems = "flex-start";
        div.style.justifyContent = "flex-start";
        div.style.overflow = "auto";
        div.style.padding = "8px";

        // Render nodes as a simple list
        const nodes = element.diagramData?.nodes || [];
        if (nodes.length > 0) {
          nodes.forEach((node) => {
            const nodeEl = document.createElement("div");
            nodeEl.style.padding = "4px 8px";
            nodeEl.style.margin = "2px 0";
            nodeEl.style.marginLeft = `${node.level * 20}px`;
            nodeEl.style.backgroundColor = "rgba(92, 184, 92, 0.1)";
            nodeEl.style.borderRadius = "3px";
            nodeEl.style.fontSize = "12px";
            nodeEl.textContent = node.text || "";
            div.appendChild(nodeEl);
          });
        } else {
          div.innerHTML = `<div style="padding: 8px; font-size: 12px; text-align: center; color: #5cb85c;">${element.textContent.replace(/\n/g, "<br>")}</div>`;
        }
      } else {
        // Placeholder
        div.style.backgroundColor = "rgba(230, 255, 230, 0.1)";
        div.style.border = "2px dashed rgba(92, 184, 92, 0.3)";
        div.style.display = "flex";
        div.style.alignItems = "center";
        div.style.justifyContent = "center";
        div.innerHTML = '<span style="color: #5cb85c; font-size: 14px;">🔀 SmartArt</span>';
      }
      break;
    default:
      div.style.backgroundColor = "rgba(240, 240, 240, 0.5)";
      div.style.border = "2px dashed #999";
      div.style.display = "flex";
      div.style.alignItems = "center";
      div.style.justifyContent = "center";
      div.innerHTML = '<span style="color: #666; font-size: 14px;">📎 Graphic</span>';
  }

  return div;
}

/**
 * Render SmartArt diagram with shapes
 */
function renderSmartArtDiagram(
  container: HTMLElement,
  diagramData: DiagramData,
  scale: number,
): void {
  // Clear container
  container.innerHTML = "";
  container.style.position = "relative";
  container.style.overflow = "hidden";

  // Create a container for the diagram
  const diagramContainer = document.createElement("div");
  diagramContainer.style.position = "relative";
  diagramContainer.style.width = "100%";
  diagramContainer.style.height = "100%";

  // Render shapes from drawing.xml - these have exact positions and styling
  if (diagramData.shapes && diagramData.shapes.length > 0) {
    diagramData.shapes.forEach((shape) => {
      const shapeEl = renderDiagramShape(shape, scale);
      if (shapeEl) {
        diagramContainer.appendChild(shapeEl);
      }
    });
  } else if (diagramData.nodes && diagramData.nodes.length > 0) {
    // Fallback: render nodes if no shapes available
    diagramData.nodes.forEach((node, index) => {
      const nodeEl = renderDiagramNode(node, index, scale);
      if (nodeEl) {
        diagramContainer.appendChild(nodeEl);
      }
    });
  }

  container.appendChild(diagramContainer);
}

/**
 * Convert path commands to SVG path string
 */
function pathCommandsToSvgPath(commands: import("./types").PathCommand[]): string {
  return commands
    .map((cmd) => {
      switch (cmd.type) {
        case "M":
          return `M ${cmd.x} ${cmd.y}`;
        case "L":
          return `L ${cmd.x} ${cmd.y}`;
        case "C":
          return `C ${cmd.cp1x} ${cmd.cp1y}, ${cmd.cp2x} ${cmd.cp2y}, ${cmd.x} ${cmd.y}`;
        case "Q":
          return `Q ${cmd.cpx} ${cmd.cpy}, ${cmd.x} ${cmd.y}`;
        case "A":
          // Convert arc to approximate curve or simplified arc
          return `A ${cmd.rx} ${cmd.ry} ${cmd.xAxisRotation} ${cmd.largeArc ? 1 : 0} ${cmd.sweep ? 1 : 0} ${cmd.x} ${cmd.y}`;
        case "Z":
          return "Z";
        default:
          return "";
      }
    })
    .join(" ");
}

/**
 * Render a single diagram shape with proper styling and text alignment
 */
function renderDiagramShape(shape: DiagramShape, scale: number): HTMLElement | null {
  const div = document.createElement("div");
  div.style.position = "absolute";
  div.style.left = `${shape.x * scale}px`;
  div.style.top = `${shape.y * scale}px`;
  div.style.width = `${shape.width * scale}px`;
  div.style.height = `${shape.height * scale}px`;
  div.style.boxSizing = "border-box";

  // Handle custom geometry (icons with path data)
  if (shape.isIcon && shape.pathData && shape.pathData.length > 0) {
    // Create SVG for custom geometry
    const svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
    svg.setAttribute("width", "100%");
    svg.setAttribute("height", "100%");
    svg.setAttribute(
      "viewBox",
      `${shape.viewBox?.x || 0} ${shape.viewBox?.y || 0} ${shape.viewBox?.width || 100} ${shape.viewBox?.height || 100}`,
    );
    svg.style.display = "block";
    svg.style.overflow = "visible";

    // Render each path
    shape.pathData.forEach((pathData) => {
      const path = document.createElementNS("http://www.w3.org/2000/svg", "path");
      const d = pathCommandsToSvgPath(pathData.commands);
      path.setAttribute("d", d);

      // Apply fill
      if (pathData.fillColor) {
        path.setAttribute("fill", pathData.fillColor);
      } else if (shape.fillColor) {
        path.setAttribute("fill", shape.fillColor);
      } else {
        path.setAttribute("fill", "currentColor");
      }

      // Apply stroke
      if (pathData.strokeColor || shape.strokeColor) {
        path.setAttribute("stroke", pathData.strokeColor || shape.strokeColor || "none");
        path.setAttribute(
          "stroke-width",
          String((pathData.strokeWidth || shape.strokeWidth || 1) * scale),
        );
        path.setAttribute("stroke-linecap", "round");
        path.setAttribute("stroke-linejoin", "round");
      }

      svg.appendChild(path);
    });

    div.appendChild(svg);

    // Add text if present
    if (shape.text) {
      const textDiv = document.createElement("div");
      textDiv.style.position = "absolute";
      textDiv.style.top = "0";
      textDiv.style.left = "0";
      textDiv.style.width = "100%";
      textDiv.style.height = "100%";
      textDiv.style.display = "flex";
      textDiv.style.alignItems =
        shape.textValign === "top"
          ? "flex-start"
          : shape.textValign === "bottom"
            ? "flex-end"
            : "center";
      textDiv.style.justifyContent =
        shape.textAlign === "left"
          ? "flex-start"
          : shape.textAlign === "right"
            ? "flex-end"
            : "center";
      textDiv.style.textAlign = shape.textAlign || "center";
      textDiv.style.padding = `${4 * scale}px`;
      textDiv.style.boxSizing = "border-box";

      const fontSizePt = shape.fontSize || 12;
      const fontSizePx = ptToPx(fontSizePt * scale);
      textDiv.style.fontSize = `${fontSizePx}px`;
      textDiv.style.color = shape.textColor || "#333";
      if (shape.bold) textDiv.style.fontWeight = "700";
      if (shape.italic) textDiv.style.fontStyle = "italic";
      if (shape.fontFamily) textDiv.style.fontFamily = shape.fontFamily;

      textDiv.textContent = shape.text;
      div.appendChild(textDiv);
    }

    return div;
  }

  // Shape styling - rounded corners
  switch (shape.type) {
    case "ellipse":
      div.style.borderRadius = "50%";
      break;
    case "roundedRect":
      // Use the roundedCorners percentage or default to 8px
      const radius = shape.roundedCorners
        ? `${Math.min(shape.width, shape.height) * shape.roundedCorners * scale}px`
        : "8px";
      div.style.borderRadius = radius;
      break;
    case "custom":
      break;
    default:
      div.style.borderRadius = "0";
  }

  // Fill color - use actual color from shape or transparent
  if (shape.fillType === "none") {
    div.style.backgroundColor = "transparent";
  } else if (shape.fillColor) {
    div.style.backgroundColor = shape.fillColor;
  } else {
    div.style.backgroundColor = "transparent";
  }

  // Border/stroke
  if (shape.strokeColor && shape.strokeWidth !== undefined) {
    div.style.border = `${shape.strokeWidth * scale}px solid ${shape.strokeColor}`;
  } else if (shape.strokeColor) {
    div.style.border = `1px solid ${shape.strokeColor}`;
  }

  // Text rendering with proper alignment and overflow handling
  if (shape.text) {
    div.style.display = "flex";

    // Vertical alignment
    switch (shape.textValign) {
      case "top":
        div.style.alignItems = "flex-start";
        break;
      case "bottom":
        div.style.alignItems = "flex-end";
        break;
      case "middle":
      default:
        div.style.alignItems = "center";
    }

    // Horizontal alignment
    switch (shape.textAlign) {
      case "left":
        div.style.justifyContent = "flex-start";
        div.style.textAlign = "left";
        break;
      case "right":
        div.style.justifyContent = "flex-end";
        div.style.textAlign = "right";
        break;
      case "center":
      default:
        div.style.justifyContent = "center";
        div.style.textAlign = "center";
    }

    // Text styling - convert pt to px for consistency
    const fontSizePt = shape.fontSize || 12;
    const fontSizePx = ptToPx(fontSizePt * scale);
    div.style.fontSize = `${fontSizePx}px`;
    div.style.color = shape.textColor || "#333";
    div.style.lineHeight = "1.3";

    // Font weight and style
    if (shape.bold) div.style.fontWeight = "700";
    if (shape.italic) div.style.fontStyle = "italic";
    if (shape.fontFamily) div.style.fontFamily = shape.fontFamily;

    // Create content container for proper overflow handling
    const contentDiv = document.createElement("div");
    contentDiv.style.width = "100%";
    contentDiv.style.height = "100%";
    contentDiv.style.padding = `${4 * scale}px ${8 * scale}px`;
    contentDiv.style.boxSizing = "border-box";
    contentDiv.style.overflow = "hidden";
    contentDiv.style.display = "flex";
    contentDiv.style.flexDirection = "column";
    contentDiv.style.justifyContent = div.style.alignItems.includes("center")
      ? "center"
      : div.style.alignItems.includes("flex-end")
        ? "flex-end"
        : "flex-start";
    contentDiv.style.alignItems =
      shape.textAlign === "center"
        ? "center"
        : shape.textAlign === "right"
          ? "flex-end"
          : "flex-start";
    contentDiv.style.textAlign = shape.textAlign || "center";

    // Create text span with better overflow handling
    const textSpan = document.createElement("span");
    textSpan.textContent = shape.text;
    textSpan.style.maxWidth = "100%";
    textSpan.style.wordWrap = "break-word";
    textSpan.style.overflowWrap = "break-word";
    textSpan.style.whiteSpace = "pre-wrap";
    textSpan.style.overflow = "hidden";
    textSpan.style.textOverflow = "clip";
    textSpan.style.display = "block";

    contentDiv.appendChild(textSpan);
    div.appendChild(contentDiv);
  }

  return div;
}

/**
 * Render a diagram node (fallback for when no shapes are available)
 */
function renderDiagramNode(node: DiagramNode, index: number, scale: number): HTMLElement | null {
  const div = document.createElement("div");
  div.style.position = "absolute";

  // Position based on hierarchy level
  const x = 10 + node.level * 20;
  const y = 10 + index * 40;
  div.style.left = `${x * scale}px`;
  div.style.top = `${y * scale}px`;
  div.style.minWidth = `${100 * scale}px`;
  div.style.padding = `${8 * scale}px`;

  div.style.backgroundColor = "#e8f5e9";
  div.style.border = "1px solid #4caf50";
  div.style.borderRadius = "4px";
  div.style.fontSize = `${12 * scale}px`;
  div.style.color = "#333";

  if (node.text) {
    div.textContent = node.text;
  }

  return div;
}

/**
 * Get wrapper styles
 */
function getWrapperStyles(theme: "light" | "dark"): string {
  const bg = theme === "dark" ? "#1a1a1a" : "#f5f5f5";
  return `
    background: ${bg};
    padding: 40px 20px;
    min-height: 100%;
  `;
}

/**
 * Get slide styles
 */
function getSlideStyles(
  slide: PptxSlide,
  scale: number,
  theme: "light" | "dark",
  background?: SlideBackground,
): string {
  const defaultBg = theme === "dark" ? "#2a2a2a" : "#ffffff";
  const shadow = theme === "dark" ? "0 4px 20px rgba(0,0,0,0.5)" : "0 4px 20px rgba(0,0,0,0.1)";

  // Build background style
  let backgroundStyle = "";
  if (background) {
    if (background.type === "solid" && background.color) {
      backgroundStyle = `background-color: ${background.color};`;
    } else if (background.type === "image" && background.imageSrc) {
      backgroundStyle = `background-image: url(${background.imageSrc}); background-size: cover; background-position: center;`;
    }
  }

  const bg = backgroundStyle || `background-color: ${defaultBg};`;

  return `
    position: relative;
    width: ${slide.width * scale}px;
    height: ${slide.height * scale}px;
    margin: 0 auto 40px auto;
    box-shadow: ${shadow};
    overflow: hidden;
    ${bg}
  `;
}

/**
 * Get slide number badge styles
 */
function getSlideNumberStyles(theme: "light" | "dark"): string {
  const bg = theme === "dark" ? "#3a3a3a" : "#f0f0f0";
  const color = theme === "dark" ? "#fff" : "#333";

  return `
    position: absolute;
    top: 12px;
    left: 12px;
    background: ${bg};
    color: ${color};
    padding: 6px 12px;
    border-radius: 4px;
    font-size: 12px;
    font-weight: 600;
    z-index: 1000;
    box-shadow: 0 2px 4px rgba(0,0,0,0.1);
  `;
}

/**
 * Inject CSS styles into document
 */
function injectStyles(theme: "light" | "dark"): void {
  const styleId = "pptx-renderer-styles";
  if (document.getElementById(styleId)) return;

  const style = document.createElement("style");
  style.id = styleId;
  style.textContent = getThemeStyles(theme);
  document.head.appendChild(style);
}

/**
 * Get theme-specific CSS
 */
function getThemeStyles(_theme: "light" | "dark"): string {
  return `
    .pptx-renderer-wrapper {
      font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Oxygen, Ubuntu, Cantarell, sans-serif;
    }

    .pptx-slide {
      page-break-inside: avoid;
      break-inside: avoid;
      background-repeat: no-repeat;
      background-size: cover;
      background-position: center;
    }

    .pptx-text-element {
      word-wrap: break-word;
      overflow-wrap: break-word;
      overflow: hidden;
      line-height: 1.3;
      box-sizing: border-box;
    }

    .pptx-text-element p {
      margin: 0;
      padding: 0;
      width: 100%;
      box-sizing: border-box;
    }

    .pptx-text-element span {
      display: inline;
      white-space: pre-wrap;
      word-wrap: break-word;
      overflow-wrap: break-word;
    }

    /* Ensure text doesn't overflow its container */
    .pptx-text-element > div {
      width: 100%;
      height: 100%;
      overflow: hidden;
    }

    .pptx-shape-element {
      box-sizing: border-box;
    }

    .pptx-image-element {
      overflow: hidden;
    }

    .pptx-image-element img {
      display: block;
      width: 100%;
      height: 100%;
      object-fit: contain;
    }

    .pptx-table-element {
      font-size: 14px;
      border-collapse: collapse;
    }

    .pptx-table-element td {
      padding: 8px;
      border: 1px solid #ccc;
    }

    /* SmartArt diagram styles */
    .pptx-graphic-frame {
      box-sizing: border-box;
    }

    /* Print styles */
    @media print {
      .pptx-slide {
        page-break-after: always;
        box-shadow: none !important;
        margin: 0 !important;
      }
      
      .pptx-text-element {
        overflow: visible !important;
      }
    }
  `;
}
