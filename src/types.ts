/**
 * PPTX HTML Renderer
 *
 * A lightweight, client-side PowerPoint to HTML renderer.
 * Converts .pptx files to semantic HTML with CSS styling.
 *
 * @license MIT
 */

// Slide content types
export interface PptxSlide {
  number: number;
  width: number;
  height: number;
  elements: SlideElement[];
  background?: SlideBackground;
}

export type SlideElement =
  | TextElement
  | ImageElement
  | ShapeElement
  | TableElement
  | GroupElement
  | GraphicFrameElement;

export interface BaseElement {
  type: string;
  x: number;
  y: number;
  width: number;
  height: number;
  rotation?: number;
  zIndex: number;
}

// Rich text run with individual formatting
export interface TextRun {
  type?: "run"; // discriminant for ContentElement union
  text: string;
  fontSize?: number; // in points
  fontFamily?: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  color?: string; // hex color
  highlight?: string; // highlight color
  baseline?: number; // superscript/subscript (percentage)
}

// Text field element (for slide numbers, dates, etc.)
export interface TextField {
  type: "field";
  text: string;
  fieldType?: string; // e.g., 'slidenum', 'datetime'
  fontSize?: number;
  fontFamily?: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  color?: string;
}

// Line break element
export interface LineBreak {
  type: "break";
}

// Content element can be a run, field, or line break
export type ContentElement = TextRun | TextField | LineBreak;

export interface TextElement extends BaseElement {
  type: "text";
  runs: TextRun[]; // Array of text runs with individual formatting
  paragraphs?: Paragraph[]; // Support for multiple paragraphs
  frame?: TextShapeFrame; // Visual frame styling for text-bearing shapes
  defaultFontSize?: number;
  defaultFontFamily?: string;
  defaultColor?: string;
  align?: "left" | "center" | "right" | "justify";
  verticalAlign?: "top" | "middle" | "bottom";
  lineSpacing?: number; // line spacing multiplier
  spaceBefore?: number; // space before paragraph in points
  spaceAfter?: number; // space after paragraph in points
  indent?: number; // left indent in points
  wrap?: boolean; // text wrapping enabled
  vertOverflow?: "overflow" | "clip" | "ellipsis"; // vertical overflow behavior
  horzOverflow?: "overflow" | "clip"; // horizontal overflow behavior
  leftInset?: number; // left margin inset in points
  topInset?: number; // top margin inset in points
  rightInset?: number; // right margin inset in points
  bottomInset?: number; // bottom margin inset in points
}

export interface TextShapeFrame {
  shapeType: string;
  fillColor?: string;
  strokeColor?: string;
  strokeWidth?: number;
  strokeDash?: string;
  roundedCorners?: number;
}

export interface Paragraph {
  runs: ContentElement[];
  align?: "left" | "center" | "right" | "justify";
  // Bullet formatting
  bullet?: boolean;
  bulletChar?: string;
  bulletFont?: string; // Font for bullet character (e.g., 'Wingdings')
  bulletSize?: number; // Bullet size in percentage of text (100 = same size)
  bulletColor?: string; // Bullet color
  autoNumbering?: {
    type: "arabic" | "roman" | "alphaLc" | "alphaUc" | "ordText";
    startAt?: number;
  };
  // Indentation and spacing
  level?: number;
  leftMargin?: number; // Left margin in EMUs
  rightMargin?: number; // Right margin in EMUs
  indent?: number; // First line indent in EMUs (hanging indent if negative)
  spaceBefore?: number;
  spaceAfter?: number;
  lineSpacing?: number;
  // Tab stops
  tabStops?: TabStop[];
  // Default run properties for this paragraph
  defaultRunProps?: TextRun;
}

export interface TabStop {
  position: number; // Position in EMUs
  type: "left" | "right" | "center" | "decimal";
}

export interface ImageElement extends BaseElement {
  type: "image";
  src: string; // base64 data URL
  mimeType: string;
  alt?: string;
}

export interface ShapeElement extends BaseElement {
  type: "shape";
  shapeType: string;
  fillColor?: string;
  fillType?: "solid" | "gradient";
  gradient?: GradientFill;
  strokeColor?: string;
  strokeWidth?: number;
  strokeDash?: string; // dashed, dotted, etc.
  roundedCorners?: number; // border radius
  pathData?: ShapePath[];
  viewBox?: { x: number; y: number; width: number; height: number };
}

export interface GradientFill {
  type: "linear" | "radial";
  angle?: number;
  stops: GradientStop[];
}

export interface GradientStop {
  position: number; // 0-1
  color: string;
}

// Shape path data for custom geometry
export interface ShapePath {
  commands: PathCommand[];
  fillColor?: string;
  strokeColor?: string;
  strokeWidth?: number;
}

// SVG path command
export type PathCommand =
  | { type: "M"; x: number; y: number } // Move to
  | { type: "L"; x: number; y: number } // Line to
  | { type: "C"; cp1x: number; cp1y: number; cp2x: number; cp2y: number; x: number; y: number } // Cubic bezier
  | { type: "Q"; cpx: number; cpy: number; x: number; y: number } // Quadratic bezier
  | {
      type: "A";
      rx: number;
      ry: number;
      xAxisRotation: number;
      largeArc: boolean;
      sweep: boolean;
      x: number;
      y: number;
    } // Arc
  | { type: "Z" }; // Close path

export interface TableElement extends BaseElement {
  type: "table";
  rows: TableRow[];
  columnWidths?: number[];
  backgroundColor?: string;
}

export interface TableRow {
  cells: TableCell[];
  height?: number;
}

export interface TableBorder {
  color?: string;
  width?: number;
  style?: string;
}

export interface TableCellPadding {
  left?: number;
  right?: number;
  top?: number;
  bottom?: number;
}

export interface TableCell {
  content: string;
  paragraphs?: Paragraph[];
  defaultFontSize?: number;
  defaultFontFamily?: string;
  defaultColor?: string;
  align?: "left" | "center" | "right" | "justify";
  verticalAlign?: "top" | "middle" | "bottom";
  padding?: TableCellPadding;
  rowSpan?: number;
  colSpan?: number;
  backgroundColor?: string;
  borders?: {
    left?: TableBorder;
    right?: TableBorder;
    top?: TableBorder;
    bottom?: TableBorder;
  };
}

export interface GroupElement extends BaseElement {
  type: "group";
  children: SlideElement[];
}

export interface GraphicFrameElement extends BaseElement {
  type: "graphicFrame";
  graphicType?: "chart" | "table" | "diagram" | "unknown";
  // For diagrams/SmartArt, we store the extracted text content
  textContent?: string;
  // SmartArt diagram data
  diagramData?: DiagramData;
}

// SmartArt diagram structure
export interface DiagramData {
  // Nodes from data.xml
  nodes: DiagramNode[];
  // Connections between nodes
  connections?: DiagramConnection[];
  // Rendered shapes from drawing.xml
  shapes?: DiagramShape[];
  // Diagram type/category
  category?: string;
  // Color scheme
  colors?: DiagramColors;
}

export interface DiagramNode {
  id: string;
  text: string;
  level: number;
  parentId?: string;
  childrenIds?: string[];
  // Visual properties
  x?: number;
  y?: number;
  width?: number;
  height?: number;
  shapeType?: string;
  fillColor?: string;
  strokeColor?: string;
}

export interface DiagramConnection {
  fromId: string;
  toId: string;
  type?: "line" | "arrow" | "elbow" | "curve";
}

export interface DiagramShape {
  id: string;
  type: "rect" | "ellipse" | "roundedRect" | "custom";
  x: number;
  y: number;
  width: number;
  height: number;
  // Fill styling
  fillColor?: string;
  fillType?: "solid" | "gradient" | "none";
  gradient?: GradientFill;
  // Stroke/border styling
  strokeColor?: string;
  strokeWidth?: number;
  strokeDash?: string;
  // Text styling
  text?: string;
  textColor?: string;
  fontSize?: number;
  fontFamily?: string;
  bold?: boolean;
  italic?: boolean;
  // Text alignment
  textAlign?: "left" | "center" | "right";
  textValign?: "top" | "middle" | "bottom";
  // Visual effects
  shadow?: boolean;
  roundedCorners?: number;
  // Custom geometry for icons/paths
  pathData?: ShapePath[];
  viewBox?: { x: number; y: number; width: number; height: number };
  // Icon properties
  isIcon?: boolean;
  iconType?: "custom" | "prst";
}

export interface DiagramColors {
  // Theme accent colors (accent1 - accent6)
  accent1?: string;
  accent2?: string;
  accent3?: string;
  accent4?: string;
  accent5?: string;
  accent6?: string;
  // Other theme colors
  dk1?: string;
  lt1?: string;
  dk2?: string;
  lt2?: string;
  hlink?: string;
  folHlink?: string;
  // Background colors
  bg1?: string;
  bg2?: string;
}

// Theme color mapping
export interface ThemeSchemeColors {
  scheme: {
    dk1?: string; // Dark 1
    lt1?: string; // Light 1
    dk2?: string; // Dark 2
    lt2?: string; // Light 2
    accent1?: string;
    accent2?: string;
    accent3?: string;
    accent4?: string;
    accent5?: string;
    accent6?: string;
    hlink?: string;
    folHlink?: string;
  };
}

// Parser result
export interface PptxParseResult {
  slides: PptxSlide[];
  slideWidth: number;
  slideHeight: number;
  title?: string;
}

// Slide background
export interface SlideBackground {
  type: "solid" | "image";
  color?: string;
  imageSrc?: string;
}

// Renderer options
export interface RenderOptions {
  container: HTMLElement;
  scale?: number;
  showSlideNumbers?: boolean;
  theme?: "light" | "dark";
}

// CSS variables for theming
export interface ThemePalette {
  background: string;
  text: string;
  border: string;
  slideBackground: string;
  accent: string;
}
