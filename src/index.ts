/**
 * PPTX HTML Renderer
 *
 * A lightweight, client-side PowerPoint to HTML renderer.
 *
 * @example
 * ```typescript
 * import { renderPptx } from 'pptx-react-renderer';
 *
 * const arrayBuffer = await fetch('presentation.pptx').then(r => r.arrayBuffer());
 * const container = document.getElementById('preview');
 *
 * await renderPptx(arrayBuffer, {
 *   container,
 *   scale: 0.5,
 *   showSlideNumbers: true,
 *   theme: 'light'
 * });
 * ```
 */

export { parsePptx } from "./parser";
export { renderSlides } from "./renderer";
export type {
  PptxSlide,
  SlideElement,
  BaseElement,
  TextElement,
  TextShapeFrame,
  TextRun,
  TextField,
  LineBreak,
  ContentElement,
  Paragraph,
  ImageElement,
  ShapeElement,
  TableElement,
  TableRow,
  TableCell,
  GroupElement,
  GraphicFrameElement,
  DiagramData,
  DiagramNode,
  DiagramShape,
  DiagramConnection,
  TabStop,
  PptxParseResult,
  SlideBackground,
  RenderOptions,
  ThemeSchemeColors,
  ThemePalette,
} from "./types";

import { parsePptx } from "./parser";
import { renderSlides } from "./renderer";
import type { RenderOptions } from "./types";

/**
 * Main render function - parses PPTX and renders to HTML
 */
export async function renderPptx(
  arrayBuffer: ArrayBuffer | Uint8Array,
  options: RenderOptions,
): Promise<void> {
  const { slides } = await parsePptx(arrayBuffer);
  renderSlides(slides, options);
}
