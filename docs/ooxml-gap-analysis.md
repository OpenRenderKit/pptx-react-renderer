# OOXML Gap Analysis

This document maps the current `pptx-react-renderer` parser and renderer against the OOXML features that most directly affect PowerPoint fidelity.

The target is not broad "PPTX support". The target is visual fidelity: text, SmartArt, images, placement, color, and layout that are as close to PowerPoint as practical.

## Primary Sources

- [Open XML SDK API reference](https://learn.microsoft.com/en-us/dotnet/api/overview/openxml/)
- [GroupShapeProperties Class](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.groupshapeproperties?view=openxml-3.0.1)
- [GroupShapeProperties.TransformGroup Property](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.groupshapeproperties.transformgroup?view=openxml-3.0.1)
- [GraphicFrame Class](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.graphicframe?view=openxml-3.0.1)
- [BodyProperties Class](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.bodyproperties?view=openxml-3.0.1)
- [Paragraph.ParagraphProperties Property](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.paragraph.paragraphproperties?view=openxml-3.0.1)
- [BlipFill Class](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.blipfill?view=openxml-3.0.1)
- [GradientFill Class](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.gradientfill?view=openxml-3.0.1)
- [SchemeColor Class](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.schemecolor?view=openxml-3.0.1)
- [Table Class](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.table?view=openxml-3.0.1)
- [TableCell.TableCellProperties Property](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.tablecell.tablecellproperties?view=openxml-3.0.1)
- [Table Cell Row Span Property](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.tablecell.rowspan?view=openxml-3.0.1)

## What The Current Library Already Does

- Parses slide dimensions and basic element trees from PresentationML.
- Parses text runs, multiple paragraphs, line breaks, text fields, bullet characters, and basic paragraph spacing.
- Extracts embedded images and renders them as data URLs.
- Parses basic shapes and some SmartArt drawing shapes.
- Loads theme color entries from `ppt/theme/theme1.xml`.
- Parses a sanitized complex regression deck and renders it in browser CI.

## Fidelity Gaps By OOXML Domain

### 1. Group Shapes and Nested Transform Math

Relevant OOXML:

- `p:grpSpPr` can define common group properties.
- `a:xfrm` on groups is a "2D Transform for Grouped Objects".
- Group children are positioned in child coordinates and remapped by `off` / `ext` and `chOff` / `chExt`.

Current state:

- Prior to the current branch work, [`src/parser.ts`](/Users/sid/Documents/Coding/pptx-react-renderer/src/parser.ts) reduced `grpSp` to a plain rectangle.
- There is no support yet for group-level fill inheritance, rotation, flip state, clipping, or effect inheritance.
- `cxnSp`, `alternateContent`, and other non-`sp` / `pic` / `graphicFrame` grouped children are still not handled.

Impact:

- Complex SmartArt and grouped infographic layouts lose structure and precise placement.
- Nested transforms are a major source of visible layout drift.

Status:

- This is the first implementation track and is being addressed on this branch.

### 2. Text Layout Fidelity

Relevant OOXML:

- `a:bodyPr` controls text box insets, wrapping, vertical overflow, horizontal overflow, vertical text, anchoring, and more.
- `a:pPr` controls paragraph alignment, bulleting, margins, indentation, spacing, tabs, and default run properties.
- `a:rPr` and related run properties control font family, size, emphasis, underline, highlighting, baseline shifts, and color.

Current state:

- The parser handles many core text features already.
- The renderer does not measure text like PowerPoint and relies on browser flow and simplified CSS.
- Font fallback is browser-dependent rather than theme-aware.
- Tabs, hanging indents, auto-numbering fidelity, East Asian fonts, vertical text, text direction, and richer overflow behavior are still incomplete.
- Theme-driven font resolution is effectively absent.

Impact:

- Text is readable, but line breaks, text wrapping, and alignment often drift from PowerPoint.
- This is one of the biggest blockers for real 1:1 fidelity.

### 3. Theme and Color Resolution

Relevant OOXML:

- `a:schemeClr` binds colors to theme entries.
- `a:schemeClr` may carry transforms such as `tint`, `shade`, `lumMod`, and `lumOff`.
- OOXML theme styles also affect fill refs, line refs, effect refs, and font refs.

Current state:

- [`src/parser.ts`](/Users/sid/Documents/Coding/pptx-react-renderer/src/parser.ts) loads base theme colors.
- Color transforms from `schemeClr` are not properly resolved.
- Style matrix references like `fillRef`, `lnRef`, and `fontRef` are not modeled.
- Group-level and table-level style inheritance is shallow.

Impact:

- Colors are often approximately correct but not truly theme-faithful.
- SmartArt and shape styles drift heavily when they rely on theme transforms rather than explicit RGB.

### 4. Graphic Frames

Relevant OOXML:

- `p:graphicFrame` is a container for externally generated graphics.
- `a:graphicData` can point to charts, tables, diagrams, and other payloads.

Current state:

- `graphicFrame` type detection exists in [`src/parser.ts`](/Users/sid/Documents/Coding/pptx-react-renderer/src/parser.ts).
- Tables inside graphic frames are not rendered as real tables.
- Charts are placeholders in [`src/renderer.ts`](/Users/sid/Documents/Coding/pptx-react-renderer/src/renderer.ts).
- SmartArt diagrams partially use `drawing*.xml` shapes but still fall back often.

Impact:

- A large class of business decks still renders as placeholder boxes or simplified diagrams.

### 5. SmartArt / DiagramML

Relevant OOXML:

- SmartArt spans multiple diagram parts such as `data*.xml`, `drawing*.xml`, and `colors*.xml`.
- True fidelity depends on combining node data, drawing geometry, style/color mapping, and layout transforms.

Current state:

- The parser loads diagram data, drawing parts, and some color information.
- The renderer can display extracted diagram shapes or a simplified node fallback.
- Layout/category semantics are not fully interpreted.
- Connector routing, text fitting, style matrix application, and some geometry are still incomplete.

Impact:

- This is the largest fidelity gap after text layout.
- Sanitized complex fixture slides already show this gap clearly.

### 6. Tables

Relevant OOXML:

- DrawingML tables contain `tblPr`, `tblGrid`, `tr`, `tc`, and `tcPr`.
- Table cells can span rows and columns and carry fills, borders, and margins.

Current state:

- [`src/parser.ts`](/Users/sid/Documents/Coding/pptx-react-renderer/src/parser.ts) currently extracts mostly plain text content from table cells.
- Table widths, heights, margins, borders, and styles are not modeled well.
- Embedded tables inside `graphicFrame` still render as placeholders.

Impact:

- Evaluation matrices and comparison tables from real decks lose both styling and layout fidelity.

### 7. Images and Picture Fills

Relevant OOXML:

- `a:blipFill` can specify embedded images, source cropping, tiling, stretching, and rotation behavior.

Current state:

- Embedded images are extracted and rendered.
- Cropping via `srcRect`, tile modes, and image fill behavior are not modeled.
- Shape picture fills and background picture fills are only partially supported.

Impact:

- Photos often appear, but their crop and fit differ from PowerPoint.

### 8. Shape Geometry and Stroke / Fill Effects

Relevant OOXML:

- Shapes can use preset or custom geometry.
- Custom geometry can include arcs, bezier curves, guides, and adjustable values.
- Fill and line styling include gradients, patterns, transparency, dashes, caps, joins, and effects.

Current state:

- Basic preset shapes render.
- Some custom geometry parsing exists for diagram shapes.
- `arcTo` is currently skipped in [`src/parser.ts`](/Users/sid/Documents/Coding/pptx-react-renderer/src/parser.ts).
- Line dash patterns, caps, joins, gradient fidelity, and many effects are not implemented.

Impact:

- Icons and more complex branded shapes still drift visually.

### 9. Slide Backgrounds and Master / Layout Inheritance

Relevant OOXML:

- Slide appearance can inherit from masters, layouts, theme style matrices, and background properties.

Current state:

- Per-slide solid and some image backgrounds are parsed.
- Layout/master inheritance is minimal.
- Theme format scheme usage is shallow.

Impact:

- Branded decks can render with the wrong background, font, or accent behavior.

## Current Priority Order

1. Group transform fidelity and nested content preservation.
2. Graphic-frame tables and chart model extraction.
3. Theme color transforms and style reference resolution.
4. Text layout and font/theme fidelity.
5. SmartArt layout fidelity beyond fallback shape rendering.
6. Picture cropping and shape fills.
7. Custom geometry and stroke/fill effects.
8. Master/layout inheritance.

## Definition Of Done For “1:1 Enough”

The library is not done when it "parses most decks". It is done when a realistic presentation:

- preserves the right element tree and stacking order
- places grouped and nested content in the right coordinates
- matches text wrapping, alignment, margins, and sizing closely
- resolves theme colors and style references accurately
- renders SmartArt and tables without placeholder fallbacks
- crops and scales images like PowerPoint
- handles the real complex fixture without obvious visual drift

## Branch Note

The first implementation pass on this branch focuses on `grpSp` fidelity because it is the clearest structural prerequisite for the rest of the roadmap.
