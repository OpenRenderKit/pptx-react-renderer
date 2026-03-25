# pptx-react-renderer

`pptx-react-renderer` is a browser-first library for parsing `.pptx` files and rendering them to semantic HTML.

The package is framework-agnostic, but it works naturally in React because the main API renders into a container element you control.

## Project Status

`pptx-react-renderer` is heavily under development.

It is already usable for browser-based slide previewing, but the renderer is not feature-complete and it does not aim to match PowerPoint perfectly yet. Expect rough edges, unsupported constructs, and rendering mismatches on complex decks.

## Features

- Parses `.pptx` archives client-side with `jszip`
- Renders slides to regular DOM nodes and CSS
- Supports text, images, shapes, tables, and a subset of SmartArt/graphic frames
- Exposes both parse and render APIs
- Ships as ESM and CommonJS with TypeScript types

## Install

```bash
pnpm add pptx-react-renderer
```

## Quick Start

```ts
import { renderPptx } from "pptx-react-renderer";

const response = await fetch("/deck.pptx");
const arrayBuffer = await response.arrayBuffer();

const container = document.getElementById("preview");

if (!container) {
  throw new Error("Missing preview container");
}

await renderPptx(arrayBuffer, {
  container,
  scale: 0.6,
  showSlideNumbers: true,
  theme: "light",
});
```

## React Example

```tsx
import { useEffect, useRef } from "react";
import { renderPptx } from "pptx-react-renderer";

export function PptxPreview({ file }: { file: ArrayBuffer | null }) {
  const ref = useRef<HTMLDivElement>(null);

  useEffect(() => {
    if (!file || !ref.current) return;

    void renderPptx(file, {
      container: ref.current,
      scale: 0.7,
      showSlideNumbers: true,
      theme: "light",
    });
  }, [file]);

  return <div ref={ref} />;
}
```

## API

### `renderPptx(input, options)`

Parses a `.pptx` file and immediately renders the resulting slides into `options.container`.

`input` can be an `ArrayBuffer` or `Uint8Array`.

### `parsePptx(input)`

Returns a normalized slide model without rendering:

```ts
type PptxParseResult = {
  slides: PptxSlide[];
  slideWidth: number;
  slideHeight: number;
  title?: string;
};
```

`input` can be an `ArrayBuffer` or `Uint8Array`.

### `renderSlides(slides, options)`

Renders a previously parsed slide model into a container.

### `RenderOptions`

```ts
type RenderOptions = {
  container: HTMLElement;
  scale?: number;
  showSlideNumbers?: boolean;
  theme?: "light" | "dark";
};
```

Defaults:

- `scale`: `1`
- `showSlideNumbers`: `true`
- `theme`: `"light"`

## What It Supports

- text boxes with rich text runs and paragraph-level formatting
- images extracted from the PPTX archive and inlined as data URLs
- basic shapes and fills
- tables
- grouped slide content and a subset of graphic frames / SmartArt-like data

This is not a pixel-perfect PowerPoint clone. The goal is pragmatic HTML rendering of common slide content in browser applications.

## Current Limitations

- The package is browser-first. Parsing requires `DOMParser`, `FileReader`, and `Blob`, and rendering requires `document`.
- Server-side rendering is not supported directly. Use the parser and renderer only in browser-like client runtimes.
- Grouped shapes are not fully reconstructed yet. In the current parser, group containers are simplified rather than expanded into their original child elements.
- Chart and table `graphicFrame` content is currently rendered as placeholder blocks instead of full chart/table fidelity.
- SmartArt and diagram support is partial. The parser attempts to extract diagram data and drawing shapes, but many decks still fall back to simplified node/text rendering.
- Custom geometry support is incomplete. Some path data is handled, but arc commands are currently skipped.
- Missing or unresolved embedded images fall back to a placeholder image box.
- Theme support is intentionally minimal and limited to the library's light/dark wrapper styling.
- Large decks can use significant memory because image assets are converted to data URLs during parsing.

## Runtime Notes

- Parsing requires browser APIs: `DOMParser`, `FileReader`, and `Blob`.
- Rendering additionally requires `document`.
- Server-side rendering is not supported directly.
- For React or Next.js apps, call the renderer from client-only code such as `useEffect` in a client component.
- If you run tests in Node, use a DOM-capable environment such as `jsdom`.
- Image data is currently inlined as data URLs during parsing, which can increase memory usage for large decks.

## Package Format

- ESM and CommonJS builds are published.
- TypeScript declarations are included.
- The package exports named exports only.

## Development

```bash
pnpm install
pnpm run fixtures:generate
pnpm run verify
pnpm run pack:check
pnpm run test:browser
```

To regenerate the committed sanitized golden fixture from a local source deck:

```bash
PPTX_COMPLEX_SOURCE="/absolute/path/to/source-deck.pptx" pnpm run fixtures:sanitize-complex
```

The sanitized fixture preserves slide structure, SmartArt/diagram XML, tables, groups, relationships, and media slots while replacing sensitive text, metadata, and image payloads.

## Versioning And Releases

- Follow semver.
- Create a release commit and tag in the form `vX.Y.Z`.
- The repository includes CI and a publish workflow for GitHub Actions.
- Configure npm trusted publishing before using automated publish.
- `main` is intended to stay releasable; use branches and pull requests for normal changes.
- Release tags matching `v*` are protected against rewrites and deletions.

## Testing Strategy

- unit and parser/renderer regression tests run in `vitest` with `jsdom`
- committed `.pptx` fixtures under `test/fixtures/real/` cover realistic parse and render paths
- a Playwright smoke test renders a real fixture in Chromium against the built package
- a sanitized golden deck derived from a complex local source presentation is committed so CI can exercise realistic SmartArt, tables, groups, and mixed media without depending on private files

## Contributing

See [`CONTRIBUTING.md`](https://github.com/OpenRenderKit/pptx-react-renderer/blob/main/CONTRIBUTING.md) for local setup, verification, and release steps.

## Security

See [`SECURITY.md`](https://github.com/OpenRenderKit/pptx-react-renderer/blob/main/SECURITY.md) for vulnerability reporting guidance.
