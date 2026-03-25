# pptx-react-renderer

`pptx-react-renderer` is a browser-first library for parsing `.pptx` files and rendering them to semantic HTML.

The package is framework-agnostic, but it works naturally in React because the main API renders into a container element you control.

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

### `renderPptx(arrayBuffer, options)`

Parses a `.pptx` file and immediately renders the resulting slides into `options.container`.

### `parsePptx(arrayBuffer)`

Returns a normalized slide model without rendering:

```ts
type PptxParseResult = {
  slides: PptxSlide[];
  slideWidth: number;
  slideHeight: number;
  title?: string;
};
```

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

## Runtime Notes

- Parsing requires browser APIs: `DOMParser`, `FileReader`, and `Blob`.
- Rendering additionally requires `document`.
- Server-side rendering is not supported directly.
- If you run tests in Node, use a DOM-capable environment such as `jsdom`.
- Image data is currently inlined as data URLs during parsing, which can increase memory usage for large decks.

## Package Format

- ESM and CommonJS builds are published.
- TypeScript declarations are included.
- The package exports named exports only.

## Development

```bash
pnpm install
pnpm run verify
pnpm run pack:check
```

## Versioning And Releases

- Follow semver.
- Create a release commit and tag in the form `vX.Y.Z`.
- The repository includes CI and a publish workflow for GitHub Actions.
- Configure npm trusted publishing before using automated publish.

## Contributing

See [`CONTRIBUTING.md`](https://github.com/OpenRenderKit/pptx-react-renderer/blob/main/CONTRIBUTING.md) for local setup, verification, and release steps.

## Security

See [`SECURITY.md`](https://github.com/OpenRenderKit/pptx-react-renderer/blob/main/SECURITY.md) for vulnerability reporting guidance.
