import { renderPptx } from "../../../dist/index.js";

async function main() {
  const container = document.getElementById("preview");
  (window as Window & { __PPTX_RENDER_STATUS__?: unknown }).__PPTX_RENDER_STATUS__ = {
    state: "loading",
  };

  if (!container) {
    throw new Error("Missing preview container");
  }

  try {
    const response = await fetch("/test/fixtures/real/complex-sanitized.pptx");
    const arrayBuffer = await response.arrayBuffer();

    await renderPptx(arrayBuffer, {
      container,
      scale: 0.6,
      showSlideNumbers: true,
      theme: "light",
    });

    (window as Window & { __PPTX_RENDER_STATUS__?: unknown }).__PPTX_RENDER_STATUS__ = {
      state: "ready",
      slideCount: document.querySelectorAll(".pptx-slide").length,
    };
  } catch (error) {
    console.error(error);
    (window as Window & { __PPTX_RENDER_STATUS__?: unknown }).__PPTX_RENDER_STATUS__ = {
      state: "error",
      message: String(error),
    };
  }
}

void main();
