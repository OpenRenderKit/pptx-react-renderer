import { renderPptx } from "../../../dist/index.js";

async function main() {
  const container = document.getElementById("preview");
  const params = new URLSearchParams(window.location.search);
  const fixture = params.get("fixture") || "/test/fixtures/real/complex-sanitized.pptx";
  const scale = Number(params.get("scale") || "0.6");
  const showSlideNumbers = params.get("slideNumbers") !== "0";
  const theme = params.get("theme") === "dark" ? "dark" : "light";
  window.__PPTX_RENDER_STATUS__ = { state: "loading" };

  if (!container) {
    throw new Error("Missing preview container");
  }

  try {
    const response = await fetch(fixture);
    const arrayBuffer = await response.arrayBuffer();

    await renderPptx(arrayBuffer, {
      container,
      scale,
      showSlideNumbers,
      theme,
    });

    window.__PPTX_RENDER_STATUS__ = {
      state: "ready",
      fixture,
      scale,
      showSlideNumbers,
      slideCount: document.querySelectorAll(".pptx-slide").length,
    };
  } catch (error) {
    console.error(error);
    window.__PPTX_RENDER_STATUS__ = {
      state: "error",
      message: String(error),
    };
  }
}

void main();
