import { parsePptx, renderSlides } from "../src/index";
import type { PptxParseResult, RenderOptions } from "../src/types";

type PlaygroundStatus =
  | { state: "idle" }
  | { state: "loading"; fileName: string }
  | { state: "ready"; fileName: string; slideCount: number; slideWidth: number; slideHeight: number }
  | { state: "error"; fileName?: string; message: string };

type PlaygroundWindow = Window & {
  __PPTX_PLAYGROUND_STATUS__?: PlaygroundStatus;
};

const state = {
  fileName: "",
  parsed: null as PptxParseResult | null,
  scale: 0.65,
  theme: "light" as RenderOptions["theme"],
  showSlideNumbers: true,
};

const fileInput = must<HTMLInputElement>("#file-input");
const pickFileButton = must<HTMLButtonElement>("#pick-file-button");
const clearButton = must<HTMLButtonElement>("#clear-button");
const dropzone = must<HTMLLabelElement>("#dropzone");
const preview = must<HTMLDivElement>("#preview");
const emptyState = must<HTMLDivElement>("#empty-state");
const statusMessage = must<HTMLParagraphElement>("#status-message");
const scaleInput = must<HTMLInputElement>("#scale-input");
const scaleValue = must<HTMLOutputElement>("#scale-value");
const slideNumbersInput = must<HTMLInputElement>("#slide-numbers-input");
const fileNameValue = must<HTMLElement>("#file-name");
const slideCountValue = must<HTMLElement>("#slide-count");
const slideSizeValue = must<HTMLElement>("#slide-size");
const themeButtons = Array.from(document.querySelectorAll<HTMLButtonElement>("[data-theme]"));

setStatus({ state: "idle" });
updateScaleLabel(state.scale);
bindEvents();

function bindEvents(): void {
  pickFileButton.addEventListener("click", () => fileInput.click());
  fileInput.addEventListener("change", async () => {
    const [file] = Array.from(fileInput.files || []);
    if (file) {
      await loadFile(file);
    }
  });

  clearButton.addEventListener("click", () => {
    fileInput.value = "";
    state.fileName = "";
    state.parsed = null;
    preview.innerHTML = "";
    emptyState.classList.remove("is-hidden");
    setStatus({ state: "idle" });
    syncMetadata();
    renderStatusMessage("Waiting for a file.");
  });

  dropzone.addEventListener("dragenter", (event) => {
    event.preventDefault();
    dropzone.classList.add("is-dragging");
  });
  dropzone.addEventListener("dragover", (event) => {
    event.preventDefault();
    dropzone.classList.add("is-dragging");
  });
  dropzone.addEventListener("dragleave", (event) => {
    if (event.target === dropzone) {
      dropzone.classList.remove("is-dragging");
    }
  });
  dropzone.addEventListener("drop", async (event) => {
    event.preventDefault();
    dropzone.classList.remove("is-dragging");
    const [file] = Array.from(event.dataTransfer?.files || []);
    if (file) {
      await loadFile(file);
    }
  });

  scaleInput.addEventListener("input", () => {
    state.scale = Number(scaleInput.value);
    updateScaleLabel(state.scale);
    rerenderIfReady();
  });

  slideNumbersInput.addEventListener("change", () => {
    state.showSlideNumbers = slideNumbersInput.checked;
    rerenderIfReady();
  });

  for (const button of themeButtons) {
    button.addEventListener("click", () => {
      const nextTheme = button.dataset.theme === "dark" ? "dark" : "light";
      state.theme = nextTheme;
      themeButtons.forEach((candidate) =>
        candidate.classList.toggle("is-active", candidate.dataset.theme === nextTheme),
      );
      rerenderIfReady();
    });
  }
}

async function loadFile(file: File): Promise<void> {
  if (!file.name.toLowerCase().endsWith(".pptx")) {
    setStatus({ state: "error", fileName: file.name, message: "Please choose a .pptx file." });
    syncMetadata();
    renderStatusMessage("Please choose a .pptx file.", true);
    return;
  }

  state.fileName = file.name;
  setStatus({ state: "loading", fileName: file.name });
  syncMetadata();
  renderStatusMessage(`Parsing ${file.name}...`);

  try {
    const buffer = await file.arrayBuffer();
    const parsed = await parsePptx(buffer);
    state.parsed = parsed;
    emptyState.classList.add("is-hidden");
    renderCurrentDeck();
    renderStatusMessage(`Rendered ${parsed.slides.length} slide${parsed.slides.length === 1 ? "" : "s"} from ${file.name}.`);
  } catch (error) {
    state.parsed = null;
    preview.innerHTML = "";
    emptyState.classList.remove("is-hidden");
    const message = error instanceof Error ? error.message : String(error);
    setStatus({ state: "error", fileName: file.name, message });
    syncMetadata();
    renderStatusMessage(`Could not render ${file.name}: ${message}`, true);
  }
}

function rerenderIfReady(): void {
  if (!state.parsed) {
    return;
  }

  renderCurrentDeck();
}

function renderCurrentDeck(): void {
  if (!state.parsed) {
    return;
  }

  renderSlides(state.parsed.slides, {
    container: preview,
    scale: state.scale,
    showSlideNumbers: state.showSlideNumbers,
    theme: state.theme,
  });

  setStatus({
    state: "ready",
    fileName: state.fileName,
    slideCount: state.parsed.slides.length,
    slideWidth: state.parsed.slideWidth,
    slideHeight: state.parsed.slideHeight,
  });
  syncMetadata();
}

function syncMetadata(): void {
  fileNameValue.textContent = state.fileName || "None";
  slideCountValue.textContent = state.parsed ? String(state.parsed.slides.length) : "0";
  slideSizeValue.textContent = state.parsed
    ? `${Math.round(state.parsed.slideWidth)} x ${Math.round(state.parsed.slideHeight)}`
    : "-";
}

function updateScaleLabel(scale: number): void {
  scaleValue.value = `${Math.round(scale * 100)}%`;
}

function renderStatusMessage(message: string, isError = false): void {
  statusMessage.textContent = message;
  statusMessage.classList.toggle("is-error", isError);
}

function setStatus(status: PlaygroundStatus): void {
  (window as PlaygroundWindow).__PPTX_PLAYGROUND_STATUS__ = status;
}

function must<T extends Element>(selector: string): T {
  const element = document.querySelector<T>(selector);
  if (!element) {
    throw new Error(`Missing required element: ${selector}`);
  }
  return element;
}
