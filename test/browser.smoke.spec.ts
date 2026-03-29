import path from "node:path";
import { expect, test } from "@playwright/test";

test("renders the committed PPTX fixture in a real browser", async ({ page }) => {
  await page.goto("/test/browser/harness/");

  await expect
    .poll(
      async () =>
        page.evaluate(
          () => (window as Window & { __PPTX_RENDER_STATUS__?: unknown }).__PPTX_RENDER_STATUS__,
        ),
      {
      message: "browser harness should finish rendering",
      },
    )
    .toMatchObject({ state: "ready", slideCount: 12 });

  await expect(page.locator(".pptx-slide")).toHaveCount(12);
  await expect(page.locator(".pptx-slide-number")).toHaveCount(12);
  await expect(page.locator(".pptx-text-element")).toContainText([
    "OpenBridge",
    "Evaluation Matrix",
  ]);
  await expect(page.locator(".pptx-image-element img")).toHaveCount(16);
  await expect(page.locator(".pptx-graphic-frame")).toHaveCount(3);
  await expect(page.locator(".pptx-table-element")).toHaveCount(3);
});

test("uploads a local PPTX in the playground and rerenders with live controls", async ({ page }) => {
  const fixturePath = path.resolve("test/fixtures/real/basic-preview.pptx");

  await page.goto("/playground/");
  await expect(page.locator("#empty-state")).toBeVisible();
  await page.setInputFiles("#file-input", fixturePath);

  await expect
    .poll(
      async () =>
        page.evaluate(
          () => (window as Window & { __PPTX_PLAYGROUND_STATUS__?: unknown }).__PPTX_PLAYGROUND_STATUS__,
        ),
      { message: "playground should finish rendering the uploaded PPTX" },
    )
    .toMatchObject({
      state: "ready",
      fileName: "basic-preview.pptx",
      slideCount: 2,
      slideWidth: 960,
      slideHeight: 720,
    });

  await expect(page.locator(".pptx-slide")).toHaveCount(2);
  await expect(page.locator("#slide-count")).toHaveText("2");
  await expect(page.locator("#file-name")).toHaveText("basic-preview.pptx");
  await expect(page.locator(".pptx-text-element")).toContainText(["Quarterly Review", "Grouped Card"]);

  await page.getByRole("button", { name: "Dark" }).click();
  await page.getByLabel("Slide labels").uncheck();

  await expect(page.locator(".pptx-slide-number")).toHaveCount(0);
  await expect(page.locator("#status-message")).toContainText("Rendered 2 slides");

  await page.getByRole("button", { name: "Clear" }).click();
  await expect(page.locator("#empty-state")).toBeVisible();
  await expect(page.locator(".pptx-slide")).toHaveCount(0);
  await expect(page.locator("#file-name")).toHaveText("None");
});
