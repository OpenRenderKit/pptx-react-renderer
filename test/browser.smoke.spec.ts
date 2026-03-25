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
    .toMatchObject({ state: "ready", slideCount: 2 });

  await expect(page.locator(".pptx-slide")).toHaveCount(2);
  await expect(page.locator(".pptx-slide-number")).toHaveCount(2);
  await expect(page.locator(".pptx-text-element")).toContainText(["Quarterly Review"]);
  await expect(page.locator(".pptx-image-element img")).toHaveCount(1);
  await expect(page.locator(".pptx-graphic-frame")).toContainText("Chart");
});
