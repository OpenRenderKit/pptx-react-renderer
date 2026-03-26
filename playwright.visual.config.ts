import { defineConfig } from "@playwright/test";

export default defineConfig({
  testDir: "./test",
  testMatch: ["visual.regression.spec.ts"],
  timeout: 120_000,
  reporter: [["list"], ["html", { open: "never" }]],
  use: {
    baseURL: "http://127.0.0.1:4173",
    headless: true,
    trace: "retain-on-failure",
    screenshot: "only-on-failure",
    video: "retain-on-failure",
  },
  webServer: {
    command: "python3 -m http.server 4173",
    port: 4173,
    reuseExistingServer: true,
    stdout: "ignore",
    stderr: "pipe",
  },
});
