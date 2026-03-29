import { copyFile, mkdir } from "node:fs/promises";
import path from "node:path";
import { build } from "esbuild";

const cliArgs = process.argv.slice(2).filter((arg) => arg !== "--");
const siteRoot = path.resolve(cliArgs[0] || process.env.PPTX_PAGES_ROOT || ".tmp/pages-site");
const sourceRoot = path.resolve("playground");
const outputRoot = path.join(siteRoot, "playground");

await mkdir(outputRoot, { recursive: true });
await copyFile(path.join(sourceRoot, "index.html"), path.join(outputRoot, "index.html"));
await copyFile(path.join(sourceRoot, "styles.css"), path.join(outputRoot, "styles.css"));

await build({
  entryPoints: [path.join(sourceRoot, "app.ts")],
  outfile: path.join(outputRoot, "bundle.js"),
  bundle: true,
  platform: "browser",
  format: "esm",
  target: "es2022",
  sourcemap: false,
});
