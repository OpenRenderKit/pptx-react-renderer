import { mkdir, readdir, writeFile } from "node:fs/promises";
import path from "node:path";

const siteRoot = path.resolve(process.argv[2] || process.env.PPTX_PAGES_ROOT || ".tmp/pages-site");
const previewsRoot = path.join(siteRoot, "pr-previews");

await mkdir(previewsRoot, { recursive: true });

const previewDirs = (await readdir(previewsRoot, { withFileTypes: true }))
  .filter((entry) => entry.isDirectory())
  .map((entry) => entry.name)
  .sort((left, right) => Number(right) - Number(left));

const cards = previewDirs
  .map(
    (prNumber) => `
      <article class="card">
        <h2>PR #${prNumber}</h2>
        <p><a href="./pr-previews/${prNumber}/visual-pr-comparison/index.html">PR vs main report</a></p>
        <p><a href="./pr-previews/${prNumber}/visual-regression/index.html">Office reference vs PR report</a></p>
      </article>
    `,
  )
  .join("\n");

const html = `<!DOCTYPE html>
<html lang="en">
  <head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <title>pptx-react-renderer PR Previews</title>
    <style>
      :root {
        font-family: Inter, ui-sans-serif, system-ui, sans-serif;
        color-scheme: light;
        background: #f4f7fb;
        color: #162133;
      }
      * { box-sizing: border-box; }
      body { margin: 0; padding: 32px; }
      main { max-width: 1100px; margin: 0 auto; }
      .grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(260px, 1fr)); gap: 20px; margin-top: 24px; }
      .card { background: white; border-radius: 20px; padding: 24px; box-shadow: 0 24px 60px rgba(20, 32, 51, 0.08); }
      a { color: #0a5bd8; text-decoration: none; }
      a:hover { text-decoration: underline; }
      p { margin: 8px 0 0; }
    </style>
  </head>
  <body>
    <main>
      <h1>pptx-react-renderer PR Visual Previews</h1>
      <p>Published GitHub Pages previews for pull request visual-regression runs.</p>
      <section class="grid">
        ${cards || "<p>No PR previews published yet.</p>"}
      </section>
    </main>
  </body>
</html>`;

await writeFile(path.join(siteRoot, "index.html"), html);
