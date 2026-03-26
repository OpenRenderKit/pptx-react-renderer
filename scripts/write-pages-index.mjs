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
        <div class="eyebrow">Pull Request</div>
        <h2>#${prNumber}</h2>
        <p>Published visual review surfaces for this PR.</p>
        <div class="actions">
          <a href="./pr-previews/${prNumber}/visual-pr-comparison/index.html">PR vs main</a>
          <a href="./pr-previews/${prNumber}/visual-regression/index.html">Office vs PR</a>
        </div>
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
        font-family: "Inter", ui-sans-serif, system-ui, sans-serif;
        color-scheme: light;
        background:
          radial-gradient(circle at top left, rgba(27, 96, 255, 0.16), transparent 32%),
          radial-gradient(circle at top right, rgba(0, 182, 155, 0.14), transparent 28%),
          linear-gradient(180deg, #f3f7ff 0%, #f8fbff 100%);
        color: #122038;
      }
      * { box-sizing: border-box; }
      body { margin: 0; min-height: 100vh; padding: 40px 24px 72px; }
      main { max-width: 1180px; margin: 0 auto; display: grid; gap: 28px; }
      .hero {
        padding: 28px 32px;
        border-radius: 28px;
        background: rgba(255, 255, 255, 0.78);
        border: 1px solid rgba(203, 214, 233, 0.7);
        box-shadow: 0 28px 80px rgba(15, 32, 62, 0.10);
        backdrop-filter: blur(18px);
      }
      h1 { margin: 0 0 12px; font-size: clamp(36px, 6vw, 64px); line-height: 0.98; letter-spacing: -0.05em; }
      .hero p { max-width: 780px; margin: 0; color: #46556e; font-size: 18px; line-height: 1.5; }
      .grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(260px, 1fr)); gap: 20px; }
      .card {
        display: grid;
        gap: 12px;
        padding: 22px;
        border-radius: 22px;
        background: rgba(255, 255, 255, 0.9);
        border: 1px solid rgba(207, 217, 234, 0.9);
        box-shadow: 0 24px 70px rgba(17, 32, 58, 0.08);
      }
      .eyebrow {
        width: fit-content;
        padding: 6px 10px;
        border-radius: 999px;
        background: rgba(24, 93, 216, 0.08);
        color: #0f57c8;
        font-size: 12px;
        font-weight: 700;
        letter-spacing: 0.08em;
        text-transform: uppercase;
      }
      .card h2 { margin: 0; font-size: 30px; letter-spacing: -0.03em; }
      .card p { margin: 0; color: #46556e; line-height: 1.5; }
      .actions { display: flex; flex-wrap: wrap; gap: 12px; padding-top: 8px; }
      .actions a {
        text-decoration: none;
        color: white;
        background: linear-gradient(135deg, #1257da 0%, #0f7bdc 100%);
        padding: 11px 14px;
        border-radius: 14px;
        font-weight: 600;
        box-shadow: inset 0 1px 0 rgba(255,255,255,0.24);
      }
      .actions a:last-child {
        background: linear-gradient(135deg, #0b7a72 0%, #11948a 100%);
      }
      @media (max-width: 680px) {
        body { padding: 24px 16px 48px; }
        .hero { padding: 22px; }
      }
    </style>
  </head>
  <body>
    <main>
      <section class="hero">
        <h1>Visual PR Previews</h1>
        <p>GitHub Pages snapshots for <code>pptx-react-renderer</code> pull requests. Each preview publishes two reports: <code>PR vs main</code> to isolate renderer changes, and <code>Office vs PR</code> to measure fidelity against the office-rendered baseline.</p>
      </section>
      <section class="grid">
        ${cards || "<p>No PR previews published yet.</p>"}
      </section>
    </main>
  </body>
</html>`;

await writeFile(path.join(siteRoot, "index.html"), html);
