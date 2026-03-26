import { access, cp, mkdir, mkdtemp, readdir, rm } from "node:fs/promises";
import { constants } from "node:fs";
import os from "node:os";
import path from "node:path";
import { spawn } from "node:child_process";

const fixturePath = path.resolve(process.env.PPTX_VISUAL_FIXTURE || "test/fixtures/real/complex-sanitized.pptx");
const fixtureName = path.basename(fixturePath, ".pptx");
const outputDir = path.resolve(
  process.env.PPTX_VISUAL_REFERENCE_DIR || path.join(".tmp", "visual-reference", fixtureName),
);

const sofficePath = await findExecutable(
  process.env.SOFFICE_BIN,
  "soffice",
  "/Applications/LibreOffice.app/Contents/MacOS/soffice",
  path.join(os.homedir(), "Applications", "LibreOffice.app", "Contents", "MacOS", "soffice"),
  "/usr/bin/soffice",
  "/snap/bin/libreoffice",
);
const pdftoppmPath = await findExecutable(process.env.PDFTOPPM_BIN, "pdftoppm", "/opt/homebrew/bin/pdftoppm");

if (!sofficePath) {
  throw new Error(
    "Unable to find LibreOffice `soffice`. Install LibreOffice and expose `soffice` on PATH, or set SOFFICE_BIN.",
  );
}

if (!pdftoppmPath) {
  throw new Error(
    "Unable to find `pdftoppm`. Install poppler and expose `pdftoppm` on PATH, or set PDFTOPPM_BIN.",
  );
}

const tempDir = await mkdtemp(path.join(os.tmpdir(), "pptx-react-renderer-visual-"));

try {
  await rm(outputDir, { recursive: true, force: true });
  await mkdir(outputDir, { recursive: true });

  await runCommand(sofficePath, [
    "--headless",
    "--nologo",
    "--nodefault",
    "--nolockcheck",
    "--convert-to",
    "pdf",
    "--outdir",
    tempDir,
    fixturePath,
  ]);

  const pdfPath = path.join(tempDir, `${fixtureName}.pdf`);
  await access(pdfPath, constants.R_OK);

  const imagePrefix = path.join(tempDir, "slide");
  await runCommand(pdftoppmPath, ["-png", "-r", "96", pdfPath, imagePrefix]);

  const slideImages = (await readdir(tempDir))
    .filter((name) => /^slide-\d+\.png$/.test(name))
    .sort((left, right) => left.localeCompare(right, undefined, { numeric: true }));

  if (slideImages.length === 0) {
    throw new Error("Reference rendering produced no slide PNGs.");
  }

  await Promise.all(
    slideImages.map((fileName, index) =>
      cp(path.join(tempDir, fileName), path.join(outputDir, `slide-${String(index + 1).padStart(2, "0")}.png`)),
    ),
  );

  console.log(`Rendered ${slideImages.length} reference slides into ${outputDir}`);
} finally {
  await rm(tempDir, { recursive: true, force: true });
}

async function findExecutable(...candidates) {
  for (const candidate of candidates) {
    if (!candidate) continue;
    const resolved = await resolveExecutable(candidate);
    if (resolved) return resolved;
  }

  return undefined;
}

async function resolveExecutable(candidate) {
  if (candidate.includes(path.sep)) {
    try {
      await access(candidate, constants.X_OK);
      return candidate;
    } catch {
      return undefined;
    }
  }

  return new Promise((resolve) => {
    const child = spawn("sh", ["-lc", `command -v ${candidate}`], {
      stdio: ["ignore", "pipe", "ignore"],
    });

    let output = "";
    child.stdout.on("data", (chunk) => {
      output += String(chunk);
    });
    child.on("close", (code) => {
      resolve(code === 0 ? output.trim() || undefined : undefined);
    });
  });
}

function runCommand(command, args) {
  return new Promise((resolve, reject) => {
    const child = spawn(command, args, { stdio: "pipe" });
    let stderr = "";

    child.stderr.on("data", (chunk) => {
      stderr += String(chunk);
    });

    child.on("close", (code) => {
      if (code === 0) {
        resolve();
        return;
      }

      reject(new Error(`Command failed (${command} ${args.join(" ")}): ${stderr.trim()}`));
    });

    child.on("error", reject);
  });
}
