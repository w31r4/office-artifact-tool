import { spawn } from "node:child_process";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";

const packageRoot = path.resolve(import.meta.dirname, "..");
const npmCommand = process.platform === "win32" ? "npm.cmd" : "npm";

function run(command, args, options = {}) {
  return new Promise((resolve, reject) => {
    const child = spawn(command, args, {
      cwd: options.cwd,
      env: options.env ?? process.env,
      stdio: ["ignore", "pipe", "pipe"],
    });
    let stdout = "";
    let stderr = "";
    child.stdout.setEncoding("utf8");
    child.stderr.setEncoding("utf8");
    child.stdout.on("data", (chunk) => {
      stdout += chunk;
    });
    child.stderr.on("data", (chunk) => {
      stderr += chunk;
    });
    child.once("error", reject);
    child.once("close", (code) => resolve({ code, stderr, stdout }));
  });
}

async function pathExists(filePath) {
  try {
    await fs.access(filePath);
    return true;
  } catch {
    return false;
  }
}

async function assertExists(filePath, label) {
  if (!(await pathExists(filePath))) {
    throw new Error(`Missing ${label}: ${filePath}`);
  }
}

async function assertAbsent(filePath, label) {
  if (await pathExists(filePath)) {
    throw new Error(`Unexpected ${label}: ${filePath}`);
  }
}

const tempRoot = await fs.mkdtemp(
  path.join(os.tmpdir(), "office-artifact-tool-packed-install-"),
);
const tarballDirectory = path.join(tempRoot, "tarball");
const installRoot = path.join(tempRoot, "install");

try {
  await Promise.all([
    fs.mkdir(tarballDirectory, { recursive: true }),
    fs.mkdir(installRoot, { recursive: true }),
  ]);

  const packed = await run(
    npmCommand,
    ["pack", "--pack-destination", tarballDirectory, "--json"],
    { cwd: packageRoot },
  );
  if (packed.code !== 0) {
    throw new Error(`npm pack failed (${packed.code}): ${packed.stderr}`);
  }
  const packResult = JSON.parse(packed.stdout);
  const filename = packResult[0]?.filename;
  if (typeof filename !== "string" || filename.length === 0) {
    throw new Error(`npm pack did not report a tarball filename: ${packed.stdout}`);
  }
  const tarballPath = path.join(tarballDirectory, filename);
  await assertExists(tarballPath, "packed tarball");

  await fs.writeFile(
    path.join(installRoot, "package.json"),
    `${JSON.stringify({ name: "office-artifact-tool-packed-install-smoke", private: true, type: "module" })}\n`,
    "utf8",
  );
  const installed = await run(
    npmCommand,
    [
      "install",
      "--ignore-scripts",
      "--no-audit",
      "--no-fund",
      "--offline",
      "--no-package-lock",
      tarballPath,
    ],
    { cwd: installRoot },
  );
  if (installed.code !== 0) {
    throw new Error(`clean tarball install failed (${installed.code}): ${installed.stderr}`);
  }

  const installedPackageRoot = path.join(
    installRoot,
    "node_modules",
    "office-artifact-tool",
  );
  await assertExists(installedPackageRoot, "installed package root");
  await assertAbsent(path.join(installedPackageRoot, "skills"), "repository-only skills directory");

  const smokePath = path.join(installRoot, "runtime-smoke.mjs");
  await fs.writeFile(
    smokePath,
    `import { createRequire } from "node:module";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import {
  Presentation,
  PresentationFile,
  SpreadsheetFile,
  Workbook,
} from "office-artifact-tool";
import { Fragment } from "office-artifact-tool/presentation-jsx";

if (typeof Fragment !== "symbol") {
  throw new Error("Installed presentation JSX export is unavailable.");
}

const rootRequire = createRequire(import.meta.url);
const artifactEntry = rootRequire.resolve("office-artifact-tool");
const artifactRequire = createRequire(artifactEntry);
const embeddedFonts = await import(
  artifactRequire.resolve("@officer/embedded-fonts"),
);
if ((await embeddedFonts.loadEmbeddedFontDecoder()) == null) {
  throw new Error("Installed embedded font decoder did not initialize.");
}

const outputRoot = await fs.mkdtemp(
  path.join(os.tmpdir(), "office-artifact-tool-packed-runtime-"),
);
try {
  const presentation = Presentation.create({
    slideSize: { width: 1280, height: 720 },
  });
  presentation.slides.add();
  const pptxPath = path.join(outputRoot, "smoke.pptx");
  await (await PresentationFile.exportPptx(presentation)).save(pptxPath);

  const workbook = Workbook.create();
  const worksheet = workbook.worksheets.add("Smoke");
  worksheet.getRange("A1:B2").values = [["Kind", "Value"], ["runtime", 1]];
  const xlsxPath = path.join(outputRoot, "smoke.xlsx");
  await (await SpreadsheetFile.exportXlsx(workbook)).save(xlsxPath);

  for (const outputPath of [pptxPath, xlsxPath]) {
    const stat = await fs.stat(outputPath);
    if (!stat.isFile() || stat.size <= 0) {
      throw new Error("Installed runtime produced an empty Office artifact: " + outputPath);
    }
  }
} finally {
  await fs.rm(outputRoot, { force: true, recursive: true });
}
`,
    "utf8",
  );
  const smoke = await run(process.execPath, [smokePath], { cwd: installRoot });
  if (smoke.code !== 0) {
    throw new Error(`installed runtime smoke failed (${smoke.code}): ${smoke.stderr}`);
  }

  console.log("packed install smoke ok");
} finally {
  await fs.rm(tempRoot, { force: true, recursive: true });
}
