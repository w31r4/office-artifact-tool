import assert from "node:assert/strict";
import crypto from "node:crypto";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { crc32 } from "node:zlib";

const packageRoot = path.resolve(import.meta.dirname, "..");
const libraryRoot = path.join(packageRoot, "skills", "default-template-library");
const configuredSourceRoot = process.env.OFFICE_TEMPLATE_SOURCE_ROOT;
const sourceRoot = configuredSourceRoot
  ?? path.join(os.homedir(), ".codex", "plugins", "cache", "openai-curated-remote", "openai-templates", "0.1.0");
const expectedSourceAssetAggregateSha256 = "b3b42fd92cab7f0c9fc1bd4ec8bf0ff9c6b199247b1a9729ad04ef6fc0a49b25";

const templates = [
  { id: "artifact-template-design-report", displayName: "Design Report", kind: "document", referenceExtension: ".docx" },
  { id: "artifact-template-experiment-analysis", displayName: "Experiment Analysis", kind: "document", referenceExtension: ".docx" },
  { id: "artifact-template-investment-committee-memo", displayName: "Investment Committee Memo", kind: "document", referenceExtension: ".docx" },
  { id: "artifact-template-legal-memorandum", displayName: "Legal Memorandum", kind: "document", referenceExtension: ".docx" },
  { id: "artifact-template-minimal-letterhead", displayName: "Minimal Letterhead", kind: "document", referenceExtension: ".docx" },
  { id: "artifact-template-strategy-memorandum", displayName: "Strategy Memorandum", kind: "document", referenceExtension: ".docx" },
  { id: "artifact-template-system-design", displayName: "System Design", kind: "document", referenceExtension: ".docx" },
  { id: "artifact-template-business-review", displayName: "Business Review", kind: "presentation", referenceExtension: ".pptx" },
  { id: "artifact-template-market-trends-report", displayName: "Market Trends Report", kind: "presentation", referenceExtension: ".pptx" },
  { id: "artifact-template-operating-review", displayName: "Operating Review", kind: "presentation", referenceExtension: ".pptx" },
  { id: "artifact-template-project-kickoff", displayName: "Project Kickoff", kind: "presentation", referenceExtension: ".pptx" },
  { id: "artifact-template-simple-dark-mode", displayName: "Simple Dark Mode", kind: "presentation", referenceExtension: ".pptx" },
  { id: "artifact-template-simple-light-mode", displayName: "Simple Light Mode", kind: "presentation", referenceExtension: ".pptx" },
  { id: "artifact-template-team-alignment", displayName: "Team Alignment", kind: "presentation", referenceExtension: ".pptx" },
  { id: "artifact-template-analytics-dashboard", displayName: "Analytics Dashboard", kind: "spreadsheet", referenceExtension: ".xlsx" },
  { id: "artifact-template-financial-budget", displayName: "Financial Budget", kind: "spreadsheet", referenceExtension: ".xlsx" },
  { id: "artifact-template-operating-calendar", displayName: "Operating Calendar", kind: "spreadsheet", referenceExtension: ".xlsx" },
  { id: "artifact-template-project-tracker", displayName: "Project Tracker", kind: "spreadsheet", referenceExtension: ".xlsx" },
  { id: "artifact-template-sales-pipeline", displayName: "Sales Pipeline", kind: "spreadsheet", referenceExtension: ".xlsx" },
  { id: "artifact-template-three-statement-forecast", displayName: "Three-Statement Forecast", kind: "spreadsheet", referenceExtension: ".xlsx" },
];

const sourceMarker = /\bopenai\b|\bcodex\b|plugin:\/\/|connector_[a-z0-9_]+|\.codex-plugin|openai-curated-remote|(?:^|[\\/])\.codex[\\/]plugins[\\/]cache/iu;

async function exists(filePath) {
  return fs.access(filePath).then(() => true).catch(() => false);
}

async function walkFiles(root) {
  const files = [];
  const queue = [root];
  for (let index = 0; index < queue.length; index += 1) {
    const entries = await fs.readdir(queue[index], { withFileTypes: true });
    for (const entry of entries) {
      const fullPath = path.join(queue[index], entry.name);
      if (entry.isDirectory()) queue.push(fullPath);
      else if (entry.isFile()) files.push(fullPath);
    }
  }
  return files.sort();
}

function yamlValue(text, key) {
  const value = text.match(new RegExp(`^\\s{2}${key}:\\s*(.+)$`, "mu"))?.[1]?.trim();
  if (!value) return undefined;
  if ((value.startsWith('"') && value.endsWith('"')) || (value.startsWith("'") && value.endsWith("'"))) {
    try {
      return JSON.parse(value);
    } catch {
      return value.slice(1, -1);
    }
  }
  return value;
}

function hasValidPngStructure(bytes) {
  const signature = Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]);
  if (bytes.length < signature.length + 12 || !bytes.subarray(0, signature.length).equals(signature)) return false;
  let offset = signature.length;
  let sawIhdr = false;
  while (offset + 12 <= bytes.length) {
    const length = bytes.readUInt32BE(offset);
    const dataStart = offset + 8;
    const dataEnd = dataStart + length;
    const checksumEnd = dataEnd + 4;
    if (dataEnd < dataStart || checksumEnd > bytes.length) return false;
    const type = bytes.toString("ascii", offset + 4, offset + 8);
    const data = bytes.subarray(dataStart, dataEnd);
    if ((crc32(Buffer.concat([Buffer.from(type, "ascii"), data])) >>> 0) !== bytes.readUInt32BE(dataEnd)) return false;
    if (!sawIhdr) {
      if (type !== "IHDR" || length !== 13) return false;
      sawIhdr = true;
    }
    offset = checksumEnd;
    if (type === "IEND") return sawIhdr && length === 0 && offset === bytes.length;
  }
  return false;
}

function isZipOfficeFile(bytes) {
  return bytes.length >= 4 && bytes[0] === 0x50 && bytes[1] === 0x4b && bytes[2] === 0x03 && bytes[3] === 0x04;
}

function updateAssetAggregate(hash, relativePath, bytes) {
  hash.update(relativePath, "utf8");
  hash.update("\0", "utf8");
  hash.update(bytes);
  hash.update("\0", "utf8");
}

try {
  await fs.access(libraryRoot);
} catch (error) {
  if (error?.code === "ENOENT") {
    console.log("default template library smoke skipped: repository-only skills are not packaged");
    process.exit(0);
  }
  throw error;
}

const manifest = JSON.parse(await fs.readFile(path.join(libraryRoot, "manifest.json"), "utf8"));
assert.equal(manifest.schemaVersion, 1);
assert.equal(manifest.name, "default-template-library");
assert.deepEqual(manifest.skills, templates.map((template) => `skills/${template.id}`));
assert.equal(manifest.assets?.icon, "assets/icon.svg");

const kindCounts = templates.reduce((counts, template) => {
  counts[template.kind] += 1;
  return counts;
}, { document: 0, presentation: 0, spreadsheet: 0 });
assert.deepEqual(kindCounts, { document: 7, presentation: 7, spreadsheet: 6 });

const expectedFiles = new Set(["README.md", "manifest.json", "assets/icon.svg"]);
const sourceAvailable = await exists(sourceRoot);
if (configuredSourceRoot && !sourceAvailable) {
  throw new Error(`OFFICE_TEMPLATE_SOURCE_ROOT does not exist: ${sourceRoot}`);
}
const targetAssetAggregate = crypto.createHash("sha256");
const sourceAssetAggregate = sourceAvailable ? crypto.createHash("sha256") : null;
for (const template of [...templates].sort((left, right) => left.id.localeCompare(right.id))) {
  const targetSkillRoot = path.join(libraryRoot, "skills", template.id);
  const referenceRelativePath = `assets/reference${template.referenceExtension}`;
  const expectedSkillFiles = [
    "SKILL.md",
    "artifact-template.json",
    "agents/agent.yaml",
    "assets/preview.png",
    referenceRelativePath,
  ];
  for (const relativePath of expectedSkillFiles) expectedFiles.add(path.posix.join("skills", template.id, relativePath));

  const [skillText, targetSidecarBytes, agentText, previewBytes, referenceBytes] = await Promise.all([
    fs.readFile(path.join(targetSkillRoot, "SKILL.md"), "utf8"),
    fs.readFile(path.join(targetSkillRoot, "artifact-template.json")),
    fs.readFile(path.join(targetSkillRoot, "agents", "agent.yaml"), "utf8"),
    fs.readFile(path.join(targetSkillRoot, "assets", "preview.png")),
    fs.readFile(path.join(targetSkillRoot, referenceRelativePath)),
  ]);
  const sidecar = JSON.parse(targetSidecarBytes.toString("utf8"));

  assert.equal(skillText.match(/^name:\s*(.+)$/mu)?.[1]?.trim(), template.id, `${template.id} frontmatter ID`);
  assert.equal(yamlValue(agentText, "display_name"), template.displayName, `${template.id} agent display name`);
  assert.equal(yamlValue(agentText, "icon_large"), "./assets/preview.png", `${template.id} preview icon`);
  assert.equal(await exists(path.join(targetSkillRoot, "agents", "openai.yaml")), false, `${template.id} must not retain source agent metadata`);
  assert.deepEqual(sidecar, {
    schemaVersion: 1,
    kind: template.kind,
    reference: referenceRelativePath,
    preview: "assets/preview.png",
  }, `${template.id} sidecar`);
  assert.equal(hasValidPngStructure(previewBytes), true, `${template.id} preview PNG`);
  assert.equal(isZipOfficeFile(referenceBytes), true, `${template.id} Office reference`);
  assert(previewBytes.length > 0 && referenceBytes.length > 0, `${template.id} assets must be nonempty`);
  assert.equal(sourceMarker.test(skillText), false, `${template.id} canonical SKILL.md source marker`);
  assert.equal(sourceMarker.test(agentText), false, `${template.id} canonical agent metadata source marker`);

  const previewRelativePath = path.posix.join("skills", template.id, "assets/preview.png");
  const referenceAssetRelativePath = path.posix.join("skills", template.id, referenceRelativePath);
  updateAssetAggregate(targetAssetAggregate, previewRelativePath, previewBytes);
  updateAssetAggregate(targetAssetAggregate, referenceAssetRelativePath, referenceBytes);

  if (sourceAvailable) {
    const sourceSkillRoot = path.join(sourceRoot, "skills", template.id);
    const [sourcePreviewBytes, sourceReferenceBytes, sourceSidecarBytes] = await Promise.all([
      fs.readFile(path.join(sourceSkillRoot, "assets", "preview.png")),
      fs.readFile(path.join(sourceSkillRoot, referenceRelativePath)),
      fs.readFile(path.join(sourceSkillRoot, "artifact-template.json")),
    ]);
    assert.deepEqual(targetSidecarBytes, sourceSidecarBytes, `${template.id} sidecar bytes`);
    assert.deepEqual(previewBytes, sourcePreviewBytes, `${template.id} preview bytes`);
    assert.deepEqual(referenceBytes, sourceReferenceBytes, `${template.id} reference bytes`);
    updateAssetAggregate(sourceAssetAggregate, previewRelativePath, sourcePreviewBytes);
    updateAssetAggregate(sourceAssetAggregate, referenceAssetRelativePath, sourceReferenceBytes);
  }
}

const actualFiles = (await walkFiles(libraryRoot))
  .map((filePath) => path.relative(libraryRoot, filePath).split(path.sep).join("/"))
  .sort();
assert.deepEqual(actualFiles, [...expectedFiles].sort(), "catalog contains only its canonical files");
for (const filePath of await walkFiles(libraryRoot)) {
  const relativePath = path.relative(libraryRoot, filePath).split(path.sep).join("/");
  if (relativePath.endsWith(".png") || /\/assets\/reference\.(docx|pptx|xlsx)$/u.test(relativePath)) continue;
  assert.equal(sourceMarker.test(await fs.readFile(filePath, "utf8")), false, `${relativePath} source marker`);
}

const targetAssetAggregateSha256 = targetAssetAggregate.digest("hex");
assert.equal(targetAssetAggregateSha256, expectedSourceAssetAggregateSha256, "target binary aggregate must match the retained source inventory");
if (sourceAssetAggregate) {
  assert.equal(sourceAssetAggregate.digest("hex"), targetAssetAggregateSha256, "available source binary aggregate must match the target catalog");
}
console.log("default template library smoke ok");
