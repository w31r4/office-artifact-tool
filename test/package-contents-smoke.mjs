import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

const packageRoot = fileURLToPath(new URL("../", import.meta.url));

async function assertExists(relativePath) {
  const absolutePath = path.join(packageRoot, relativePath);
  try {
    await fs.access(absolutePath);
  } catch {
    throw new Error(`Missing required package content: ${relativePath}`);
  }
}

async function assertFileWithPrefix(directoryRelativePath, prefix, suffix) {
  const absoluteDirectory = path.join(packageRoot, directoryRelativePath);
  let entries;
  try {
    entries = await fs.readdir(absoluteDirectory);
  } catch {
    throw new Error(`Missing required directory: ${directoryRelativePath}`);
  }
  const match = entries.find((entry) => entry.startsWith(prefix) && entry.endsWith(suffix));
  if (!match) {
    throw new Error(`Missing required file matching ${prefix}*${suffix} in ${directoryRelativePath}`);
  }
}

await assertExists("dist/artifact_tool.mjs");
await assertExists("dist/presentation-jsx/index.mjs");
await assertExists("dist/presentation-jsx/jsx-runtime.mjs");
await assertExists("dist/presentation-jsx/jsx-dev-runtime.mjs");
await assertExists("node_modules/@officer/walnut/package.json");
await assertExists("node_modules/@officer/walnut/wasm/blazor.boot.json");
await assertFileWithPrefix("node_modules/@officer/walnut/wasm", "DocumentFormat.OpenXml", ".wasm");
await assertFileWithPrefix("node_modules/@officer/walnut/wasm", "System.IO.Packaging", ".wasm");
await assertFileWithPrefix("node_modules/@officer/walnut/wasm", "Google.Protobuf", ".wasm");
await assertExists("node_modules/skia-canvas/package.json");
await assertExists("node_modules/skia-canvas/lib/skia.node");

console.log("package contents smoke ok");
