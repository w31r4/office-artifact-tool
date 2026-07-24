import crypto from "node:crypto";
import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

const packageRoot = fileURLToPath(new URL("../", import.meta.url));
const EMBEDDED_FONT_DECODER_SHA256 =
  "e84b106e4b664cc7d6722092cc22deb17b71c4df8d9c42676203e97da4c1586d";

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

async function assertPackageMetadata(relativePath, expectedName, expectedVersion) {
  const metadata = JSON.parse(await fs.readFile(path.join(packageRoot, relativePath), "utf8"));
  if (metadata.name !== expectedName || metadata.version !== expectedVersion) {
    throw new Error(`Unexpected package metadata in ${relativePath}: ${metadata.name}@${metadata.version}`);
  }
}

async function assertFileIntegrity(absolutePath, expectedIntegrity, label) {
  const actualIntegrity = `sha256-${crypto
    .createHash("sha256")
    .update(await fs.readFile(absolutePath))
    .digest("base64")}`;
  if (actualIntegrity !== expectedIntegrity) {
    throw new Error(
      `${label} integrity mismatch: expected ${expectedIntegrity}, received ${actualIntegrity}`,
    );
  }
}

function collectDeclaredWalnutIntegrity(resources) {
  const declaredIntegrity = new Map();
  function visit(value, pathParts = []) {
    if (value == null || typeof value !== "object" || Array.isArray(value)) {
      return;
    }
    for (const [name, child] of Object.entries(value)) {
      if (pathParts.length === 0 && (name === "fingerprinting" || name === "hash")) {
        continue;
      }
      if (typeof child === "string") {
        if (!child.startsWith("sha256-")) continue;
        if (declaredIntegrity.has(name) && declaredIntegrity.get(name) !== child) {
          throw new Error(`Walnut boot manifest declares conflicting integrity for ${name}`);
        }
        declaredIntegrity.set(name, child);
        continue;
      }
      visit(child, [...pathParts, name]);
    }
  }
  visit(resources);
  return declaredIntegrity;
}

async function assertWalnutFingerprintAssets() {
  const wasmRoot = path.join(packageRoot, "node_modules/@officer/walnut/wasm");
  const boot = JSON.parse(await fs.readFile(path.join(wasmRoot, "blazor.boot.json"), "utf8"));
  const referenced = new Set(Object.keys(boot.resources?.fingerprinting ?? {}));
  if (referenced.size === 0) throw new Error("Walnut boot manifest has no fingerprinted assets");
  for (const relativePath of referenced) {
    await fs.access(path.join(wasmRoot, relativePath));
  }

  const declaredIntegrity = collectDeclaredWalnutIntegrity(boot.resources);
  if (declaredIntegrity.size === 0) {
    throw new Error("Walnut boot manifest has no integrity declarations");
  }
  for (const [relativePath, integrity] of declaredIntegrity) {
    if (!referenced.has(relativePath)) {
      throw new Error(`Walnut integrity references a non-fingerprinted asset: ${relativePath}`);
    }
    await assertFileIntegrity(
      path.join(wasmRoot, relativePath),
      integrity,
      `Walnut ${relativePath}`,
    );
  }
}

async function assertEmbeddedFontDecoder() {
  const decoderPath = path.join(
    packageRoot,
    "node_modules/@officer/embedded-fonts/embedded_font_decoder.wasm",
  );
  const actual = crypto
    .createHash("sha256")
    .update(await fs.readFile(decoderPath))
    .digest("hex");
  if (actual !== EMBEDDED_FONT_DECODER_SHA256) {
    throw new Error(
      `Embedded font decoder checksum mismatch: expected ${EMBEDDED_FONT_DECODER_SHA256}, received ${actual}`,
    );
  }
}

await assertExists("dist/artifact_tool.mjs");
await assertExists("dist/presentation-jsx/index.mjs");
await assertExists("dist/presentation-jsx/jsx-runtime.mjs");
await assertExists("dist/presentation-jsx/jsx-dev-runtime.mjs");
await assertExists("node_modules/@officer/walnut/package.json");
await assertPackageMetadata("node_modules/@officer/walnut/package.json", "@officer/walnut", "0.1.231");
await assertExists("node_modules/@officer/walnut/wasm/blazor.boot.json");
await assertWalnutFingerprintAssets();
await assertFileWithPrefix("node_modules/@officer/walnut/wasm", "DocumentFormat.OpenXml", ".wasm");
await assertFileWithPrefix("node_modules/@officer/walnut/wasm", "System.IO.Packaging", ".wasm");
await assertFileWithPrefix("node_modules/@officer/walnut/wasm", "Google.Protobuf", ".wasm");
await assertExists("node_modules/@officer/embedded-fonts/package.json");
await assertPackageMetadata("node_modules/@officer/embedded-fonts/package.json", "@officer/embedded-fonts", "0.1.1");
await assertExists("node_modules/@officer/embedded-fonts/dist/index.js");
await assertExists("node_modules/@officer/embedded-fonts/embedded_font_decoder.wasm");
await assertEmbeddedFontDecoder();
await assertExists("node_modules/skia-canvas/package.json");
await assertPackageMetadata("node_modules/skia-canvas/package.json", "skia-canvas", "3.0.8");
await assertExists("node_modules/skia-canvas/lib/skia.node");

console.log("package contents smoke ok");
