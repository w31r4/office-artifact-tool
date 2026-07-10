const mainModule = await import("office-artifact-tool");

const requiredExports = [
  "Workbook",
  "SpreadsheetFile",
  "Presentation",
  "PresentationFile",
  "DocumentFile",
];

const missing = requiredExports.filter((name) => !(name in mainModule));
if (missing.length > 0) {
  throw new Error(`Missing expected exports: ${missing.join(", ")}`);
}

const jsxModule = await import("office-artifact-tool/presentation-jsx");
if (!jsxModule || Object.keys(jsxModule).length === 0) {
  throw new Error("presentation-jsx subpath imported but exposed no exports");
}

const jsxRuntime = await import("office-artifact-tool/presentation-jsx/jsx-runtime");
for (const name of ["jsx", "jsxs", "Fragment"]) {
  if (!(name in jsxRuntime)) {
    throw new Error(`Missing jsx-runtime export: ${name}`);
  }
}

const jsxDevRuntime = await import("office-artifact-tool/presentation-jsx/jsx-dev-runtime");
for (const name of ["jsxDEV", "Fragment"]) {
  if (!(name in jsxDevRuntime)) {
    throw new Error(`Missing jsx-dev-runtime export: ${name}`);
  }
}

console.log("import smoke ok");
