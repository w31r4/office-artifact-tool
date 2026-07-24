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

const officerElementSymbol = Symbol.for("@officer/granola/presentation-jsx.element");
const officerFragmentSymbol = Symbol.for("@officer/granola/presentation-jsx.fragment");
if (jsxRuntime.jsx("fixture", {}).$$type !== officerElementSymbol) {
  throw new Error("jsx-runtime did not emit the Officer element symbol");
}
if (jsxDevRuntime.jsxDEV("fixture", {}).$$type !== officerElementSymbol) {
  throw new Error("jsx-dev-runtime did not emit the Officer element symbol");
}
if (jsxRuntime.Fragment !== officerFragmentSymbol || jsxDevRuntime.Fragment !== officerFragmentSymbol) {
  throw new Error("JSX runtimes did not emit the Officer fragment symbol");
}

console.log("import smoke ok");
