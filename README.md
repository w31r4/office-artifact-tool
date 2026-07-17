# office-artifact-tool

Standalone ESM package for Office-style artifact generation and manipulation.

This package is a standalone Office artifact toolkit, distributed under the MIT license.

## Usage

```js
import { Presentation, PresentationFile, Workbook, SpreadsheetFile, DocumentFile } from "office-artifact-tool";

const presentation = Presentation.create({
  slideSize: { width: 1280, height: 720 },
});

const file = await PresentationFile.exportPptx(presentation);
await file.save("deck.pptx");
```

JSX helpers for presentation composition are exposed as subpath exports:

```js
import { Fragment } from "office-artifact-tool/presentation-jsx";
import { jsx, jsxs } from "office-artifact-tool/presentation-jsx/jsx-runtime";
```

## Exports

- `office-artifact-tool`
- `office-artifact-tool/presentation-jsx`
- `office-artifact-tool/presentation-jsx/jsx-runtime`
- `office-artifact-tool/presentation-jsx/jsx-dev-runtime`

## Runtime contents

The package intentionally vendors the runtime assets required by its ESM runtime:

- `@officer/walnut@0.1.225`, including the OpenXML/.NET WASM assets under `node_modules/@officer/walnut/wasm/`
- `skia-canvas`, including the native `lib/skia.node` binding

Do not remove those files from the published tarball unless the package runtime is updated to resolve them from another location.

## Verification

```sh
npm test
npm run test:pack
```

For release preparation, also install the generated tarball into a clean temporary project and run an import smoke test there.

## License

See `LICENSE.md`. This package is distributed under the MIT license.
