import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";

import { Presentation, PresentationFile } from "office-artifact-tool";

const outputPath = path.join(os.tmpdir(), `office-artifact-tool-presentation-${process.pid}.pptx`);
const inspectPath = `${outputPath}.inspect.ndjson`;

try {
  await fs.rm(outputPath, { force: true });
  await fs.rm(inspectPath, { force: true });

  const presentation = Presentation.create({
    slideSize: { width: 1280, height: 720 },
  });

  const file = await PresentationFile.exportPptx(presentation);
  await file.save(outputPath);

  const stat = await fs.stat(outputPath);
  if (!stat.isFile() || stat.size <= 0) {
    throw new Error(`Presentation export produced an empty or invalid file: ${outputPath}`);
  }

  console.log("presentation export smoke ok");
} finally {
  await fs.rm(outputPath, { force: true });
  await fs.rm(inspectPath, { force: true });
}
