import { spawn } from "node:child_process";
import path from "node:path";

const packageRoot = path.resolve(import.meta.dirname, "..");
const npmCommand = process.platform === "win32" ? "npm.cmd" : "npm";

const result = await new Promise((resolve, reject) => {
  const child = spawn(npmCommand, ["pack", "--dry-run", "--json"], {
    cwd: packageRoot,
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

if (result.code !== 0) {
  throw new Error(`npm pack --dry-run failed (${result.code}): ${result.stderr}`);
}

const pack = JSON.parse(result.stdout);
const files = pack[0]?.files?.map((file) => file.path) ?? [];
const includedSkills = files.filter(
  (file) => file === "skills" || file.startsWith("skills/"),
);
if (includedSkills.length > 0) {
  throw new Error(`Repository-only skills unexpectedly entered the npm tarball: ${includedSkills.join(", ")}`);
}

console.log("package skill exclusion smoke ok");
