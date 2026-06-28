import { spawnSync } from "node:child_process";
import { existsSync } from "node:fs";

const bundlePath = "bundle/miku-md2pptx.mjs";

if (!existsSync(bundlePath)) {
  throw new Error("bundle/miku-md2pptx.mjs does not exist. Run npm run build:bundle first.");
}

for (const arg of ["--version", "--help"]) {
  const result = spawnSync(process.execPath, [bundlePath, arg], { encoding: "utf8" });
  if (result.status !== 0) {
    process.stderr.write(result.stderr);
    throw new Error(`Bundle smoke failed for ${arg}`);
  }
  if (!result.stdout.trim()) {
    throw new Error(`Bundle smoke produced empty stdout for ${arg}`);
  }
}

console.log("[smoke:bundle] cli ok");
