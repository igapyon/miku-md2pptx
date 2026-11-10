import { existsSync } from "node:fs";

const runtimeBundlePath = "bundle/miku-md2pptx-runtime.mjs";

if (!existsSync(runtimeBundlePath)) {
  throw new Error("bundle/miku-md2pptx-runtime.mjs does not exist. Run npm run build:bundle first.");
}

const runtime = await import(`../${runtimeBundlePath}`);

const runtimeSource = await (await import("node:fs/promises")).readFile(runtimeBundlePath, "utf8");
if (/(?:from|import)\s*["\']node:/.test(runtimeSource) || /\bprocess\.(?:getBuiltinModule|env|argv|cwd|execPath)\b|\bBuffer\.from\b|\bglobalThis\.(?:Buffer|process|document)\b|\bnew DOMParser\b/.test(runtimeSource)) {
  throw new Error("Runtime bundle contains a Node-only or DOM-only dependency.");
}

if (runtime.mikuMd2PptxMetadata?.productName !== "miku-md2pptx") {
  throw new Error("Runtime bundle metadata has an unexpected productName.");
}
if (runtime.mikuMd2PptxMetadata?.artifactRole !== "markdown-to-pptx-runtime") {
  throw new Error("Runtime bundle metadata has an unexpected artifactRole.");
}

for (const exportName of ["markdownToPptx", "markdownToPptxResult", "markdownToSlides"]) {
  if (typeof runtime[exportName] !== "function") {
    throw new Error(`Runtime bundle does not export function: ${exportName}`);
  }
}

const result = await runtime.markdownToPptxResult("# Smoke\n\nRuntime bundle check");
if (!(result.pptx instanceof Uint8Array) || result.pptx.byteLength === 0) {
  throw new Error("Runtime bundle did not generate a non-empty PPTX buffer.");
}
if (!Array.isArray(result.diagnostics)) {
  throw new Error("Runtime bundle result diagnostics is not an array.");
}

console.log("[smoke:runtime] ok");
