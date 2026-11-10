import { mkdir, readFile, writeFile } from "node:fs/promises";
import { build } from "esbuild";

const outputDirectory = "workplace/browser-tests";
await mkdir(outputDirectory, { recursive: true });
await build({ entryPoints: ["bundle/miku-md2pptx-runtime.mjs"], outfile: `${outputDirectory}/miku-md2pptx-runtime.iife.js`, bundle: true, format: "iife", globalName: "mikuMd2pptxRuntime", platform: "neutral", conditions: [], mainFields: ["module", "main"], target: "es2022" });
await writeFile(`${outputDirectory}/index.html`, `<!doctype html><html lang="en"><meta charset="utf-8"><title>miku-md2pptx browser verification</title><script src="./miku-md2pptx-runtime.iife.js"></script><script>window.runtimeReady = Boolean(window.mikuMd2pptxRuntime?.markdownToPptxResult);</script><body>Offline runtime harness</body></html>`);
const iifeSource = await readFile(`${outputDirectory}/miku-md2pptx-runtime.iife.js`, "utf8");
await writeFile(`${outputDirectory}/miku-md2pptx-worker.js`, `${iifeSource}
self.onmessage = async () => { try { const bytes = await self.mikuMd2pptxRuntime.markdownToPptx("# Worker deck\\n\\nWorker generated this slide."); let binary = ""; for (let i = 0; i < bytes.length; i += 0x8000) binary += String.fromCharCode(...bytes.subarray(i, i + 0x8000)); self.postMessage({ pptxBase64: btoa(binary) }); } catch (error) { self.postMessage({ error: error instanceof Error ? error.message : String(error) }); } };
`);
console.log(`Generated browser harness in ${outputDirectory}.`);
