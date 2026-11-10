import { spawnSync } from "node:child_process";
import { fileURLToPath, pathToFileURL } from "node:url";
import { expect, test } from "@playwright/test";
import { getZipTextEntry, readZipPackage } from "../../src/vendor/miku-ms-office-core-0.8.0.mjs";

const harnessPath = fileURLToPath(new URL("../../workplace/browser-tests/index.html", import.meta.url));
const templatePptx = createPythonTemplate();
const image = Buffer.from("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+jfS8AAAAASUVORK5CYII=", "base64");

test("offline runtime reuses a compressed template and runs in a Worker", async ({ browser }) => {
  const context = await browser.newContext();
  try {
    await context.route("**/*", async (route) => { const url = new URL(route.request().url()); if (url.protocol === "file:" || (url.protocol === "http:" && url.hostname === "127.0.0.1" && url.port === "4176")) await route.continue(); else await route.abort("blockedbyclient"); });
    const page = await context.newPage();
    await page.goto(pathToFileURL(harnessPath).href);
    await page.waitForFunction(() => window.runtimeReady === true);
    const result = await page.evaluate(async ({ templateBase64, imageBase64 }) => {
      const bytes = (value) => Uint8Array.from(atob(value), (character) => character.charCodeAt(0));
      const pending = window.mikuMd2pptxRuntime.markdownToPptxResult("# Browser deck\n\nThis is **bold text** with 日本語 😀 𠮷野家.\n\n[Open](https://example.com/)\n\n![Chart](assets/chart.png)\n\n<!-- speaker-notes: Presenter reminder -->", { templatePptx: bytes(templateBase64), resolveImage: () => ({ bytes: bytes(imageBase64), extension: "png" }) });
      const promiseReturned = typeof pending?.then === "function";
      const output = await pending;
      let binary = ""; for (let i = 0; i < output.pptx.length; i += 0x8000) binary += String.fromCharCode(...output.pptx.subarray(i, i + 0x8000));
      let damagedTemplateRejected = false; try { await window.mikuMd2pptxRuntime.markdownToPptx("fail", { templatePptx: new Uint8Array([1, 2, 3]) }); } catch { damagedTemplateRejected = true; }
      return { promiseReturned, pptxBase64: btoa(binary), diagnostics: output.diagnostics, damagedTemplateRejected };
    }, { templateBase64: templatePptx.toString("base64"), imageBase64: image.toString("base64") });
    expect(result.promiseReturned).toBe(true);
    expect(result.damagedTemplateRejected).toBe(true);
    expect(result.diagnostics).toContainEqual(expect.objectContaining({ code: "template-layout-selected" }));
    const verified = verifyWithPython(Buffer.from(result.pptxBase64, "base64"), image);
    expect(verified.slide).toContain("日本語");
    expect(verified.slide).toContain("𠮷野家");
    expect(verified.slide).toContain("bold text");
    expect(verified.slide).toContain("Chart");
    expect(verified.relationships).toContain("https://example.com/");
    expect(verified.relationships).toContain("slideLayout1.xml");
    expect(verified.layout).toContain("BrowserWideLayout");
    expect(verified.theme).toContain("BrowserTemplateTheme");
    expect(verified.presentation).toContain('cx="9144000" cy="5143500"');
    expect(verified.notes).toContain("Presenter reminder");

    const workerPage = await context.newPage(); await workerPage.goto("http://127.0.0.1:4176/health");
    const workerBase64 = await workerPage.evaluate(() => new Promise((resolve, reject) => { const worker = new Worker("/miku-md2pptx-worker.js"); worker.onmessage = (event) => { worker.terminate(); event.data.error ? reject(new Error(event.data.error)) : resolve(event.data.pptxBase64); }; worker.onerror = (event) => { worker.terminate(); reject(new Error(event.message)); }; worker.postMessage("run"); }));
    await workerPage.close();
    const workerPackage = await readZipPackage(new Uint8Array(Buffer.from(workerBase64, "base64")));
    expect(workerPackage.diagnostics).toEqual([]);
    expect(getZipTextEntry(workerPackage.entries, "ppt/slides/slide1.xml")).toContain("Worker generated this slide.");
  } finally { await context.close(); }
});

function createPythonTemplate() {
  const source = [
    "import io, sys, zipfile",
    "files = {",
    "'[Content_Types].xml': '<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Default Extension=\"xml\" ContentType=\"application/xml\"/><Override PartName=\"/ppt/presentation.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml\"/><Override PartName=\"/ppt/slideMasters/slideMaster1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml\"/><Override PartName=\"/ppt/slideLayouts/slideLayout1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml\"/><Override PartName=\"/ppt/theme/theme1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.theme+xml\"/></Types>',",
    "'_rels/.rels': '<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"ppt/presentation.xml\"/></Relationships>',",
    "'ppt/presentation.xml': '<p:presentation xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><p:sldIdLst></p:sldIdLst><p:sldSz cx=\"9144000\" cy=\"5143500\"/><p:notesSz cx=\"6858000\" cy=\"9144000\"/></p:presentation>',",
    "'ppt/_rels/presentation.xml.rels': '<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster\" Target=\"slideMasters/slideMaster1.xml\"/></Relationships>',",
    "'ppt/slideMasters/slideMaster1.xml': '<p:sldMaster xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\"><p:cSld><p:spTree></p:spTree></p:cSld></p:sldMaster>',",
    "'ppt/slideMasters/_rels/slideMaster1.xml.rels': '<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout\" Target=\"../slideLayouts/slideLayout1.xml\"/><Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme\" Target=\"../theme/theme1.xml\"/></Relationships>',",
    "'ppt/slideLayouts/slideLayout1.xml': '<p:sldLayout xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" type=\"obj\" name=\"BrowserWideLayout\"><p:cSld><p:spTree><p:sp><p:nvSpPr><p:nvPr><p:ph type=\"title\"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x=\"457200\" y=\"228600\"/><a:ext cx=\"8229600\" cy=\"914400\"/></a:xfrm></p:spPr></p:sp><p:sp><p:nvSpPr><p:nvPr><p:ph type=\"body\"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x=\"457200\" y=\"1371600\"/><a:ext cx=\"8229600\" cy=\"3200400\"/></a:xfrm></p:spPr></p:sp></p:spTree></p:cSld></p:sldLayout>',",
    "'ppt/theme/theme1.xml': '<a:theme xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" name=\"BrowserTemplateTheme\"><a:themeElements/></a:theme>',",
    "}",
    "output = io.BytesIO()",
    "with zipfile.ZipFile(output, 'w', compression=zipfile.ZIP_DEFLATED) as archive:",
    "  for name, content in files.items(): archive.writestr(name, content)",
    "sys.stdout.buffer.write(output.getvalue())"
  ].join("\n");
  const result = spawnSync("python3", ["-c", source], { encoding: "buffer" });
  if (result.status !== 0) throw new Error(`Unable to create PPTX template: ${result.stderr.toString()}`);
  return result.stdout;
}

function verifyWithPython(bytes, image) {
  const source = [
    "import base64, io, json, sys, zipfile",
    "archive = zipfile.ZipFile(io.BytesIO(sys.stdin.buffer.read()))",
    "assert archive.testzip() is None, 'ZIP CRC check failed'",
    "infos = archive.infolist()",
    "assert infos and all(info.compress_type == zipfile.ZIP_DEFLATED for info in infos), 'Unexpected compression method'",
    "slide = archive.read('ppt/slides/slide1.xml').decode('utf-8')",
    "rels = archive.read('ppt/slides/_rels/slide1.xml.rels').decode('utf-8')",
    "presentation = archive.read('ppt/presentation.xml').decode('utf-8')",
    "layout = archive.read('ppt/slideLayouts/slideLayout1.xml').decode('utf-8')",
    "theme = archive.read('ppt/theme/theme1.xml').decode('utf-8')",
    "notes = archive.read('ppt/notesSlides/notesSlide1.xml').decode('utf-8')",
    "assert archive.read('ppt/media/image1.png') == base64.b64decode(sys.argv[1]), 'Image bytes changed'",
    "print(json.dumps({'slide': slide, 'relationships': rels, 'presentation': presentation, 'layout': layout, 'theme': theme, 'notes': notes}, ensure_ascii=False))"
  ].join("\n");
  const result = spawnSync("python3", ["-c", source, image.toString("base64")], { input: bytes, encoding: "utf8" });
  if (result.status !== 0) throw new Error(`Independent PPTX verification failed: ${result.stderr}`);
  return JSON.parse(result.stdout);
}
