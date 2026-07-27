import { describe, expect, it } from "vitest";
import { markdownToSlides, markdownToPptx, markdownToPptxResult, mikuMd2PptxMetadata } from "../dist/core.js";
import { unzipStoredEntries, zipCompressionMethods } from "./helpers/zip.js";

const onePixelPng = new Uint8Array([
  137, 80, 78, 71, 13, 10, 26, 10, 0, 0, 0, 13, 73, 72, 68, 82,
  0, 0, 0, 1, 0, 0, 0, 1, 8, 6, 0, 0, 0, 31, 21, 196,
  137, 0, 0, 0, 13, 73, 68, 65, 84, 120, 156, 99, 248, 255, 255,
  63, 0, 5, 254, 2, 254, 167, 53, 129, 132, 0, 0, 0, 0,
  73, 69, 78, 68, 174, 66, 96, 130
]);

describe("miku-md2pptx core", () => {
  it("exports runtime metadata for downstream adapters", () => {
    expect(mikuMd2PptxMetadata).toEqual({
      productName: "miku-md2pptx",
      artifactRole: "markdown-to-pptx-runtime",
      primaryInput: "markdown",
      primaryOutput: "pptx",
      coreApi: ["markdownToSlides", "markdownToPptx", "markdownToPptxResult"]
    });
  });

  it("splits Markdown into slides at level 1 and 2 headings", () => {
    const slides = markdownToSlides("# One\n\nBody\n\n## Two\n\n- Item\n");

    expect(slides).toEqual([
      { title: "One", blocks: [{ kind: "text", text: "Body", runs: [{ text: "Body" }] }], notes: [] },
      { title: "Two", blocks: [{ kind: "text", text: "- Item", runs: [{ text: "- " }, { text: "Item" }] }], notes: [] }
    ]);
  });

  it("creates a pptx package with presentation and slide parts", () => {
    const pptx = markdownToPptx("# Deck\n\n## Slide\n\nBody");
    const entries = unzipStoredEntries(pptx);
    expect(new Set(zipCompressionMethods(pptx))).toEqual(new Set([8]));

    expect(entries.get("[Content_Types].xml")).toContain("presentationml.presentation.main+xml");
    expect(entries.get("ppt/_rels/presentation.xml.rels")).toContain("slides/slide1.xml");
    expect(entries.get("ppt/slideLayouts/slideLayout1.xml")).toContain('type="obj"');
    expect(entries.get("ppt/slideLayouts/slideLayout1.xml")).not.toContain('type="titleAndContent"');
    expect(entries.get("ppt/slides/slide1.xml")).toContain("Deck");
    expect(entries.get("ppt/slides/slide2.xml")).toContain("Body");
  });

  it("preserves supplementary Unicode and removes invalid XML characters", () => {
    const entries = unzipStoredEntries(markdownToPptx("# Deck 😀\n\nBody\u0001 text \uD800"));
    const slideXml = entries.get("ppt/slides/slide1.xml");

    expect(slideXml).toContain("Deck 😀");
    expect(slideXml).toContain("Body text ");
    expect(slideXml).not.toContain("\u0001");
    expect(slideXml).not.toContain("\uFFFD");
  });

  it("writes Markdown tables as native PowerPoint table parts", () => {
    const entries = unzipStoredEntries(markdownToPptx("# Deck\n\n| A | B |\n| --- | --- |\n| 1 | 2 |"));

    expect(entries.get("ppt/slides/slide1.xml")).toContain("<p:graphicFrame>");
    expect(entries.get("ppt/slides/slide1.xml")).toContain("<a:tbl>");
    expect(entries.get("ppt/slides/slide1.xml")).toContain("<a:t>A</a:t>");
  });

  it("writes Markdown links as PowerPoint hyperlink relationships", () => {
    const entries = unzipStoredEntries(markdownToPptx("# Deck\n\nSee [example](https://example.com/)."));

    expect(entries.get("ppt/slides/slide1.xml")).toContain("<a:hlinkClick r:id=\"rId2\"/>");
    expect(entries.get("ppt/slides/_rels/slide1.xml.rels")).toContain("Target=\"https://example.com/\"");
    expect(entries.get("ppt/slides/_rels/slide1.xml.rels")).toContain("TargetMode=\"External\"");
  });

  it("embeds resolved local Markdown images as PowerPoint picture parts", () => {
    const entries = unzipStoredEntries(markdownToPptx("# Deck\n\n![Chart](assets/chart.png)", {
      resolveImage(url) {
        if (url === "assets/chart.png") {
          return { bytes: onePixelPng, extension: "png" };
        }
        return undefined;
      }
    }));

    expect(entries.get("[Content_Types].xml")).toContain('Extension="png" ContentType="image/png"');
    expect(entries.get("ppt/slides/slide1.xml")).toContain("<p:pic>");
    expect(entries.get("ppt/slides/slide1.xml")).toContain('descr="Chart"');
    expect(entries.get("ppt/slides/_rels/slide1.xml.rels")).toContain("Target=\"../media/image1.png\"");
    expect(entries.has("ppt/media/image1.png")).toBe(true);
  });

  it("writes speaker notes comments as PowerPoint notes slides", () => {
    const entries = unzipStoredEntries(markdownToPptx("# Deck\n\nBody\n\n<!-- speaker-notes: Presenter reminder -->"));

    expect(entries.get("[Content_Types].xml")).toContain("presentationml.notesSlide+xml");
    expect(entries.get("ppt/slides/_rels/slide1.xml.rels")).toContain("relationships/notesSlide");
    expect(entries.get("ppt/notesSlides/notesSlide1.xml")).toContain("Presenter reminder");
    expect(entries.get("ppt/notesSlides/notesSlide1.xml")).toContain('ph type="sldImg"');
  });

  it("reports diagnostics for skipped images and raw HTML-like text", () => {
    const result = markdownToPptxResult("# Deck\n\n![Missing](assets/missing.png)\n\n<span>raw</span>");

    expect(result.pptx.byteLength).toBeGreaterThan(0);
    expect(result.diagnostics).toEqual([
      expect.objectContaining({
        severity: "warning",
        code: "possible-raw-html-text"
      }),
      expect.objectContaining({
        severity: "warning",
        code: "skipped-image",
        message: "Markdown image was not embedded: assets/missing.png"
      })
    ]);
  });

  it("uses a template title/body layout without copying existing template slides", () => {
    const templatePptx = markdownToPptx("# Template cover\n\nTemplate-only body");
    const result = markdownToPptxResult("# Generated deck\n\nGenerated body", { templatePptx });
    const entries = unzipStoredEntries(result.pptx);

    expect(new Set(zipCompressionMethods(result.pptx))).toEqual(new Set([8]));
    expect(result.diagnostics).toContainEqual(expect.objectContaining({
      severity: "info",
      code: "template-layout-selected"
    }));
    expect(entries.get("ppt/slides/slide1.xml")).toContain("Generated deck");
    expect(entries.get("ppt/slides/slide1.xml")).toContain("Generated body");
    expect(entries.get("ppt/slides/slide1.xml")).not.toContain(' sz="2400"');
    expect(entries.get("ppt/slides/slide1.xml")).not.toContain(' sz="1800"');
    expect(entries.get("ppt/slides/slide1.xml")).not.toContain("Template-only body");
    expect(entries.get("ppt/slides/_rels/slide1.xml.rels")).toContain("Target=\"../slideLayouts/slideLayout1.xml\"");
    expect(entries.get("ppt/presentation.xml")).not.toContain("Template cover");
  });
});
