import type { MarkdownToPptxDiagnostic, MarkdownToPptxOptions, SlideModel } from "./types.ts";

export function collectDiagnostics(slides: SlideModel[], options: MarkdownToPptxOptions): MarkdownToPptxDiagnostic[] {
  const diagnostics: MarkdownToPptxDiagnostic[] = [];
  for (const slide of slides) {
    for (const block of slide.blocks) {
      if (block.kind === "text" && block.text.includes("<") && block.text.includes(">")) {
        diagnostics.push({
          severity: "warning",
          code: "possible-raw-html-text",
          message: "Raw HTML-like text was emitted as plain slide text.",
          ...(options.sourcePath ? { source: options.sourcePath } : {})
        });
      }
    }
  }
  return diagnostics;
}
