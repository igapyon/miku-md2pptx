export interface MarkdownToPptxOptions {
  title?: string;
  sourcePath?: string;
  templatePptx?: Uint8Array;
  resolveImage?: (url: string, sourcePath?: string) => ResolvedImage | undefined;
}

export interface MarkdownToPptxResult {
  pptx: Uint8Array;
  diagnostics: MarkdownToPptxDiagnostic[];
}

export interface MarkdownToPptxDiagnostic {
  severity: "info" | "warning" | "error";
  code: string;
  message: string;
  source?: string;
}

export interface SlideModel {
  title: string;
  blocks: SlideBlock[];
  notes: TextRun[][];
}

export type SlideBlock = TextSlideBlock | TableSlideBlock | ImageSlideBlock;

export interface TextSlideBlock {
  kind: "text";
  text: string;
  runs: TextRun[];
}

export interface TableSlideBlock {
  kind: "table";
  rows: TableCell[][];
}

export interface TableCell {
  text: string;
  runs: TextRun[];
}

export interface TextRun {
  text: string;
  href?: string;
}

export interface ImageSlideBlock {
  kind: "image";
  altText: string;
  url: string;
}

export interface ResolvedImage {
  bytes: Uint8Array;
  extension: "png" | "jpg" | "jpeg" | "gif";
}
