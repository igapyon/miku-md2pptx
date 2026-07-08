import type {
  ImageSlideBlock,
  MarkdownToPptxDiagnostic,
  MarkdownToPptxOptions,
  ResolvedImage
} from "./types.ts";

export interface MediaEntry {
  path: string;
  data: Uint8Array;
}

export interface MediaManager {
  entries: MediaEntry[];
  addImage(image: ImageSlideBlock): { target: string; packagePath: string; extension: string } | undefined;
}

function normalizeImageExtension(image: ResolvedImage): "png" | "jpg" | "gif" {
  if (image.extension === "jpeg" || image.extension === "jpg") {
    return "jpg";
  }
  if (image.extension === "gif") {
    return "gif";
  }
  return "png";
}

export function createMediaManager(options: MarkdownToPptxOptions, diagnostics: MarkdownToPptxDiagnostic[], startImageIndex = 1): MediaManager {
  const entries: MediaEntry[] = [];
  let nextImageIndex = startImageIndex;
  return {
    entries,
    addImage(image: ImageSlideBlock) {
      const resolved = options.resolveImage?.(image.url, options.sourcePath);
      if (!resolved) {
        diagnostics.push({
          severity: "warning",
          code: "skipped-image",
          message: `Markdown image was not embedded: ${image.url}`,
          ...(options.sourcePath ? { source: options.sourcePath } : {})
        });
        return undefined;
      }
      const extension = normalizeImageExtension(resolved);
      const fileName = `image${nextImageIndex}.${extension}`;
      nextImageIndex += 1;
      const packagePath = `ppt/media/${fileName}`;
      entries.push({ path: packagePath, data: resolved.bytes });
      return {
        target: `../media/${fileName}`,
        packagePath,
        extension
      };
    }
  };
}
