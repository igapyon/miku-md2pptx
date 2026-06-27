import { unified } from "unified";
import remarkGfm from "remark-gfm";
import remarkParse from "remark-parse";
import type { TextRun } from "./types.ts";

export function parseMarkdown(markdown: string): any {
  return unified().use(remarkParse).use(remarkGfm).parse(markdown) as any;
}

export function extractText(node: any): string {
  if (!node) {
    return "";
  }
  if (typeof node.value === "string") {
    return node.value;
  }
  if (Array.isArray(node.children)) {
    return node.children.map((child: any) => extractText(child)).join(node.type === "paragraph" ? "" : " ");
  }
  return "";
}

export function extractTextRuns(node: any): TextRun[] {
  if (!node) {
    return [];
  }
  if (node.type === "link" && typeof node.url === "string") {
    const text = extractText(node);
    return text ? [{ text, href: node.url }] : [];
  }
  if (typeof node.value === "string") {
    return node.value ? [{ text: node.value }] : [];
  }
  if (Array.isArray(node.children)) {
    const runs: TextRun[] = [];
    for (const [index, child] of node.children.entries()) {
      if (index > 0 && node.type !== "paragraph" && runs.length > 0) {
        runs.push({ text: " " });
      }
      runs.push(...extractTextRuns(child));
    }
    return runs;
  }
  return [];
}
