import { parseMarkdown, extractText, extractTextRuns } from "./markdown-parser.ts";
import type {
  MarkdownToPptxOptions,
  SlideBlock,
  SlideModel,
  TableCell,
  TextRun
} from "./types.ts";

function normalizeLine(value: string): string {
  return value.replace(/\s+/g, " ").trim();
}

function normalizeRuns(runs: TextRun[]): TextRun[] {
  const merged: TextRun[] = [];
  for (const run of runs) {
    const text = run.text.replace(/\s+/g, " ");
    if (!text) {
      continue;
    }
    const previous = merged[merged.length - 1];
    if (previous && previous.href === run.href) {
      previous.text += text;
    } else {
      merged.push(run.href ? { text, href: run.href } : { text });
    }
  }
  const joined = merged.map((run) => run.text).join("").trim();
  if (!joined) {
    return [];
  }
  let trimStart = merged.findIndex((run) => run.text.trimStart().length > 0);
  if (trimStart < 0) {
    return [];
  }
  merged.splice(0, trimStart);
  trimStart = 0;
  merged[trimStart].text = merged[trimStart].text.trimStart();
  merged[merged.length - 1].text = merged[merged.length - 1].text.trimEnd();
  return merged.filter((run) => run.text.length > 0);
}

function textBlockFromRuns(runs: TextRun[], prefix = ""): SlideBlock[] {
  const normalized = normalizeRuns(runs);
  if (normalized.length === 0) {
    return [];
  }
  const prefixedRuns = prefix ? [{ text: prefix }, ...normalized] : normalized;
  return [{ kind: "text", text: prefixedRuns.map((run) => run.text).join(""), runs: prefixedRuns }];
}

function parseSpeakerNotesHtml(value: string): TextRun[][] {
  const match = value.match(/^\s*<!--\s*speaker-notes(?::|\s)([\s\S]*?)-->\s*$/i);
  if (!match) {
    return [];
  }
  return match[1]
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean)
    .map((line) => [{ text: line }]);
}

function tableRows(node: any): TableCell[][] {
  const rows = Array.isArray(node.children) ? node.children : [];
  return rows
    .map((row: any) => {
      const cells = Array.isArray(row.children) ? row.children : [];
      return cells.map((cell: any) => {
        const runs = normalizeRuns(extractTextRuns(cell));
        return { text: runs.map((run) => run.text).join(""), runs };
      });
    })
    .filter((row: TableCell[]) => row.length > 0);
}

function listBlocks(node: any, depth = 0): SlideBlock[] {
  const blocks: SlideBlock[] = [];
  const items = Array.isArray(node.children) ? node.children : [];
  for (const item of items) {
    const children = Array.isArray(item.children) ? item.children : [];
    const runs: TextRun[] = [];
    for (const child of children) {
      if (child.type === "list") {
        continue;
      }
      runs.push(...extractTextRuns(child));
    }
    const normalized = normalizeRuns(runs);
    if (normalized.length > 0) {
      const prefix = `${"  ".repeat(depth)}- `;
      blocks.push({ kind: "text", text: `${prefix}${normalized.map((run) => run.text).join("")}`, runs: [{ text: prefix }, ...normalized] });
    }
    for (const child of children) {
      if (child.type === "list") {
        blocks.push(...listBlocks(child, depth + 1));
      }
    }
  }
  return blocks;
}

function nodeBlocks(node: any): SlideBlock[] {
  switch (node.type) {
    case "image":
      return [{
        kind: "image",
        altText: typeof node.alt === "string" ? node.alt : "",
        url: typeof node.url === "string" ? node.url : ""
      }];
    case "paragraph":
      if (Array.isArray(node.children) && node.children.length === 1 && node.children[0]?.type === "image") {
        return nodeBlocks(node.children[0]);
      }
      return textBlockFromRuns(extractTextRuns(node));
    case "list":
      return listBlocks(node);
    case "code":
      return String(node.value ?? "").split(/\r?\n/).map((line) => ({ kind: "text", text: `    ${line}`, runs: [{ text: `    ${line}` }] }));
    case "blockquote":
      return (node.children ?? []).flatMap((child: any) => nodeBlocks(child).map((block: SlideBlock) =>
        block.kind === "text" ? { kind: "text", text: `> ${block.text}`, runs: [{ text: "> " }, ...block.runs] } : block
      ));
    case "table":
      return [{ kind: "table", rows: tableRows(node) }];
    case "thematicBreak":
      return [{ kind: "text", text: "---", runs: [{ text: "---" }] }];
    default: {
      const text = normalizeLine(extractText(node));
      return text ? [{ kind: "text", text, runs: [{ text }] }] : [];
    }
  }
}

export function markdownToSlides(markdown: string, options: MarkdownToPptxOptions = {}): SlideModel[] {
  const tree = parseMarkdown(markdown);
  const slides: SlideModel[] = [];
  let current: SlideModel | undefined;

  function ensureSlide(): SlideModel {
    if (!current) {
      current = { title: options.title ?? "Markdown deck", blocks: [], notes: [] };
      slides.push(current);
    }
    return current;
  }

  for (const node of tree.children ?? []) {
    if (node.type === "heading" && (node.depth === 1 || node.depth === 2)) {
      current = { title: normalizeLine(extractText(node)) || "Untitled slide", blocks: [], notes: [] };
      slides.push(current);
      continue;
    }
    if (node.type === "html" && typeof node.value === "string") {
      const notes = parseSpeakerNotesHtml(node.value);
      if (notes.length > 0) {
        ensureSlide().notes.push(...notes);
        continue;
      }
    }
    const blocks = nodeBlocks(node);
    if (blocks.length > 0) {
      ensureSlide().blocks.push(...blocks);
    }
  }

  if (slides.length === 0) {
    slides.push({ title: options.title ?? "Markdown deck", blocks: [], notes: [] });
  }
  if (options.title && slides[0]) {
    slides[0].title = options.title;
  }
  return slides;
}
