import { mkdir, readFile, writeFile } from "node:fs/promises";
import { readFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { markdownToPptxResult } from "../../dist/core.js";
import packageJson from "../../package.json" with { type: "json" };

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const rootDir = path.resolve(__dirname, "../..");

function usage() {
  return `miku-md2pptx converts a Markdown file into a PowerPoint .pptx deck.

Usage:
  miku-md2pptx <input.md> --out <output.pptx> [--template <template.pptx>]
  miku-md2pptx --help
  miku-md2pptx --version

Options:
  --out <path>             Output .pptx path.
  --template <path>        Use a PowerPoint template's design information and
                           first title+body slide layout. Existing template
                           slides are not copied.
  --title <text>           Override the generated presentation title.
  --help                   Show this help.
  --version                Show the package version.

Template behavior:
  --template reads slide size, theme, slide masters, slide layouts, and related
  design parts from the template PPTX.
  Generated output contains only slides created from the Markdown input.
  Existing slides in the template are not copied, prepended, appended, or
  edited.
  Generated slides reference the first title+body/content layout found in the
  template. If no such layout is found, the converter tries a title-only layout,
  then falls back to the built-in generated layout with a diagnostic.
  If the template PPTX cannot be read, conversion fails instead of silently
  falling back.
  Template mode is structural, not pixel-perfect. Tables, images, and dense
  content may need final positioning in PowerPoint.

Markdown handling notes:
  Heading level 1 and 2 blocks start new slides.
  Paragraphs, lists, code blocks, and tables become simple editable slide text.
  The first implementation prioritizes structure and local generation over
  pixel-perfect PowerPoint layout.

Examples:
  npm run cli -- ./sample.md --out ./sample.pptx
  npm run cli -- ./sample.md --out ./sample.pptx --template ./template.pptx
  npm run cli -- ./sample.md --out ./sample.pptx --title "Project brief"
`;
}

function parseArgs(args) {
  const options = { input: undefined, out: undefined, title: undefined, template: undefined };
  for (let i = 0; i < args.length; i += 1) {
    const arg = args[i];
    if (arg === "--help" || arg === "-h") {
      return { help: true };
    }
    if (arg === "--version") {
      return { version: true };
    }
    if (arg === "--out") {
      options.out = args[++i];
      if (!options.out) {
        throw new Error("--out requires a path.");
      }
      continue;
    }
    if (arg === "--title") {
      options.title = args[++i];
      if (!options.title) {
        throw new Error("--title requires text.");
      }
      continue;
    }
    if (arg === "--template") {
      options.template = args[++i];
      if (!options.template) {
        throw new Error("--template requires a .pptx path.");
      }
      continue;
    }
    if (arg.startsWith("--")) {
      throw new Error(`Unknown option: ${arg}`);
    }
    if (options.input) {
      throw new Error(`Unexpected argument: ${arg}`);
    }
    options.input = arg;
  }
  if (!options.input) {
    throw new Error("Input Markdown path is required. Use --help for usage.");
  }
  if (!options.out) {
    throw new Error("--out is required. Use --help for usage.");
  }
  return options;
}

function resolveLocalImage(url, sourcePath) {
  if (!sourcePath || /^(?:[a-z][a-z0-9+.-]*:|\/)/i.test(url)) {
    return undefined;
  }
  const cleanUrl = url.split("#")[0].split("?")[0];
  const extension = path.extname(cleanUrl).slice(1).toLowerCase();
  if (!["png", "jpg", "jpeg", "gif"].includes(extension)) {
    return undefined;
  }
  try {
    const imagePath = path.resolve(path.dirname(sourcePath), decodeURIComponent(cleanUrl));
    const sourceDir = path.resolve(path.dirname(sourcePath));
    const relative = path.relative(sourceDir, imagePath);
    if (relative.startsWith("..") || path.isAbsolute(relative)) {
      return undefined;
    }
    return {
      bytes: readFileSync(imagePath),
      extension
    };
  } catch {
    return undefined;
  }
}

export async function main(args) {
  const options = parseArgs(args);
  if (options.help) {
    process.stdout.write(usage());
    return;
  }
  if (options.version) {
    process.stdout.write(`${packageJson.version}\n`);
    return;
  }

  const inputPath = path.resolve(rootDir, options.input);
  const outputPath = path.resolve(rootDir, options.out);
  const templatePath = options.template ? path.resolve(rootDir, options.template) : undefined;
  const markdown = await readFile(inputPath, "utf8");
  const result = markdownToPptxResult(markdown, {
    title: options.title,
    sourcePath: inputPath,
    ...(templatePath ? { templatePptx: await readFile(templatePath) } : {}),
    resolveImage: resolveLocalImage
  });
  await mkdir(path.dirname(outputPath), { recursive: true });
  await writeFile(outputPath, result.pptx);
  for (const diagnostic of result.diagnostics) {
    process.stderr.write(`${diagnostic.severity}: ${diagnostic.code}: ${diagnostic.message}\n`);
  }
  process.stdout.write(`Wrote ${path.relative(rootDir, outputPath)}\n`);
}
