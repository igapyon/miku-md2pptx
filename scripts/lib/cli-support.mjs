import { mkdir, readFile, writeFile } from "node:fs/promises";
import { readFileSync } from "node:fs";
import path from "node:path";
import { markdownToPptxResult } from "../../dist/core.js";
import packageJson from "../../package.json" with { type: "json" };

export class CliUsageError extends Error {}

function usage() {
  const releaseAsset = `miku-md2pptx-${packageJson.version}.mjs`;
  return `miku-md2pptx converts a Markdown file into a PowerPoint .pptx deck.

Usage:
  node ${releaseAsset} <input.md> --out <output.pptx> [options]
  node ${releaseAsset} --help
  node ${releaseAsset} --version

Inputs:
  <input.md>              UTF-8 Markdown source file.
  --template <path>       Optional PowerPoint .pptx design source.

Outputs:
  stdout                  Human-readable success status, help, or version.
  stderr                  Usage errors, conversion diagnostics, and failures.
  <output.pptx>           Primary generated PowerPoint file.

Generated artifacts:
  A conversion writes only the .pptx path supplied with --out. Build outputs
  such as dist/ and bundle/ are development artifacts, not conversion output.

Overwrite behavior:
  The output parent directory is created when needed. An existing output file
  is replaced without prompting.

Machine-readable output contract:
  The generated .pptx file is the primary artifact. stdout is human-readable
  status text and is not a stable machine-readable data format.

Exit codes:
  0  Conversion succeeded, or --help/--version was shown.
  1  Input, output, template, or conversion processing failed.
  2  CLI usage error, such as missing arguments or an unknown option.

Options:
  --out <path>             Output .pptx path.
  --template <path>        Use a PowerPoint template's design information and
                           first title+body slide layout. Existing template
                           slides are not copied.
  --title <text>           Override the generated presentation title.
  --help, -h               Show this help.
  --version                Show the package version.

Execution contract:
  Input, output, template, and local image paths are processed locally. Relative
  CLI paths are resolved from the current working directory.
  On success, the command exits 0 and prints "Wrote <path>" to stdout.
  Conversion diagnostics use "<severity>: <code>: <message>" on stderr. A
  warning does not by itself make the command fail.

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
  Paragraphs, lists, fenced code blocks, and simple tables become editable
  PowerPoint content. Markdown links become external hyperlinks.
  Relative PNG, JPEG, and GIF images under the input file's directory can be
  embedded. Remote URLs, absolute paths, paths outside that directory, missing
  files, and unsupported formats are skipped with a warning.
  <!-- speaker-notes: text --> adds speaker notes to the current slide.
  The first implementation prioritizes structure and local generation over
  pixel-perfect PowerPoint layout.

Examples:
  node ${releaseAsset} input.md --out output.pptx
  node ${releaseAsset} input.md --out output.pptx --template template.pptx
`;
}

function usageError(message) {
  return new CliUsageError(message);
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
      if (!options.out || options.out.startsWith("-")) {
        throw usageError("--out requires a path.");
      }
      continue;
    }
    if (arg === "--title") {
      options.title = args[++i];
      if (!options.title || options.title.startsWith("-")) {
        throw usageError("--title requires text.");
      }
      continue;
    }
    if (arg === "--template") {
      options.template = args[++i];
      if (!options.template || options.template.startsWith("-")) {
        throw usageError("--template requires a .pptx path.");
      }
      continue;
    }
    if (arg.startsWith("-")) {
      throw usageError(`Unknown option: ${arg}`);
    }
    if (options.input) {
      throw usageError(`Unexpected argument: ${arg}`);
    }
    options.input = arg;
  }
  if (!options.input) {
    throw usageError("Input Markdown path is required. Use --help for usage.");
  }
  if (!options.out) {
    throw usageError("--out is required. Use --help for usage.");
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

  const workingDir = process.cwd();
  const inputPath = path.resolve(workingDir, options.input);
  const outputPath = path.resolve(workingDir, options.out);
  const templatePath = options.template ? path.resolve(workingDir, options.template) : undefined;
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
  process.stdout.write(`Wrote ${path.relative(workingDir, outputPath)}\n`);
}
