# miku-md2pptx

`miku-md2pptx` converts Markdown files into editable PowerPoint `.pptx` decks.

It is a local tool. Your Markdown file is processed on your machine and is not
uploaded to a server.

The conversion goal is practical slide structure, not pixel-perfect PowerPoint
layout.

## What It Converts

The first implementation supports:

- heading level 1 and 2 sections as slides
- paragraphs
- bullet and numbered lists as slide text
- fenced code blocks as slide text
- Markdown links as external PowerPoint hyperlinks
- local PNG, JPEG, and GIF images referenced by relative Markdown paths
- speaker notes using `<!-- speaker-notes: ... -->` HTML comments
- Markdown tables as simple native PowerPoint tables
- PowerPoint template design reuse through `--template <pptx>`
- conversion diagnostics for skipped images and HTML-like text warnings

The generated slide titles, body text, bullet lists, external hyperlinks,
local image assets, speaker notes, and simple tables are covered by a
compatibility test that reads the generated `.pptx` back through
`miku-pptx2md`.

Generated decks in the supported scope have also been checked in Microsoft
PowerPoint for macOS. A representative deck containing titles, body text,
lists, a simple table, a local image, an external link, and speaker notes opens
without PowerPoint repair.

Known limitations:

- remote image URLs, absolute image paths, missing image files, and unsupported
  image formats are skipped with warnings
- raw HTML is not fully converted
- detailed theme editing, template placeholder replacement, multi-layout
  selection, speaker notes layout, and slide master customization are not
  supported yet

## CLI Use

Install dependencies once:

```bash
npm install
```

Convert a Markdown file:

```bash
npm run cli -- ./sample.md --out ./sample.pptx
```

Override the presentation title:

```bash
npm run cli -- ./sample.md --out ./sample.pptx --title "Project brief"
```

Use a PowerPoint template:

```bash
npm run cli -- ./sample.md --out ./sample.pptx --template ./template.pptx
```

`--template` reads design information from the template PPTX and uses the first
slide layout that has title and body/content placeholders. Generated slides
reference that layout, so they inherit its slide master and theme. The output
contains only slides generated from the Markdown input; existing slides in the
template are not copied.

Template-based generation is structural rather than pixel-perfect. It reuses
slide size, theme, masters, layouts, and placeholder geometry where practical,
but it does not edit existing template slides or perform full PowerPoint
layout flow. Tables, images, dense text, and final visual polish may still need
manual adjustment in PowerPoint.

Show help or version:

```bash
npm run cli -- --help
npm run cli -- --version
```

## Current Status

This repository is in first-cut development. It creates `.pptx` files from
Markdown headings and text-oriented document blocks. The first reverse
compatibility fixture verifies `miku-md2pptx -> miku-pptx2md` extraction for
slide order, slide titles, body text, bullet lists, external hyperlinks, local
image assets, speaker notes, and simple tables. The CLI reports conversion
warnings such as skipped images on stderr. PowerPoint repair checks currently
cover the same supported structural scope. Template mode has been smoke-tested
with a local PowerPoint template under `workplace/`; generated slides reference
the selected template layout, and existing template slides are not copied.

Related miku-soft applications:

- `miku-md2docx`: Markdown to Word `.docx`
- `miku-md2xlsx`: Markdown to Excel `.xlsx`
- `miku-pptx2md`: PowerPoint `.pptx` to Markdown-oriented artifacts

See [TODO.md](./TODO.md) for follow-up work.

Developer notes are in [docs/development.md](./docs/development.md).
Shared miku-soft reference information is in
[docs/miku-soft-reference.md](./docs/miku-soft-reference.md).

## Development Notes

Build and test:

```bash
npm run build
npm test
```

Build and smoke-test the single-file CLI bundle:

```bash
npm run build:bundle
npm run smoke:bundle
npm run smoke:runtime
```

`workplace/` is a local scratch area for sister repository checkouts, generated
verification files, and temporary artifacts. Only `workplace/.gitkeep` is
tracked.

Generated build outputs under `dist/`, `bundle/`, and `coverage/` are ignored.

GitHub Release assets are built when a `v*` GitHub Release is published. The
release workflow attaches the CLI bundle, importable runtime bundle, and source
archive for the release tag.

## License

Apache License 2.0.

See [LICENSE](./LICENSE).
