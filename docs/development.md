# miku-md2pptx Development

This project was initialized as a miku-soft Node.js main application.

Checked references:

- `igapyon-miku-soft-developer` installed skill on 2026-06-27
- main workflow: `references/10-node-app-workflow.md`
- main design: `references/miku-soft-basic/miku-soft-10-mainapp-design.md`
- sister repositories:
  - `/Users/igapyon/Documents/git/miku-md2docx`
  - `/Users/igapyon/Documents/git/miku-md2xlsx`
- reverse-direction PowerPoint reference:
  - `/Users/igapyon/Documents/git/miku-pptx2md`

Adopted repository-shape decisions from the sister repositories:

- TypeScript product core under `src/ts/`
- Node CLI wrapper under `scripts/`
- `dist/core.js` build output consumed by the CLI wrapper
- Vitest tests against core behavior and CLI subprocess behavior
- local `workplace/` scratch area tracked only by `workplace/.gitkeep`
- generated outputs under `dist/`, `bundle/`, and `coverage/`

Reference decisions from `miku-pptx2md`:

- Treat PowerPoint conversion as structure-first rather than visual-fidelity
  reconstruction.
- Keep PowerPoint-specific conversion policy in this repository's product core.
- Use the reverse converter as a compatibility reference for slide order,
  title placeholders, body placeholders, simple tables, speaker notes, images,
  links, and diagnostics.
- Generate title and body placeholders so `miku-pptx2md` treats slide headings
  as slide titles and ordinary body text as paragraphs rather than generic
  shape text.
- Generate Markdown tables as native PowerPoint table graphic frames so
  `miku-pptx2md` can recover Markdown table rows.
- Generate Markdown links as external PowerPoint hyperlink relationships so
  `miku-pptx2md` can recover Markdown link syntax and hyperlink counts.
- Embed local relative PNG, JPEG, and GIF Markdown image references under
  `ppt/media/` and generate picture relationships so `miku-pptx2md` can
  recover image asset summaries.
- Generate speaker notes from `<!-- speaker-notes: ... -->` comments as
  notesSlide parts so `miku-pptx2md` can recover the `### Speaker Notes`
  section and notes summary count.
- Return conversion diagnostics from the core result and print CLI warnings for
  skipped images and HTML-like text.
- For `--template`, read the template PPTX as design input only. Existing
  template slides are not copied into the generated deck.
- Select the first template slide layout with title and body/content
  placeholders. Generated slides reference that layout, preserving its slide
  master and theme relationship. If no title/body layout is found, fall back to
  the first title-only layout; if no title layout is found, use the default
  generated layout and report a diagnostic.
- Template layout selection failure is not the same as template file failure.
  If the template file cannot be read as a PPTX package, conversion fails. If
  the PPTX can be read but no usable title/body layout is found, conversion can
  fall back to the built-in generated layout with a diagnostic.
- Generated template-mode slides reuse the actual title and body placeholder
  tags from the selected layout when possible. This matters for templates whose
  body placeholder is identified by `idx` rather than `type="body"`.
- When a template is used, generated text runs do not set explicit font sizes;
  PowerPoint resolves text size through the selected layout, slide master, and
  theme. Non-template generation keeps the built-in 24pt title and 18pt body
  sizes.
- Template-mode tables use the selected body placeholder geometry when it can
  be resolved from the slide layout or parent slide master. Table flow is still
  simple: tables are positioned inside that body area with a small offset after
  preceding text blocks rather than using full PowerPoint text-flow behavior.
- Template mode deliberately stops short of full visual layout. Tables, images,
  and dense content may need manual positioning in PowerPoint after generation.
- Do not claim full round-trip behavior until explicit fixtures and tests cover
  `miku-md2pptx -> miku-pptx2md` expectations.

Current compatibility fixture:

- `tests/md2pptx-pptx2md-compat.test.js` generates a `.pptx`, reads it through
  local `../miku-pptx2md`, and verifies slide titles, body text, list items,
  external hyperlink extraction, image asset extraction, speaker notes
  extraction, simple table extraction, and zero reverse-converter errors.

Presentation software check:

- Microsoft PowerPoint for macOS opens generated text-only, link, image, table,
  speaker-note, and representative combined structural decks without repair.
- Microsoft PowerPoint for macOS was also used for a local template experiment
  with `workplace/サンプル.pptx`; generated slides used the selected template
  layout and preserved the "existing template slides are not copied" behavior.
- LibreOffice is not installed in the current environment.
- Keynote is installed, but automated AppleScript validation was not reliable
  enough to use as completion evidence.

Rejected or deferred decisions:

- Web App surface is not initialized in this repository.
- Java, Agent Skill, and MCP companion surfaces are not initialized.
- GitHub Actions release workflow follows the sister `miku-md2xlsx` release
  CLI/runtime bundle shape. It runs when a `v*` GitHub Release is published
  and uploads `bundle/miku-md2pptx.mjs`,
  `bundle/miku-md2pptx-runtime.mjs`, and
  `bundle/miku-md2pptx-sources.tgz` as versioned GitHub Release assets.
