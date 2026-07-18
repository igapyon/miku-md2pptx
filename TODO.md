# TODO

- Improve PowerPoint layout fidelity after the first structural conversion is stable.
- Align future image, table, speaker-note, and diagnostics behavior with
  `miku-pptx2md` where the reverse-direction contract is relevant.
- Cross-check generated `.pptx` files with Keynote and LibreOffice when those
  checks are needed for a broader compatibility claim.
- Decide whether a separated Web App repository is needed.

## AI Agent Current Tasks

This section tracks active work items for AI agents.
Update this section while working. Do not rewrite unrelated TODO items.

### Tasks

- [x] Inspect the existing Markdown-to-PPTX generation path for template layout
      support.
- [x] Add initial `--template <pptx>` support using the first title+body slide
      layout and without copying existing template slides.
- [x] Add focused core and CLI tests for template-based generation.
- [x] Update CLI help and Markdown docs with the template-mode execution
      contract for AI agents and human users.

### Blockers

- None.

### Retry Log

Use this section only when the same task or error is repeated.
If the same failure appears 3 times, stop and ask the user.

- None.

## Done

- Expanded CLI help and README documentation with an AI-readable execution
  contract covering path resolution, overwrite behavior, stdout/stderr, exit
  codes, image restrictions, and speaker-note syntax.
- Corrected the generated title-and-content slide layout type to the valid OOXML
  `obj` value so Microsoft PowerPoint does not repair the layout type.
- Adopted vendored `miku-ms-office-core` `v0.6.0` for low-level ZIP, OPC
  relationship, and XML sanitization plumbing, including supplementary Unicode
  coverage.
- Refactored `src/ts/core.ts` by moving OOXML relationship helpers to
  `src/ts/ooxml.ts` and static PPTX part generators to
  `src/ts/pptx-static-parts.ts`.
- Added a basic `miku-md2pptx -> miku-pptx2md` compatibility fixture for slide
  order, slide titles, body text, bullet lists, and simple tables.
- Generate Markdown tables as native PowerPoint table graphic frames instead
  of plain text rows.
- Generate Markdown links as external PowerPoint hyperlink relationships and
  verify `miku-pptx2md` hyperlink extraction.
- Embed local PNG, JPEG, and GIF Markdown image references and verify
  `miku-pptx2md` image asset extraction.
- Generate speaker notes from `<!-- speaker-notes: ... -->` comments and
  verify `miku-pptx2md` notes extraction.
- Return conversion diagnostics from the core and print CLI warnings for
  skipped images and HTML-like text.
- Verified generated `.pptx` files in Microsoft PowerPoint for macOS without
  repair for text-only, link, image, table, notes, and representative combined
  structural cases.
- Added `bundle/miku-md2pptx-runtime.mjs` generation and runtime smoke checks.
- Updated release asset workflow to run on GitHub Release publish and attach
  CLI bundle, runtime bundle, and source archive.
- Exported `mikuMd2PptxMetadata` from the core/runtime bundle and verified it
  in runtime smoke tests.
- Added initial template-based generation. `--template` uses template design
  information and the first title+body slide layout for Markdown-generated
  slides; existing template slides are not copied.
- Refined template-mode generation so generated slides reuse the selected
  layout's placeholder tags, template text sizing can be inherited, and native
  tables are positioned relative to the resolved body placeholder when
  possible. Detailed overlap and final visual polish remain a PowerPoint
  adjustment step.
