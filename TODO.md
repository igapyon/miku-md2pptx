# TODO

- Improve PowerPoint layout fidelity after the first structural conversion is stable.
- Align future image, table, speaker-note, and diagnostics behavior with
  `miku-pptx2md` where the reverse-direction contract is relevant.
- Cross-check generated `.pptx` files with Keynote and LibreOffice when those
  checks are needed for a broader compatibility claim.
- Decide whether a separated Web App repository is needed.

## Done

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
