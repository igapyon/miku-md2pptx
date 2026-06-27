---
purpose: ai-agent-goal
read_when:
  - before_starting_work
  - before_finishing_work
  - when_scope_is_unclear
update_when:
  - goal_changes
  - done_conditions_change
  - stop_conditions_change
---

# Goal

This file defines what the AI agent is trying to accomplish.
Read this before starting work, before deciding that work is complete, and whenever scope becomes unclear.

## Objective

Develop `miku-md2pptx` as a miku-soft Node.js main application that converts
Markdown into `.pptx` files, prioritizing the PowerPoint structures that
`miku-pptx2md` can extract reliably.

The current product direction is compatibility-centered:

- generate local `.pptx` files from Markdown without server dependency
- keep conversion semantics in the TypeScript product core
- treat `miku-pptx2md` as the reverse-direction compatibility reference
- prefer slide order, titles, body text, lists, simple tables, images, links,
  speaker notes, and diagnostics over decorative PowerPoint design
- avoid claiming round-trip behavior until fixtures prove
  `miku-md2pptx -> miku-pptx2md` expectations

## Done

- `miku-md2pptx` has a working CLI that converts Markdown input to `.pptx`
  output.
- Generated `.pptx` files open in common presentation software or any
  incompatibility is documented with a concrete blocker.
- The core supported scope is aligned with `miku-pptx2md` extraction behavior
  for slide order, titles, body text, lists, simple tables, images, links,
  speaker notes, and diagnostics.
- Compatibility fixtures verify important `miku-md2pptx -> miku-pptx2md`
  expectations.
- README, TODO, and relevant docs describe the supported scope, limitations,
  and reverse-direction compatibility policy.
- Relevant build, test, CLI smoke, bundle smoke, and generated package checks
  pass.

## Stop

- Stop and ask the user if the desired compatibility target conflicts with
  what `miku-pptx2md` currently extracts.
- Stop and ask the user before broadening the product into a visual design or
  presentation-authoring tool beyond the `miku-pptx2md`-centered scope.
- Stop and ask the user before changing repository direction toward Web App,
  Java, Agent Skill, or MCP companion surfaces.
- Stop and ask the user if validating generated `.pptx` files requires manual
  GUI checks that cannot be performed from the current environment.
- `TODO.md` の `Retry Log` に同じ原因の失敗が3回記録された
