---
purpose: ai-agent-handoff
read_when:
  - before_resuming_work
  - before_handing_off_work
  - when_context_is_missing
update_when:
  - work_is_paused
  - handoff_summary_changes
  - verification_status_changes
---

# Handoff

This file summarizes the current working state for the next human or AI agent.
Keep it concise. Do not use this as a full work log or a replacement for `TODO.md` and `DECISIONS.md`.

## Current State

- `miku-md2pptx` is a miku-soft TypeScript / Node.js main application.
- The current product direction is compatibility-centered Markdown to
  PowerPoint conversion, using `miku-pptx2md` as the reverse-direction
  extraction reference.
- The repository has working source, tests, docs, and bundle scripts.
- A behavior-preserving refactoring has now been performed: OOXML helper
  functions were moved to `src/ts/ooxml.ts`, and static PPTX part generators
  were moved to `src/ts/pptx-static-parts.ts`.

## Next Action

- Review whether deeper slide XML rendering extraction is worth doing later.
  Do not start broader behavior changes unless the user asks.

## Relevant Files

- `GOAL.md`: objective, done conditions, and stop conditions.
- `TODO.md`: active project TODOs and AI agent current tasks.
- `DECISIONS.md`: compatibility and state-management decisions.
- `README.md`: supported scope, CLI usage, and current status.
- `docs/development.md`: miku-soft references and repository-shape decisions.
- `src/ts/`: TypeScript product core.
- `src/ts/ooxml.ts`: OOXML namespaces, XML escaping, and relationship XML.
- `src/ts/pptx-static-parts.ts`: static PPTX master/layout/theme/support
  parts.
- `scripts/`: CLI and build entrypoints.
- `tests/`: behavior, CLI, and compatibility fixtures.

## Watch Outs

- Do not broaden the tool into a visual design or presentation-authoring tool
  without explicit user confirmation.
- Do not change the Web App, Java, Agent Skill, or MCP direction without
  explicit user confirmation.
- Preserve unrelated user changes and keep generated outputs under ignored
  paths such as `dist/`, `bundle/`, and `workplace/`.

## Last Verification

- `npm test` passed: 3 test files, 12 tests.
- `npm run smoke:version` passed and printed `0.2.0`.
- `npm run build:bundle` passed.
- `npm run smoke:bundle` passed.
