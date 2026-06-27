---
purpose: ai-agent-decisions
read_when:
  - before_starting_work
  - when_making_decision
  - when_looping_or_repeating_work
update_when:
  - important_decision_is_made
  - option_is_rejected
  - work_is_deferred
---

# Decisions

This file records important decisions for the AI agent.
Read this before making or revisiting decisions, especially when the work seems to loop.

## 2026-06-27: Keep `miku-md2pptx` compatibility-centered

Reason:
The repository goal defines `miku-pptx2md` as the reverse-direction
compatibility reference. The supported scope prioritizes slide order, titles,
body text, lists, simple tables, images, links, speaker notes, and diagnostics
over decorative PowerPoint design.

Impact:
Refactoring should preserve product behavior and tests around this structural
conversion scope. Broader visual presentation-authoring behavior should not be
introduced without explicit user confirmation.

## 2026-06-27: Use lightweight repository-local state files

Reason:
`GOAL.md` and `TODO.md` already existed. The agent-state workflow benefits
from adding only the missing `DECISIONS.md` and `HANDOFF.md`, while preserving
the existing human-facing `TODO.md` content.

Impact:
Future agents should read `GOAL.md`, `TODO.md`, `DECISIONS.md`, and
`HANDOFF.md` before continuing substantial work.

## 2026-06-27: Limit refactoring to behavior-preserving extraction

Reason:
The immediate user concern was that implementation refactoring had not yet
been performed. The safest useful improvement was to reduce `src/ts/core.ts`
size by moving generic OOXML helpers and static PPTX part generators out of the
conversion assembly path without changing supported Markdown-to-PPTX behavior.

Impact:
`src/ts/core.ts` remains responsible for conversion orchestration and slide
assembly. `src/ts/ooxml.ts` owns XML escaping and relationship XML helpers.
`src/ts/pptx-static-parts.ts` owns master, layout, theme, notes-support, and
property part generators.
