---
file: ./.agents/rules/doc/doc-style.md
name: Doc style - plain text only
description: doc-style - Plain-text doc style for agent-config markdown, covering bold, em-dashes, arrows, blank lines
type: feedback
scope: common
updated-at: 2026-05-30
---

Hook-enforced (the PreToolUse hook `check-doc-style.mjs` DENIES the write) for `.md` files under `.agents/` and `.claude/`:

- Bold (double-asterisk syntax): forbidden. Use plain text labels.
- Em-dash (U+2014): forbidden. Use " - " (space-hyphen-space) instead.
- Unicode arrow (U+2192 and similar): forbidden. Use `->` (ASCII) instead.

Scope note: human-facing docs under `docs/` (for example docs/plan/plan.md) are NOT watched by the hook, so they may use rich markdown (bold, tables, em-dash). The plain-text rule applies only to agent-config docs under `.agents/` and `.claude/`.

Convention only (not hook-enforced):

- Markdown tables ARE allowed. The folder-map tables in CLAUDE.md files rely on them.
- Prefer a heading over a horizontal rule `---` between sections. Frontmatter `---` at file top is required.
- At most one blank line between paragraphs.
- ASCII only. Avoid en-dash (U+2013), curly quotes, and typographic ellipsis (use three dots).

Why: em-dash and Unicode arrows are multi-byte UTF-8 (more tokens than ASCII). Bold adds visual clutter without information. Code blocks and inline code spans are exempt - they pass through verbatim.

How to apply: use plain text labels and `key: value` lines. When editing a file that has existing violations, convert the touched sections.
