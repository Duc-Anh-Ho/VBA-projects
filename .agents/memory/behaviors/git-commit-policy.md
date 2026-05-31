---
name: git-commit-policy
description: "git-commit-policy - commit message format, scope separation, and no-bypass rules enforced by the commit hooks"
metadata: 
  node_type: memory
  file: ./.agents/memory/behaviors/git-commit-policy.md
  type: feedback
  scope: agent
  updated-at: 2026-05-30
  originSessionId: 51facd40-9c05-4990-a7d5-6db2064bd159
---

Commit only when the user asks. If on the default branch (main), branch first.

Message format (enforced by hooks/commit/block-message-format.mjs):

Two patterns, determined by what is staged:
- Project commit (VBA files, workbooks, docs): "Add: description"
- Agent commit (.agents/, .claude/, CLAUDE.md):  "Add (agent): description"

Type words: Add, Update, Fix, Refactor, Remove, Docs, Chore, Style, Test, Perf, Build, CI.
Line endings LF only. On Windows use HEREDOC (git commit -m "$(cat <<'EOF' ... EOF)").
If a body is present, separate from subject with one blank line.

Hook enforcement:
- Mixed scope (agent + project staged together) -> BLOCKED.
- Agent-only staged without "(agent)" in subject -> BLOCKED.
- Project-only staged with "(agent)" in subject -> BLOCKED.
- Wrong type word -> BLOCKED.

Scope separation (enforced, required by /commit-clean):
- .agents/, .claude/, CLAUDE.md = agent scope -> commit separately with "(agent)".
- VBA files, workbooks, docs = project scope -> commit separately without "(agent)".
- /commit-clean strips agent scope from the deploy branch, so mixed commits corrupt it.

No bypass (enforced by hooks/commit/block-no-verify.mjs):
- Never use --no-verify. If a hook is wrong, fix the hook under .agents/hooks/, do not
  bypass it (root-fix-only).

Never-commit: settings.local.json and tools.local.json are machine-local (gitignored).
Do not stage secrets or machine-specific paths.
