---
type: canonical-rules
auto-loaded-by: root CLAUDE.md (inlined behavioral rules)
sync-targets: CLAUDE.md "Agent Behavioral Rules"
updated-at: 2026-05-30
---

# Shared Agent Rules

Edit this file first when updating cross-agent rules, then sync the changes into the root CLAUDE.md "Agent Behavioral Rules" section.

## 1. Language policy

- Respond to the user in the language they write in (Vietnamese here). Use full Vietnamese diacritics in multi-paragraph replies.
- Write all AI-managed artifacts in English (memory files, configs, hook code, agent docs).
- Human-facing docs under docs/ may be Vietnamese.

## 2. Clarification protocol

Stop and ask the user before acting when any of these apply:
- Requirements are ambiguous, incomplete, or contradictory - do not guess intent.
- The action is destructive or hard to reverse (delete, reset, force-push, overwrite a binary workbook).
- The action has external side effects (git push, publish, send, post).
- Something in the codebase contradicts the request.
- The requested scope is significantly larger than stated.
- A security-sensitive file or credential would be modified or exposed.

## 3. Tool and fetch failure protocol

When a tool call fails: do NOT retry the same call more than once, do NOT use workarounds or guessed values, do NOT write anything that depends on the unverified result. Stop and report what failed and what you did not touch.

## 4. No-assumption rule

Before writing any config field, schema, API parameter, or CLI flag you do not have verified knowledge of: ask the user for the official URL and fetch it once. Uncertainty about a field name is enough to stop.

## 5. Storage policy

All rules, memories, configs, and reference docs must live inside the project directory and be git-tracked. Never write to the user home dir (~/.claude, %USERPROFILE%\.claude) or system paths. Git is the sync mechanism across machines. Memory files go to `.agents/memory/<scope>/`. The check-write-rules hook enforces this.

## 6. Shell environment

This machine runs Windows with PowerShell as the default shell; Bash (Git for Windows) is also available. Use `.agents/tools.local.json` for external CLI paths. Do not silently substitute a missing tool. Avoid `/dev/stdin` (not portable on Windows shells).

## 7. Docs and memory file style

Agent-config markdown under .agents/ and .claude/ is plain text: no bold, no em-dash, no Unicode arrows, ASCII only. Tables are allowed. See rules/doc/doc-style.md. Code source files follow their own language conventions (VBA: rules/code/vba-conventions.md).
