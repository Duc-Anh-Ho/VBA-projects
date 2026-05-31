#!/usr/bin/env node
/**
 * @file ./.agents/hooks/edit/inject-vba-reminder.mjs
 * @description PreToolUse Edit/Write hook. When the target is a VBA source file (.bas/.cls/.frm), injects a non-blocking reminder of the project VBA coding conventions as additionalContext, so each edit follows .agents/rules/code/vba-conventions.md without a full file read. Advisory only - never blocks. Defensive: any internal error fails-open (allow).
 * @scope project
 * @updated-at 2026-05-30
 */
import { readFileSync } from "fs";

const VBA_EXT = /\.(bas|cls|frm)$/i;

const REMINDER =
    `VBA CONVENTIONS (.agents/rules/code/vba-conventions.md): `
  + `Option Explicit at top; prefer string-specific functions (Left$ not Left, Mid$, Chr$); `
  + `fully-qualified calls (VBA.Strings.Left$); default Private, use Friend for in-project API, Public only for external; `
  + `PascalCase for modules/classes, camelCase for vars; class shape = Class_Initialize / Class_Terminate (Set members Nothing); `
  + `error pattern = On Error GoTo ErrorHandle ... GoTo ExecuteProcedure / ErrorHandle: tackleErrors / ExecuteProcedure; `
  + `no IIf in core/Utils/Controller modules. Edit text files; the runnable code lives in the binary workbook (see root CLAUDE.md).`;

const allow = (ctx) => {
  if (!ctx) { process.exit(0); }
  process.stdout.write(JSON.stringify({ hookSpecificOutput: { hookEventName: "PreToolUse", additionalContext: ctx } }) + "\n");
  process.exit(0);
};

try {
  const input    = JSON.parse(readFileSync(0, "utf8"));
  if (input.tool_name !== "Edit" && input.tool_name !== "Write") allow();
  const filePath = (input.tool_input?.file_path ?? "").replace(/\\/g, "/");
  if (!VBA_EXT.test(filePath)) allow();
  allow(REMINDER);
} catch {
  allow();
}
