#!/usr/bin/env node
/**
 * @file ./.agents/hooks/memory/inject-index.mjs
 * @description UserPromptSubmit hook. Reads .agents/memory/MEMORY.md and injects its content into the prompt context as additionalContext, so the memory index is always visible regardless of whether the Session Bootstrap Rule was followed. No-ops when the file is missing or the input is not a UserPromptSubmit event.
 * @scope project
 * @updated-at 2026-05-30
 */
import {
    readFileSync
  , existsSync
} from "fs";
import { join } from "path";

const PROJECT_ROOT = process.env.CLAUDE_PROJECT_DIR || process.cwd();
const MEMORY_INDEX = join(PROJECT_ROOT, ".agents", "memory", "MEMORY.md");

let input;
try { input = JSON.parse(readFileSync(0, "utf8")); } catch { process.exit(0); }
if (typeof input.prompt !== "string") process.exit(0);

let indexBody = "";
try { indexBody = existsSync(MEMORY_INDEX) ? readFileSync(MEMORY_INDEX, "utf8") : ""; } catch {}
if (!indexBody) process.exit(0);

process.stdout.write(JSON.stringify({
  hookSpecificOutput: {
      hookEventName     : "UserPromptSubmit"
    , additionalContext : `PROJECT MEMORY INDEX (.agents/memory/MEMORY.md):\n---\n${indexBody.trim()}\n---\nTo read a memory file: use path relative to project root.`
  }
}));
process.exit(0);
