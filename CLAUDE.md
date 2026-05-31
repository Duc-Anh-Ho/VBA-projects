# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Agent Behavioral Rules

Canonical source: `.agents/shared-rules.md`. Always active:

- **Language:** reply to the user in Vietnamese with full diacritics. Write all AI-managed artifacts (memory, configs, hook code, agent docs) in English. Human-facing docs under `docs/` may be Vietnamese.
- **Clarify before acting** when requirements are ambiguous, the action is destructive/hard to reverse (delete, reset, force-push, overwriting a binary workbook), has external side effects (push/publish/send), the scope is larger than stated, or a secret would be exposed.
- **Tool failure:** do not retry a failing call more than once, do not use guessed values, do not write anything depending on an unverified result. Stop and report.
- **No assumptions:** before writing an unverified config field / API param / CLI flag, ask the user for the official URL and fetch once. Uncertainty is enough to stop.
- **Storage:** all rules, memory, and config live inside the project and are git-tracked. Memory goes in `.agents/memory/<scope>/`, never the user home dir. Enforced by hooks.

### Session bootstrap

1. The memory index (`.agents/memory/MEMORY.md`) is auto-injected each prompt. Read the topic files it names that are relevant to the task.
2. For project direction/decisions, read `.agents/memory/project/vba-project-direction.md` and `docs/plan/plan.md`.

## Agent configuration (this repo)

Two-layer agent tooling, seeded 2026-05-30 from the QA-agent template:

- `.agents/` — cross-agent shared layer (memory, rules, hooks, tools manifest). Map: `.agents/CLAUDE.md`.
- `.claude/` — Claude-native (settings.json hook wiring, skills). Map: `.claude/CLAUDE.md`.
- Hooks enforce: memory-stays-in-project, plain-text agent docs, tool full-path manifest, commit message format + no `--no-verify`, VBA conventions reminder. Inventory: `.agents/hooks/CLAUDE.md`.
- Skills: `/tools-init` (rebuild the machine tool manifest), `/memory-merge` (pull user-local memories into the project).
- VBA coding standard: `.agents/rules/code/vba-conventions.md`.

## What this is

"Danh-Tools" is an Excel VBA add-in (ribbon-based productivity toolkit for sheets, charts, pivots, ranges, pictures, files, and keyboard shortcuts). The shipped artifact is `Danh-Tools-Installation.xlsb`: when a user opens it, `Workbook_Open` self-installs the workbook as `Danh-Tools.xlam` into Excel's `Application.UserLibraryPath` and enables it as an add-in. There is no compiler, no test runner, and no package manager — the "build" is opening the workbook in Excel.

## Source-of-truth model (critical)

The runnable code lives **inside** the binary `.xlsb`/`.xlam`/`.xlsm` files, which are not diffable or editable as text. Git tracks **exported text** of the VBA components under `VBA-files-<workbookBaseName>/` folders, produced by the in-house "Auto Backup" tool (`part-tools/auto-backup`, class `AutoVBE.cls`). Each export folder is laid out by component type:

- `Modules/*.bas` — standard modules
- `Classes/*.cls` — class modules
- `Forms/*.frm` (+ `.frx` binary resource) — UserForms
- `Else/*.cls` — document modules (`ThisWorkbook`, `SheetN`) exported with a `.cls` extension

Editing workflow: changes are made in the Excel VBE, then **exported** to these text files (Auto Backup tool, or the add-in's own *Export all VBA files* ribbon button) before committing. To apply text-file changes back into a workbook, use *Import all* — `AutoVBE.importAllVBAfiles` removes the matching component and re-imports from the folder; for document modules it replaces code in place and strips the 4 `VERSION/BEGIN/MultiUse/END` header lines. All import/export requires Excel's **"Trust access to the VBA project object model"** setting enabled.

When asked to change behavior, edit the `.bas`/`.cls`/`.frm` text files; never attempt to edit the binary workbooks directly.

## Repository layout

- `final-installation/` — **the main, current product.** `Danh-Tools-Installation.xlsb` + its `VBA-files-Danh-Tools-Installation/` export + `CustomUI14.xml` (the ribbon definition).
- `part-tools/` — standalone single-feature workbooks (auto-add-in, auto-backup, combine-sheets, invert-color, hide/show chart labels, …), each with its own `VBA-files-*` export. `auto-backup` is the tool used to produce all the exports.
- `draft-and-reference/` — old versions, password-broken copies, and backups. Reference only; do not treat as current.
- `malware-investigation/` — unrelated quarantined sample files. Do not open or execute.
- `xlwings/`, `app/`, `vbe-theme-regedit/`, `Images/` — experiments, ribbon-editor assets, VBE color theme, and ribbon icon source images.

## Architecture (main add-in)

The add-in is **ribbon-driven and stateless across callbacks**, which shapes everything:

- **`CustomUI14.xml`** declares the ribbon tab/groups/controls and wires every control to a callback (`onAction`, `getLabel`, `getImage`, `getSize`, `getEnabled`, `getVisible`, …). All callbacks resolve to public subs in **`Modules/CustomUi.bas`**.
- **`CustomUi.bas`** is the controller hub. Each callback `Select Case`s on `control.id` and reads/writes a per-control **`CustomUITag`** object (wraps that control's label/image/size/enabled state). `onAction` handlers instantiate a domain **Controller** class, call it, then release it.
- **Ribbon handle recovery**: Office gives the `IRibbonUI` only once, in `CustomUIOnLoad`. The pointer (`ObjPtr`) is stashed in a **Named range** inside the add-in (`InfoConstants.getRibbonID`). `Preprocessor.bas` declares the Win32 `CopyMemory`/`Sleep` APIs (guarded by `#If VBA7/Win64`) and `GetRibbon` rehydrates the `IRibbonUI` from that pointer so callbacks can call `ribbon.Invalidate` to refresh button state on demand.
- **Controller classes** (`Classes/*.cls`: `SheetsController`, `ChartsController`, `PivotTablesController`, `RangesController`, `FilesController`, `PicturesController`, `FormatController`, `ModeController`, `InternetConnector`, …) each own one feature area and are created/destroyed per action.
- **`SystemUpdate.cls`** is the shared utility used by nearly every class: `speedOn`/`speedOff` (toggle ScreenUpdating/Calculation/Events around an operation), `tackleErrors` (central error dispatcher), file-system and clipboard helpers, and `hasWorkPlace` (active-workbook guard).
- **`AutoAddin.cls`** implements self install/update/remove of the `.xlam`. **`InfoConstants.cls`** centralizes version, author, add-in name (`Danh-Tools`), and ribbon ID constants — change version/metadata here.
- **Keyboard shortcuts**: `ShortcutController` registers `Application.OnKey` bindings in `Workbook_Open` and unregisters them in `Workbook_BeforeClose`. The bound procedures live in **`Modules/Shortcuts.bas`** because `OnKey` targets must be in a standard module (global scope), not a class.
- **Live events** (selection change, etc.) are routed through `customEvents.cls` and the `MouseOverControl`/`Custom*Event` classes; `ThisWorkbook` also drives the auto-arrange picture loop via `Application.OnTime` self-rescheduling.

## VBA conventions used throughout

- `Option Explicit` in every component.
- Classes follow a fixed shape: `Class_Initialize` constructor, `Class_Terminate` destructor that `Set`s members to `Nothing`, a `hasVariables()` initializer, then methods grouped under `'ASSESSORS' / 'METHODS' / 'MAIN'` comment banners.
- Standard error pattern in public entry points: `On Error GoTo ErrorHandle` … `GoTo ExecuteProcedure` / `ErrorHandle: Call system.tackleErrors` / `ExecuteProcedure: Call system.speedOff`. Wrap heavy operations in `speedOn`/`speedOff`.
- Explicit `Let`/`Set` on assignments; controllers instantiated with `New`, used, then released.
- Match the surrounding style when adding code — comments are sparse and often reference "Check README.md for more information".

## Developer-only notes

- `Modules/Developer.bas` contains dev-machine specifics: `GIT_LOCAL_PATH = "S:\VBA-projects\"` and `aSaveBackup`, which saves the running `.xlam` back over the install `.xlsb` (`FileFormat:=xlExcel12`) and reopens it. These paths are author-specific.
- README warns: do **not** enable the auto-send-email feature (`Developer.autoSendWifi` / `EmailCDO` / `InternetConnector`) — antivirus flags it; it is intentionally left disabled in `Workbook_Open`.
