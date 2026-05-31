---
file: ./.agents/rules/code/vba-conventions.md
name: VBA coding conventions
description: vba-conventions - coding standard for Danh-Tools VBA source (Option Explicit, string-specific functions, fully-qualified calls, access modifiers, class shape, error pattern, naming)
type: project
scope: project
updated-at: 2026-05-30
---

Canonical source: the longer prose version lives at
xlwings/DANH-project/docs/CODING_CONVENTIONS.md. This file is the short rule the
agent applies on every VBA edit (also injected by hooks/edit/inject-vba-reminder.mjs).

Language best practices:

- Option Explicit at the top of every module and class.
- Prefer string-specific functions with the $ suffix for performance and type
  safety: Left$, Mid$, Right$, Chr$, Trim$, Format$, Str$ instead of the Variant
  forms. Exception: use the Variant form only when the value may be Null.
- Use fully-qualified calls for built-in functions: VBA.Strings.Left$,
  VBA.FileSystem.Dir, VBA.Interaction.MsgBox. Avoids ambiguity and self-documents.
- If/Else over IIf in core, Utils_, and Controller modules (IIf evaluates both
  branches - unsafe). IIf is tolerated only in project-specific implementation modules.

Access modifiers (principle of least privilege):

- Default to Private. Use Friend for the in-project API (callable across modules in
  this project, hidden from external workbooks). Use Public only for the official
  external-facing API.

Naming:

- PascalCase for file/module/class/interface names. Prefixes: Utils_ for utility
  classes, Interface_ for interfaces.
- camelCase for variables and parameters.

Class shape:

- Class_Initialize constructor and Class_Terminate destructor that Sets every member
  to Nothing. A hasVariables() initializer where used. Group methods under
  'ASSESSORS / METHODS / MAIN' comment banners.

Error and performance pattern (public entry points):

- On Error GoTo ErrorHandle ... GoTo ExecuteProcedure / ErrorHandle: Call
  system.tackleErrors / ExecuteProcedure: Call system.speedOff. Wrap heavy work in
  speedOn / speedOff.

Source-of-truth reminder:

- The runnable code lives INSIDE the binary workbook (.xlsb/.xlam). Git tracks the
  exported text under VBA-files-*/ (Modules/.bas, Classes/.cls, Forms/.frm+.frx,
  Else/.cls). Edit the text files; never edit the binary. Round-trip via the
  Auto-Backup tool or the add-in's Import/Export ribbon buttons. See root CLAUDE.md.
