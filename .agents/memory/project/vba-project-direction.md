---
file: ./.agents/memory/project/vba-project-direction.md
name: VBA project direction and decisions
description: vba-project-direction - strategic direction for Danh-Tools (portfolio + free/open-source + donate, not a heavy commercial rewrite), tech decision (Excel-DNA over VSTO if porting), and repo cleanup decisions
type: project
scope: project
updated-at: 2026-05-30
---

Context: this repo stores the author's VBA Excel tools. The most advanced version is
xlwings/DANH-project (47 VBA files, own git, own docs), newer than final-installation
(35 files). The author asked about converting to VSTO to sell commercially.

Direction decided (2026-05-30):
- Primary goal is PORTFOLIO + FREE / OPEN-SOURCE + DONATE, not a heavy commercial
  rewrite. Reason: the generic Excel-utility market is saturated (Kutools, ASAP
  Utilities); selling a cobbled-together personal toolset cold is low-odds. A clean,
  open-source portfolio plus a donate button is cheaper, lower-risk, and matches the
  project's original stated purpose (resume portfolio). Real income path is reputation
  -> custom-tool work / consulting / courses, not license sales.
- Validate willingness-to-pay with the existing VBA version BEFORE any rewrite.

Tech decision if/when porting to .NET:
- Prefer Excel-DNA over VSTO. Both are full .NET on Windows desktop with equal Excel
  COM power; Excel-DNA deploys as a single .xll (no ClickOnce pain), supports modern
  .NET, and supports UDFs. VSTO is in maintenance mode and stuck on .NET Framework.
- Office.js (cross-platform) cannot do this feature set (ShareX snip, PowerShell,
  global hotkeys, VBA import/export, Win32) - only consider for a future cut-down Lite.
- "Convert" = full rewrite VBA -> C#; no auto-converter. ~3-4 months full-time for the
  full suite if proficient in C#; a focused 1.0 is much less.

Repo cleanup decisions (planned, see docs/plan/plan.md):
- Single source of truth: treat xlwings/DANH-project as the main version.
- xlwings/DANH-project/.git is a nested independent repo (not a submodule, no remote) -
  must be resolved (submodule or subtree).
- Separate the personal PJ1_*/UT-log automation (sfcmaplg, EUC-JP, D:\share\ASS paths)
  from the product. Remove the auto-send-wifi/email feature (antivirus flags it).
- Binary .xlsb/.xlam belong in Git LFS or are built from the text export.

Full plan: docs/plan/plan.md.
