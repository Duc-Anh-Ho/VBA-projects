# Danh-Tools - Current Architecture

> Current-state architecture of the Danh-Tools VBA add-in (xlwings/DANH-project).
> Complements the future-state `ARCHITECTURE_REFACTORING_PLAN.md`. Generated from a
> full read of all 47 VBA source files (~15,300 lines), 2026-05-30.
> Companion: `MODULE-REFERENCE.md` (per-file API).

## 1. Big picture

Danh-Tools is a ribbon-driven Excel add-in (`Danh-Tools.xlam`) shipped as a self-installing
workbook (`Danh-Tools-Installation.xlsb`). The architecture has four layers plus one
self-contained personal subsystem that is **not** part of the product:

| Layer | Members | Role |
|---|---|---|
| Presentation | `CustomUI` XML + `CustomUi.bas` + Forms (`KeyboardShortcutForm`, `MultipleReplaceForm`, `SnippingToolForm`, `CommandRunnerForm`) | Ribbon callbacks + dialogs |
| Application (controllers) | `*Controller.cls` (Sheets, Charts, Pivot, Ranges, Format, Pictures, Files, Mode, Internet, Email, PowerShell, ShareX, Shortcut) | One feature area each; orchestrate, do not hold cross-cutting logic |
| Core / infrastructure | `SystemUpdate`, `InfoConstants`, `AutoAddin`, `Preprocessor.bas`, `MouseScroll.bas`, `Shortcuts.bas`, event classes | Shared services, Win32, lifecycle, global input |
| Utils / DI | `Utils_*` (Scripts, Test, Address, Stringify, Error, Clipboard, VBE) + `Interface_Address` | Newer reusable utility surface + dependency-injection plumbing |
| Personal subsystem (NOT product) | `PJ1_Logic`, `PJ1_Address` (+ `Utils_Test`) | UT-log verification tool tied to one external test workflow |

This is the more evolved version (47 files) vs `final-installation` (35 files): it adds the
`Utils_`/`Interface_` DI layer, `ShareXController`, `CommandRunnerForm`, and the PJ1 subsystem.

## 2. Ribbon system (how a click flows)

1. **CustomUI XML** declares the `danh-tools` tab/groups/controls and binds each control to
   named callbacks (`onLoad`, `getLabel`, `getImage`, `getEnabled`, `getVisible`, `getSize`,
   `getShowImage/Label`, `getKeytip`, `getScreentip`, `getSupertip`, `getPressed`, dropdown
   `getItemCount/ID/Label/getSelectedItemIndex`, and `onAction`). All resolve to public procs
   in `CustomUi.bas`. Control `id`s match the IDs assigned in `configTags`.
2. **onLoad captures the ribbon.** `CustomUIOnLoad` stores the live `IRibbonUI` in the public
   `loadedRibbon` and persists `ObjPtr(ribbon)` into a **Named range** in the add-in workbook
   (key = `InfoConstants.getRibbonID`). It then builds state via `setDefaultSettings` +
   `configTags`.
3. **State objects.** `configTags` creates ~70 `CustomUITag` objects (one per control), each a
   plain holder of label/image/size/enabled/visible/tips computed from module-level **mode
   flags** (`hasWorksheet`, `hasHighlight`, `isAutoArrange`, `hasListSheet`, `hasSYNCPivot`,
   `isArranging`, ...). `setButtonAsMode` rewrites toggle-button label/image to reflect mode.
   These tags are the single source of truth the getters read.
4. **Getters serve attributes.** Each `get*` callback does `If loadedRibbon Is Nothing Then
   refreshCustomRibbon`, then `Select Case control.id` to return the matching `CustomUITag`
   property.
5. **onAction -> domain controllers.** Click handlers (`sheetController`, `chartController`,
   `rangeController`, `pictureController`, `pivotControllerEvent`, `VBAFilesController`,
   `internetController`, `addinController`, `hidePageBreakChange`, `offsetSelect`, ...)
   instantiate the relevant controller, perform the action, update mode flags, and call
   `refreshCustomRibbon(loadedRibbon)`.
6. **Pointer recovery + invalidate.** When `loadedRibbon Is Nothing` (add-in re-entry, project
   reset, error 91), `refreshCustomRibbon` calls `Preprocessor.GetRibbon(...Names(getRibbonID))`,
   which reads the stored pointer string and `CopyMemory`s it back into an `IRibbonUI`. It then
   rebuilds tags and calls `Invalidate` so Excel re-queries every getter.

Trade-off: every callback is a hand-maintained `Select Case` over control IDs, so adding one
control means editing ~10 callbacks. This is the single biggest source of ribbon bugs (see 7).

## 3. Event + input system

- **Workbook/sheet events (`customEvents.cls`).** A `CustomEvents` instance (created in
  `setDefaultSettings`) sinks `Application` events. Sheet/workbook activate/deactivate/close
  recompute `CustomUi.hasWorksheet/hasWorkChart/hasWorkDialog` and call `refreshCustomRibbon`,
  so ribbon enabled/visible state always matches context. While a stateful mode is on (List
  Sheets, SYNC Pivot, Highlight, Arrange, Auto Arrange), switching workbooks is blocked by
  `popup` (re-activates `previousWb` under `EnableEvents=False`).
- **Event-driven controllers.** `SheetsController`, `PivotTablesController`, and
  `RangesController` hold their own `WithEvents` sinks and react to
  SheetActivate/Deactivate/Change/SelectionChange (auto-maintain list columns, refresh pivot
  caches, draw the crosshair highlight, store/paste highlight format on selection).
- **Form control events (`MouseOverControl` + `CustomLabelEvent` + `CustomTextBoxEvent`).**
  Because `KeyboardShortcutForm` builds its rows at runtime, each created label/textbox is
  wrapped in an event-bridge object held in the form's `eventColl`; their `WithEvents`
  handlers forward MouseMove/Click/DblClick/KeyDown/Change to the form's public methods.
- **Global mouse wheel (`MouseScroll.bas`).** `EnableMouseScroll` installs a thread-local
  `WH_MOUSE` hook (cristianbuse/VBA-UserForm-MouseScroll, MIT). `MouseProc` defers work via a
  self-terminating `MouseOverControl`, and `ProcessMouseData` routes wheel deltas to
  Scroll/Zoom against the hovered control. The hook auto-removes when all forms are destroyed.
- **Keyboard shortcuts (`Shortcuts.bas` + OnKey).** `ShortcutController.install` reads a
  worksheet keybinding table and registers `Application.OnKey` -> `Shortcuts.<procedure>`. OnKey
  targets must live in a standard module, hence `Shortcuts.bas`. `KeyboardShortcutForm` is the
  editor that re-installs bindings.

## 4. Add-in lifecycle

- **Open (`ThisWorkbook.Workbook_Open`):** instantiate `AutoAddin` -> `install()`, then
  `ShortcutController` -> `install()`. (A commented-out `Developer.autoSendWifi` call is wired
  but disabled.)
- **Install/update (`AutoAddin`):** resolves `Application.UserLibraryPath` + `Danh-Tools.xlam`;
  if not already the add-in, prompts to update when it exists, disables+`Kill`s the old file,
  `SaveAs xlOpenXMLAddIn`, enables via `AddIns(...).Installed = True`, then closes the installer.
- **Remove (`AutoAddin.remove`):** optional confirm, switch current wb to read-only if it is the
  add-in, `Kill` the `.xlam`, disable.
- **Close (`Workbook_BeforeClose`):** `ShortcutController.unInstall` removes the OnKey bindings.
- **OnTime loop:** `Auto_Run_Continuously` self-reschedules every ~0.625s to arrange shapes;
  `Stop_Run_Continously` cancels it.
- **Sheets:** Sheet3 = the "ping" network-status sheet (toggle button runs a ping loop). Sheet1,
  2, 4, 5, 6 have empty code-behind (config/data sheets; roles not determinable from code).

## 5. Utils / DI layer

`Utils_Scripts` is a lazy COM-object factory/cache (FileSystemObject, WScript.Shell,
Shell.Application, ADODB Stream/Connection/Recordset, Dictionary, RegExp, CDO, WinHTTP, WinSCP,
MSForms clipboard) behind one `safeCreate` wrapper. `Utils_Test` is the file-system / IO / test
workhorse (path checks, recursive create/copy/delete, binary diff, encoding detection via nkf,
tail-reading, ADO-OLEDB reads of closed Excel workbooks, regex). `Utils_Address` maps logical
keys to sheet-qualified ranges and reads/writes scalars/arrays. `Utils_Stringify` does sheet-name
quoting + `{{placeholder}}` templating. `Utils_Error` centralizes an `ECODE` enum + file logging.
`Utils_Clipboard` does dual COM/Win32 Unicode clipboard. `Utils_VBE` automates the VBE IDE.

### Dependency-injection pattern (the template to reuse)

```
Interface_Address  (contract: Public Sub Pair(Addr As Utils_Address))
        ^ Implements
PJ1_Address        (concrete: Pair -> ~45 Addr.Add("KEY","cell"))
        | injected into
Utils_Address.Init(sheet, addressSource As Interface_Address)
        -> addressSource.Pair(Me)   ' inversion of control
```

Composition root is `PJ1_Logic.Init`, which does `Addr.Init(targetWorksheet, New PJ1_Address)`.
`Utils_Address` knows nothing about PJ1, only the interface. To reuse: write
`PJ2_Address Implements Interface_Address`, fill `Pair`, inject it - no change to `Utils_Address`.
This is the pattern the refactoring plan wants to spread across the whole codebase.

## 6. PJ1 subsystem (personal, NOT product)

`PJ1_Logic` (940 lines) + `PJ1_Address` (66) are a personal work-automation tool that verifies
unit-test (UT) runs against an external test environment. For a `TASKS` list it checks
checklist/UT/log file existence, counts folder entries, detects encoding (nkf), and parses
Japanese (EUC-JP) `sfcmaplg`/`sfcmerlg`/`.log` output plus UT Excel sheets (read via ADO OLEDB
without opening them) for run times, PIDs, errors, and patterns, writing PASS/FAIL back to the
sheet.

It is hard-wired to one workflow: hardcoded paths (`D:\share\ASS\Tasks\template\...`), the
`sfcmaplg`/`sfcmerlg`/`testcsh`/`ALS` log conventions, EUC-JP encoding, the external `nkf`/
`autotest.exe`/`node`/`git` tools, and a fixed 229-row sheet layout. **Per the project plan,
this must be split out of the product.** Only the `Utils_`/DI plumbing under it is reusable.

## 7. Cross-cutting patterns

- **Controller shape:** most controllers declare `info As InfoConstants` + `system As
  SystemUpdate`, defer real init to a `hasVariables()` that news them up and gates on
  `system.hasWorkPlace(hasMsg:=True, workPlaceType:="xlWorksheet")`. `Class_Initialize` is
  usually empty; `Class_Terminate` sets references to Nothing. Exceptions: `ShareXController`
  and `ShortcutController` follow neither pattern.
- **Error pattern:** `On Error GoTo ErrorHandle` / `If Not hasVariables Then GoTo
  ExecuteProcedure` / body / `GoTo ExecuteProcedure` / `ErrorHandle: Call system.tackleErrors` /
  `ExecuteProcedure:`. `SystemUpdate.tackleErrors` is the single error sink.
- **Performance:** public entry points wrap work in `system.speedOn` / `speedOff`
  (screen/events/alerts off). `FormatController` deliberately omits it for copy/paste;
  `ChartsController` omits it.
- **External/OS surface concentrates risk:** `PowerShellController`, `InternetConnector`,
  `EmailCDO`, `ShareXController` reach outside Excel. See section 8.

## 8. Known issues, bugs, and risks (found during the read)

Security / privacy (must address before any release or open-source):
- `EmailCDO.cls`: hardcoded Gmail username + app password in source; sends host name, user, and
  full `ipconfig /all` to hardcoded recipients (data-exfiltration shape).
- `InternetConnector.cls`: exports saved Wi-Fi passwords in clear text via `netsh ... key=clear`.
- `Developer.autoSendWifi`: flagged "DON'T RELEASE"; wired (commented) into `Workbook_Open`.
- `Preprocessor.GetRibbon` / `MouseScroll`: raw `CopyMemory` from a stored pointer can crash the
  process if the pointer is stale.

Portability:
- Hardcoded machine paths: `ShareXController` (`D:\Program Files\ShareX\ShareX.exe`),
  `Developer.bas` (`S:\VBA-projects\`), `Utils_Test`/`PJ1_Logic`/`Shortcuts.bas`
  (`D:\share\ASS\...`, `D:\environments\node\node.exe`), `Utils_Error` (`C:\DT_Logs`).

Correctness bugs flagged by the read (verify before fixing; see self-review note in chat):
- `CustomUi.getShowLabel`: `hidePageBreakDropDown` case returns `multipleReplaceButton.getShowLabel`.
- `SystemUpdate.createDictionary`: assigns its result to `createClipboard` instead of itself.
- `SystemUpdate.getLastRow/getLastColumn`: test `cell.value = False` to detect empties (fragile).
- `ShortcutController`: Numpad6/7 names swapped in conversions; `PrintKey` mapped twice.
- `KeyboardShortcutForm.formatLabel`: first two `Case isPickingLine` branches identical
  (duplicate-keybinding branch unreachable).
- `Utils_Test.WINDOW`/`MAC`: typed `Boolean` but assigned String constants.
- `Utils_Scripts.HTTPS`: progID misspelled (`WinHttpRequets.5.1`) - will fail to create.
- `PJ1_Logic.LastUpdate`: copy-paste duplicate of `CheckLogContent`; several `Dim ... As String`
  vars assigned Booleans.

Maintainability:
- `SystemUpdate` is a ~970-line God class (app state + file system + clipboard + arrays + sheet
  detection + COM factory + error handling). The COM factory block is self-flagged `TODO: Remove`.
- The ribbon callbacks are ~10 parallel `Select Case` tables that must stay in sync by hand.
