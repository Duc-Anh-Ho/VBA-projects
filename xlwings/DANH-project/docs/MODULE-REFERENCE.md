# Danh-Tools - Module Reference

> Per-file API reference for all 47 VBA source files, grouped by layer. Generated from a
> full read of the source, 2026-05-30. See `ARCHITECTURE.md` for the big picture.
> "API" lists Public/Friend members; Private members appear only when architecturally notable.

## Layer 1 - Presentation (ribbon + forms + input)

### CustomUi.bas (2522 lines)
Backs the entire custom ribbon: builds control-state objects, serves every ribbon callback,
dispatches clicks to controllers.
- State: `configTags` builds ~70 `CustomUITag` objects from mode flags; `setButtonAsMode`
  rewrites toggle button label/image; `setDefaultSettings` resets flags + controller refs.
- Lifecycle: `CustomUIOnLoad` (onLoad) stores `IRibbonUI` in `loadedRibbon` + `ObjPtr` in a
  Named range; `refreshCustomRibbon([rb])` recovers the pointer, rebuilds tags, calls `Invalidate`.
- Getter callbacks (each `Select Case control.id`): `getSize`, `getImage`, `getEnabled`,
  `getShowImage`, `getKeytip`, `getLabel`, `getShowLabel`, `getScreentip`, `getSupertip`,
  `getVisible`, `getPressed`; dropdown `getItemCount/getItemID/getItemLabel/getText/getSelectedItemIndex`.
- onAction dispatchers: `sheetController`, `sheetControllerEvent`, `chartController`,
  `pivotControllerEvent`, `VBAFilesController`, `rangeController`, `rangeControllerEvent`,
  `hidePageBreakChange`, `accessSettings` (stub), `internetController`, `addinController`,
  `pictureController`, `offsetSelect`.
- Public fields read elsewhere: `loadedRibbon`, `arrangeButton`, the boolean mode flags.
- Risk: ~10 parallel `Select Case` tables to keep in sync; known bug in `getShowLabel`
  (hidePageBreakDropDown returns multipleReplaceButton).

### CustomUITag.cls
Plain value holder for one ribbon control. Get/Let pairs for `id, size, Description, isEnabled,
image, isShowImage, keytip, label, isShowLabel, screentip, supertip, isVisible`. No logic.

### KeyboardShortcutForm.frm (1217 lines)
VS-Code-style keyboard-shortcut editor: renders one row per shortcut, captures keypresses,
flags duplicates, persists via `ShortcutController`.
- Lifecycle: `UserForm_Initialize` builds rows (`initRow`/`createRow` add Label/TextBox/Line
  controls at runtime), enables `MouseScroll`, wires `storeCustomEvent`. Terminate/QueryClose
  disable scroll + `clearUp`; `closeForm` does `Unload` + `End`.
- Interaction: `labelMoveOn/labelClick/labelDbClick` (from `CustomLabelEvent`),
  `textBoxKeyDown` (from `CustomTextBoxEvent`) convert keys via `ShortcutController.convertKeyToName`
  (Enter=save, Esc=cancel, Backspace=clear); `KeyboardFrame_KeyDown` adds nav.
- State: `editedArr()`, `applyArr()`; `isEditedLine/isDuplicatedLine/hasDuplicated`; `formatLabel`
  styles rows.
- Buttons: `ApplyButton_Click` (blocks on duplicates -> convertNameToCode -> setColData + install
  -> Save), `ApplyAndCloseButton_Click`, `CancelButton_Click`, `EditButton_Click`,
  `RestoreDefaultButton_Click` (stub).
- Risk: first two `Case isPickingLine` branches in `formatLabel` are identical (dup branch dead);
  `closeForm`/Terminate call `End` (hard-stops all VBA state).

### MultipleReplaceForm.frm (267 lines)
Find/replace where Find and Replace are parallel ranges, scoped by Selection/Sheet/Workbook.
- `UserForm_Initialize` sets checkboxes/combos/tab order; `*Label_Click` move focus;
  `WithInComboBox_Change` shows `SelectedAreaInput` only for "Selection"; `*Input_Exit` validate
  via `Application.Evaluate`; `ReplaceAllButton_Click` validates (incl. equal Find/Replace counts)
  then calls `RangesController.multipleReplace(...)`. `closeForm` does `Unload` + `End`.

### SnippingToolForm.frm
Front-end to `PicturesController.snip` with a lock-ratio checkbox. `UserForm_Activate/Deactivate`
create/clear a `PicturesController`; `Scisors_Icon_Click` -> `pic.snip`. Secondary entry point
(most snipping is invoked directly from the ribbon).

### CommandRunnerForm.frm
Caption "Command Runner". The exported `.frm` has only `Option Explicit` and no procedures - all
layout is in the binary `.frx`. Logic-less in source (placeholder/WIP; not wired to the ribbon).

### customEvents.cls
Application-level event sink keeping the ribbon in sync and blocking workbook switching during a
stateful mode. `WithEvents appEvent/wbEvent/WsEvent`. Handlers: `appEvent_SheetActivate/Deactivate`
(recompute hasWorksheet/Chart/Dialog + refresh), `appEvent_WorkbookBeforeClose/Activate/Deactivate`;
private `popup` forces focus back to `previousWb` under `EnableEvents=False`.

### MouseOverControl.cls
Wraps any of 13 MSForms control types (or the form) to capture `MouseMove`, and doubles as an
async callback whose `Terminate` drives `ProcessMouseData`. Factories `CreateFromControl/CreateFromForm`;
`getControl`, `FormHandle`, `IsAsyncCallback`.

### CustomLabelEvent.cls / CustomTextBoxEvent.cls
Event bridges wrapping a single MSForms.Label / .TextBox so dynamically-created form controls can
raise events. Label: forwards MouseMove/Click/DblClick to `KeyboardShortcutForm.labelMoveOn/labelClick/labelDbClick`.
TextBox: forwards KeyDown/Change to `textBoxKeyDown/textBoxChange`. Must be held in a collection to stay alive.

### MouseScroll.bas (1160 lines)
Self-contained MIT library (cristianbuse/VBA-UserForm-MouseScroll) installing a thread-local
`WH_MOUSE` hook so forms/controls scroll/zoom with the wheel.
- Public: `EnableMouseScroll(uForm, passScrollToParentAtMargins, useShiftForPerpendicularScroll,
  useCtrlToZoom)`, `DisableMouseScroll(uForm)`, `SetHoveredControl(moCtrl)`, `ProcessMouseData`.
- Plumbing: `SetWindowsHookEx`/`UnhookWindowsHookEx`/`CallNextHookEx`, `AddressOf MouseProc`
  (x64 ASM ret-offset patch via `CopyMemory`), window/input APIs (`WindowFromPoint`, `IsChild`,
  `GetKeyState`, `SystemParametersInfo`, `PostMessage`), scroll engine (`ScrollY/ScrollX/Zoom`).
- Local edits marked "DANH EDIT" (`CollectionHasKey`, `UpdateLastCombo` guards).
- Risk: pressing the VBE Reset button inside the hook scope is dangerous (warned in comments).

### Preprocessor.bas
Conditionally-compiled Win32 declarations + the ribbon-pointer helper. `Option Private Module`.
- kernel32: `Sleep`, `GetTickCount`, `CopyMemory`/`RtlMoveMemory`, `GlobalAlloc/Lock/Unlock/Size`,
  `lstrcpy`. user32: clipboard APIs (`OpenClipboard`/`SetClipboardData`/...).
- `GetRibbon(ribbonName As Name) As IRibbonUI`: parses the Named-range pointer string, `CopyMemory`s
  it back into an object (VBA7 + VBA6 overloads). MAC branch is unsupported.

### Shortcuts.bas (1130 lines)
Houses every macro target invoked by `Application.OnKey` (must be in a standard module).
- Thin wrappers: copyName/copyFullName/copyShortName/copyPath/copyExtensionName (FilesController);
  copyF/pasteF/pasteV/clearContent/clearFormat/clearAll, shapeMoveAndSize/shapeMove/shapeFree
  (FormatController); sheetSelectN/P, sheetFocusRename (SheetsController); toggleZenMode/toggleZoomMode*
  (ModeController); openMultipleReplaceForm, openShortcutForm; captureShareX* (PicturesController).
- Richer logic: shape grouping (Union-Find: isIntersect/findRoot/mergeRoot/groupIntersectedShapes/...),
  listAllShapes/renameAllShapes, convertGroupToImage, autoFill, plus an external-tooling suite that
  shells node/autotest/git and reads Excel via ADODB (uatTest, checkEncodeAndEOF, gitCheckout*, ...).
- Risk: hardcoded paths (`D:\environments\node\node.exe`); PJ1-style business logic mixed into the
  generic shortcut layer.

## Layer 2 - Application (controllers)

### SheetsController.cls (744)
Bulk worksheet ops + live events. API: `add()`, `deleteAll()`, `list([onSheet])`,
`rename([onSheet])`, `hide(isHide, [isVeryHide])`, `selectNext()`, `selectPrevious()`,
`focusRename([onSheet])`; props `hasListSheet`, `hasRenameSheet`. Uses magic A1:C1 markers
(`~No.~`, `~SHEET NAMES~`, `~RENAME~`) + heavy WithEvents.

### ChartsController.cls
`hide([isHide=True])` - blanks/restores error-valued chart data labels (sets `datalabel.text=""`
as a workaround). No speedOn/speedOff.

### PivotTablesController.cls
`refreshAllPivotTableCaches()` - refreshes every PivotCache; also auto-runs in the constructor and
on every SheetChange (perf cost on large books).

### RangesController.cls (379)
API: `invertColor()`, `boldFirstLine()`, `storeHighlightFormat([onSheet])`,
`pasteHighlightFormat([onSheet])`, `highlight(target)`, `displayPageBreak(isDisplay, [isApplyAll])`,
`multipleReplace(findArea, replaceArea, withinIndex, [selectedArea], [isMatchCase], [isMatchByte],
[isMatchContent], [searchOrderCd], [isOrderByLength]) As Boolean`; highlight props
(color/bold/blurRate/addSize). Uses temp sheets `formatStored` + `mult-replace-tmp`; crosshair
highlight runs on selection events.

### FormatController.cls (241)
API: `copyFormat()`, `pasteFormat()`, `pasteValue()`, `setPlacement([placementStt=xlMoveAndSize])`,
`clearContent()`, `clearFormat()`, `clearAll()` (+ low-level clear* variants). Dispatches via
TypeName to Shapes.PickUp/Apply; relies on `Application.CutCopyMode`. speedOn/Off intentionally
commented out.

### PicturesController.cls (657)
Screen-clip/snip pictures into a target cell/shape + shape arrangement. API: `snip()`,
`snipShareX(workFlow, [savedPath])`, `assign()`, `arrange(objectName)`, `clearArrange()`,
`autoArrange(isOn)`, `arrangeToMerge()`, `selectShapeInRange([mode])`, `getShapeInRange(target,
[mode]) As String()`, `lockRatio(targetObject)`, `sendTo(targetObject, [orderType])`; props
`letOffset`, `letLockRatio`, `selectObjectMode`, `hasFakeBorder`, `hasMarkTargetPlace/Object`;
Public Enum `TOUCH_MODE {OVERLAP, INSIDE}`. Uses ShareXController; `capturePicture` recurses on failure.

### FilesController.cls (334)
VBA component import/export + filename copy. API: `importSelectedVBAfiles()`,
`importAllVbaFiles()`, `exportAllVbaFiles()`, `copyFileName([typeName="name"])`. Needs "Trust access
to the VBA project object model"; security-sensitive (self-modifies the running VBProject).

### ModeController.cls
Zen mode + zoom. API: `toggleZenMode()`, `toggleZoomMode([isMax])`, `zoom100()`. Uses legacy
`ExecuteExcel4Macro` to hide the ribbon.

### InternetConnector.cls (190)
Ping + Wi-Fi export. API: `isConnect([link="google.com"]) As Boolean`, `saveWifiAsTxt/Json/Csv([useThisFolder])`;
Public Enum `exportType {TXT, CSV, JSON}`. SECURITY: extracts saved Wi-Fi passwords in clear text via
`netsh ... key=clear`.

### EmailCDO.cls (154)
`send([attachmentPath])` - Gmail SMTP via CDO, body = machine name + user + `ipconfig /all`.
SECURITY: hardcoded Gmail credentials + recipients; data-exfiltration shape.

### PowerShellController.cls
Wrapper around WScript.Shell. API: `createPWShellCommand As String`, `runScript(scrpit) As Byte`
(waits, returns exit code), `executeScript(scrpit) As String` (captures stdout/stderr; busy-waits
with 1s `Application.Wait`).

### ShareXController.cls
Launches external ShareX to run a named workflow, waits via WMI. API: `letWorkflow`, `getWorkflow`,
`START([timeoutMs=5000])`. Hardcoded path `D:\Program Files\ShareX\ShareX.exe`; does not use the
standard controller/error pattern.

### ShortcutController.cls (800)
Install/uninstall OnKey shortcuts from a worksheet table + key-name conversions. API: `install()`,
`unInstall()`, `convertKeyToName(KeyCode, Shift) As String`, `convertCodeToName(code) As String`,
`convertNameToCode(name) As String`, `getShortcutTable As ListObject`, `setColData(data(), colNo)`,
plus column-index / sentinel / key-name property getters; Public Enum `KEY_CODE`. Reads
`Sheets("keyboard-shortcut").ListObjects("Keybinding")`. Bugs: Numpad6/7 names swapped; RightControl/
RightWin mismapped; PrintKey mapped twice. No tackleErrors pattern.

## Layer 3 - Core / infrastructure

### SystemUpdate.cls (970) - God class
Grouped responsibilities (decomposition candidates):
- App state/perf: `speedOn`, `speedOff`, `setStatusBar` (Let), `isHideRibbon` (Let), `getFomulaSeparator`.
- Context detection: `hasWorkPlace([hasMsg],[workPlaceType])`, `hasApplication([appName])`,
  `hasWorkbook(wbName)`, `hasSheet(shName)`; cached fields `app/wb/ws/wd/sheetOb/chartSh/dialogSh`.
- File system: `getExcelPath([FileFilter])`, `getExcelFile(path)`, `getFileName([typeName])`, `getFolder()`.
- Position/array: `getLastRow(ws,[atColumn])`, `getLastColumn(ws,[atRow])`, `getArrayLength(arr,[dim])`,
  `getArrayDimension(arr)`, `mergeTwoArrays(a,b)`, `mergeMulArrays(ParamArray)`.
- Clipboard: `getClipboard()` (Get), `setClipboard` (Let).
- Sheet formatting: `storeSheetFormat(fromSheet,[toSheetName])`, `pasteSheetFormat(toSheet,[fromSheet])`.
- Timer: `getTimerMilestone()`, `restartTimer()`.
- COM factory (self-flagged TODO: Remove): `createFileSystem/createWShell/createFileStream/createCDOConfig/
  createCDOMess/createHTMLFile/createClipboard/createDictionary`.
- Error: `tackleErrors()` (+ private `errorDisplay`). Plus Public `WithEvents app/wb/ws`.
- Bugs: `createDictionary` assigns to `createClipboard`; `getLastRow/Column` use `cell.value=False`.

### InfoConstants.cls
Read-only metadata. API: `getVersion` ("v2.3.4"), `getUpdate`, `getPrompt`, `getAuthor` ("DANH"),
`getAddinShortName` ("Danh-Tools"), `getAddinExtension` (".xlam"), `getAddinName`, `getRibbonID`.
Comment warns short name must NOT include `_AddIn`.

### AutoAddin.cls (220)
Self-installer. API: `install()`, `remove([hasConfirm])`. SaveAs `xlOpenXMLAddIn` into
`Application.UserLibraryPath`; on update disables+`Kill`s old then re-enables; closes installer.
Mixed literal `AddIns("danh-tools")` vs parameterized name.

### Developer.bas (332)
Dev-only helpers (many for the Immediate window). API includes `showDPI()`, `aSaveBackup()` (saves
as `.xlsb`, OnTime reopen), `autoSendWifi()` (TEST-only Wi-Fi email - DO NOT SHIP), `aaTestCode()`,
`Clip/Clip2`, VBE window helpers, collection/range/file converters (`CollectionToArray`, `RangeToArray`,
`RangeToFile`, `FileToRange`, ...). Hardcoded `GIT_LOCAL_PATH = "S:\VBA-projects\"`.

### ThisWorkbook.cls
Workbook event host. API: `Workbook_Open` (AutoAddin.install + ShortcutController.install),
`Workbook_BeforeClose` (ShortcutController.unInstall), `Auto_Run_Continuously`/`Stop_Run_Continously`
(OnTime shape-arrange loop), private `Auto_Arrange_Shape`. Commented-out `autoSendWifi` block.

### Sheet1..Sheet6.cls
Sheet3 = "ping" network-status sheet (toggle button runs a ping loop over an address list: columns
No/Name/Address/Status, data from row 3; busy-loop without DoEvents). Sheet1/2/4/5/6 = empty
code-behind (`Option Explicit` only); roles not determinable from code.

## Layer 4 - Utils / DI

### Utils_Scripts.cls
Lazy COM-object factory/cache via one `safeCreate`. Get props: `FileSystem`, `WShell`, `AShell`,
`FileStream`, `Connection`, `RecordSet`, `CDOConfig`, `CDOMess`, `HTMLFile`, `Clipboard`,
`Dictionary`, `HTTPS` (progID typo `WinHttpRequets`), `WinSCP`, `RegexExpression`. `safeCreate`
shows a MsgBox on failure.

### Utils_Test.cls (887)
File-system / IO / test workhorse. API (selected): OS/path props (`OS`, `WINDOW`, `MAC`, `GIT`,
`SPLASH`, `Where(execName)`); existence/count (`IsExistPath/Folder/File`, `CountFilesInFolder`,
`CountInFolder`); path (`GetParentFolder`, `GetShortName`); fs ops (`CreateFolder`, `RenameFile`,
`CopyFile`, `CopyFolder`, `DeleteFolder`); zip (`ExtractExcelByZip` via PowerShell Expand-Archive);
diff/encoding (`DiffCheck`, `GetEndOfFile`/`GetEndOfFiles` via nkf, `ReadLastLines`); Excel/text
search (`IsSheetExistInExcel`, `SearchInExcelByPattern(...)` via ACE OLEDB, `SearchInFileWithPattern`,
`ExtracByRegex`, `ColumnNameToIndex`). Hardcoded nkf fallback `D:\share\ASS\...`; `WINDOW`/`MAC`
typed Boolean but assigned strings.

### Utils_Address.cls
Key->range mapping engine. API: `Init(targetWorksheet, addressSource As Interface_Address)`,
`UseSheet` (Set/Get), `Dictionary` (Set/Get), `Add(key, value, [isNote])`, `GetByKey(key)`,
`SetValue/GetValue`, `GetValueArray`/`SetValueArray` (errors+`Stop` on size mismatch),
`AddAllNotes`/`ClearAllNotes`. Stores sheet-qualified addresses.

### Utils_Stringify.cls
`SheetName(ws) As String` ('Sheet'! with escaping); `ReplacePlaceholder(templateText, dict) As String`
(`{{key}}` substitution).

### Utils_Error.cls
Public Enum `ECODE` + `Raise(errorCode, sourceName, message)`, `ToText(...)`, `Log(filePath, ...)`,
`Fail(errorCode, ParamArray sourceNames())`. Hardcoded log path `C:\DT_Logs\AppError.log`.

### Utils_Clipboard.cls
Dual clipboard. API: `SaveByCOM/LoadByCOM` (DataObject), `SaveByAPI/LoadByAPI` (Win32 Unicode).
Win32 declares live in a shared module.

### Utils_VBE.cls (213)
VBE IDE automation. API: `CloseProjectWindow/ClosePropertiesWindow/CloseImmediateWindow/CloseAllWindows`,
`ListComponentsByTypeName([typeName])`, `OpenComponentByName(name)`, `CloseAllComponents`,
`ClearImmediateWindowUnix/ClearImmediateWindow`, `DebugAssert(condition)`, `PrintParsed*`,
`ToggleToolbars`, `GetCommandBar(name)`. Needs "Trust access to the VBA project object model".

### Interface_Address.cls
DI contract: `Public Sub Pair(Addr As Utils_Address)` (empty; used via `Implements`).

### PJ1_Address.cls (66)
Concrete `Implements Interface_Address`. `Interface_Address_Pair(Addr)` calls `Addr.Add(...)` ~45
times mapping PJ1 keys (TASKS, *_PATTERN, *_FULLNAME, RESULT_*, CHECK_*, CONFIG_LIST) to cells.
`MAX_LINE = "229"`.

### PJ1_Logic.cls (940) - personal subsystem, NOT product
UT-log verification engine. `Init([targetWorksheet])` injects `New PJ1_Address` into `Utils_Address`.
Friend methods by purpose:
- Existence: `ChecklistExits`, `UTExits`, `LogExist`, `SfcmaplgExist`, `SfcmerlgExist`.
- Counts: `CountInUT`, `CountInOutput`.
- Encoding/EOF (nkf): `LogEOF`, `SfcmaplgEOF`, `SfcmerlgEOF`.
- Log content: `CheckLogContent`, `LastUpdate` (duplicate of CheckLogContent).
- UT-Excel extraction (ADO): `LastRuntime`, `GetTestcsh`, `GetTestcshExec`, `GetResultExec`, `GetALELog`.
- Cross-checks: `CheckSfcmaplgOutput`, `CheckSfcmerlgOutput`, `CheckSfcmerlgLine`, `CheckSfcmaplgLine`,
  `CheckSfcmerlgError`, `CheckElsePattern`.
- Asset export: `ExportImages`. Scratch: `XXXTEST`.
Hardcoded `D:\share\ASS\...`, EUC-JP, nkf - tied to one external workflow. Split from product per plan.
