# Temporary path
xlwings vba edit --file "C:\Users\Administrator\AppData\Roaming\Microsoft\AddIns\Danh-Tools.xlam"
xlwings vba edit -f "C:\Users\Administrator\AppData\Roaming\Microsoft\AddIns\Danh-Tools.xlam"

# Danh-Tools for Excel

**Author:** DANH
**Version:** v2.3.4
**Last Updated:** 2024/08/26

## Introduction

Danh-Tools is a powerful and comprehensive VBA add-in for Microsoft Excel. It provides a rich set of tools to automate common tasks, manage worksheets and VBA code, and enhance the overall Excel experience. The project is built with a modular and maintainable architecture, with a clear separation of concerns between different functional areas.

## Features

### Sheet Management (`SheetsController.cls`)

This controller provides a set of tools for managing worksheets.

- **`add()`**: Adds new sheets to the workbook based on the values in the currently selected range.
- **`deleteAll()`**: Deletes all worksheets in the workbook except for the active one.
- **`list(Optional ByVal onSheet As Worksheet)`**: Toggles a temporary column on the specified sheet (or the active sheet if not specified) that lists all sheets in the workbook with hyperlinks for easy navigation.
- **`rename(Optional ByVal onSheet As Worksheet)`**: Renames sheets based on a list of new names in a worksheet column. This is designed to work in conjunction with the `list` method.
- **`hide(ByVal isHide As Boolean, Optional ByVal isVeryHide As Boolean = False)`**: Hides, very hides (cannot be unhidden from the Excel UI), or unhides all sheets except the active one.
- **`selectNext()`**: Selects the next sheet in the workbook.
- **`selectPrevious()`**: Selects the previous sheet in the workbook.

### File Management (`FilesController.cls`)

This controller provides tools for importing and exporting VBA code, allowing for easier version control and code sharing.

- **`importSelectedVBAfiles()`**: Opens a file dialog for the user to select and import multiple VBA files (`.bas`, `.frm`, `.cls`) into the project.
- **`importAllVbaFiles()`**: Imports all VBA files from a predefined folder structure (`VBA-files-<workbook_name>/Modules`, `.../Classes`, `.../Forms`, `.../Else`).
- **`exportAllVbaFiles()`**: Exports all VBA components from the project into the same predefined folder structure.
- **`copyFileName(Optional ByRef TypeName As String = "name")`**: Copies the file name, path, or full name of the active workbook to the clipboard. The `TypeName` argument can be "name", "path", "fullName", "shortName", or "extension".

### Chart Management (`ChartsController.cls`)

This controller provides tools for managing charts.

- **`hide(Optional ByVal isHide As Boolean = True)`**: Hides or shows data labels that contain Excel errors (e.g., `#DIV/0!`, `#N/A`) on the active chart.

### Formatting (`FormatController.cls`)

This controller provides a set of tools for copying, pasting, and clearing formatting and content. It is designed to work with custom keyboard shortcuts.

- **`copyFormat()`**: Copies the formatting of the selected object (range, shape, chart, etc.).
- **`pasteFormat()`**: Pastes the copied formatting to the selected object.
- **`pasteValue()`**: Pastes the copied values to the selected range.
- **`setPlacement(Optional ByVal placementStt As Byte = xlMoveAndSize)`**: Sets the placement of all shapes on the worksheet. The `placementStt` argument can be `xlMoveAndSize`, `xlMove`, or `xlFreeFloating`.
- **`clearContent()`**: Clears the content of the selected range.
- **`clearFormat()`**: Clears the formatting of the selected range.
- **`clearAll()`**: Clears both the content and formatting of the selected range.

### Picture Management (`PicturesController.cls`)

This controller provides a rich set of tools for capturing, arranging, and manipulating pictures within Excel.

- **`snipShareX(Optional savedPath As String = vbNullString)`**: Captures a screen region using ShareX, and then inserts and scales the captured image into the selected area in Excel.
- **`snip()`**: Captures a screen region using Excel's built-in snipping tool, then inserts and scales the picture into the selected area.
- **`assign()`**: Sets up an interactive arranging feature by assigning a macro to all shapes, adding a border and a marker to the selected area, and entering "Select Objects" mode.
- **`arrange(ByRef objectName As String)`**: The macro that is assigned by the `assign` method. It arranges the clicked shape within the selected area.
- **`clearArrange()`**: Clears the markers and borders created by `assign` and removes the assigned macro from the shapes.
- **`autoArrange(ByRef isOn As Boolean)`**: Toggles an "auto arrange" mode, which marks all shapes with a "tear" icon.
- **`arrangeToMerge()`**: Arranges a selected grouped object to fit within its merged cell area.

### PivotTable Management (`PivotTablesController.cls`)

This controller provides tools for managing PivotTables.

- **`refreshAllPivotTableCaches()`**: Refreshes all PivotTable caches in the active workbook. This method is also triggered automatically whenever a sheet is changed.

### PowerShell Integration (`PowerShellController.cls`)

This controller provides an interface for executing PowerShell scripts from within VBA.

- **`createPWShellCommand As String` (Property Get)**: Returns the basic PowerShell command string (`PowerShell.exe -NoLogo -WindowStyle Hidden -Command `).
- **`runScript(ByRef scrpit As String) As Byte` (Property Get)**: Executes a PowerShell script without returning any output. It returns the exit code of the script.
- **`executeScript(ByRef scrpit As String) As String` (Property Get)**: Executes a PowerShell script and returns the output from `StdOut` or `StdErr`.

### Range Manipulation (`RangesController.cls`)

This controller provides a variety of tools for manipulating and formatting ranges.

- **`invertColor()`**: Inverts the colors of the selected range (background, font, and borders).
- **`boldFirstLine()`**: Toggles the bold formatting of the first line of each cell in the selected range.
- **`highlight(ByVal target As Object)`**: Highlights the target range with a specified color, blur, and font size increase. This feature also has event handlers to automatically highlight the selected range and to store/paste the format when switching between sheets.
- **`displayPageBreak(ByRef isDisplay As Boolean, Optional isApplyAll As Boolean = False)`**: Shows or hides page breaks for the active sheet or all sheets.
- **`multipleReplace(...) As Boolean`**: A powerful function for performing multiple find-and-replace operations. It takes two ranges (find and replace), and a set of options (match case, match content, search order, etc.).

### Shortcut Management (`ShortcutController.cls`)

This controller manages custom keyboard shortcuts. The shortcuts are defined in a table on the "keyboard-shortcut" worksheet, making them user-configurable.

- **`install()`**: Applies all the custom shortcuts defined in the "keyboard-shortcut" sheet.
- **`unInstall()`**: Restores the default behavior of all the custom shortcuts.
- **`convertKeyToName(...) As String`**: Converts a key code and a shift mask into a user-friendly name (e.g., "Ctrl + Shift + A").
- **`convertCodeToName(ByRef code As String) As String`**: Converts an `Application.OnKey` string (e.g., `^+A`) into a user-friendly name.
- **`convertNameToCode(ByVal name As String) As String`**: Converts a user-friendly name back into an `Application.OnKey` string.

## Core Components

- **`SystemUpdate.cls`:** The core engine of the add-in, providing a wide range of utilities for interacting with the Excel application, the file system, and other system objects. It also includes performance optimization features and a centralized error handling mechanism.
- **`InfoConstants.cls`:** A centralized class for storing static information such as the author's name, the add-in's version, and update dates.