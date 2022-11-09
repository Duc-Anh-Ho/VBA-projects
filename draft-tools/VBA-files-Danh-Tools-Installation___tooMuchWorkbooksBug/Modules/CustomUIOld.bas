Attribute VB_Name = "CustomUIOld"
''Attribute VB_Name = "CustomUI"
'Option Explicit
'Dim YorN As VbMsgBoxResult 'for yes/no/cancel msgbox
''Callback for labelName getLabel
'Sub GetName(control As IRibbonControl, ByRef returnedVal)
'    returnedVal = "Tool:" & TOOLNAME
'End Sub
'
''Callback for labelVersion getLabel
'Sub GetVersion(control As IRibbonControl, ByRef returnedVal)
'    returnedVal = "Version: " & VERSION
'End Sub
'
''Callback for labelToday getLabel
'Sub GetToday(control As IRibbonControl, ByRef returnedVal)
'    returnedVal = "Today: " & Date
'End Sub
'
''SNIPPING TOOL
''Callback for buttonSnipping onAction
'Sub CallSnipping(control As IRibbonControl)
'    Call Snipping.Snipping
'End Sub
'
''Callback for buttonMoveAndSizeWithCells onAction
'Sub CallMoveAndSizeWithCells(control As IRibbonControl)
'    Call MoveAndSizeWithCells.Move_And_Size_With_Cells
'End Sub
'
''Callback for buttonPasteFormat onAction
'Sub CallPastFormat(control As IRibbonControl)
'    Call pasteFormat.Paste_Format
'End Sub
'
''Callback for buttonListSheet onAction
'Sub CallListSheet(control As IRibbonControl)
'    Call SheetsController.ListSheets
'End Sub
'
''SHEETS CONTROLLER
''Callback for buttonAddSheet onAction
'Sub CallAddSheet(control As IRibbonControl)
'    Call SheetsController.addSheets
'End Sub
'
''Callback for buttonDeleteSheet onAction
'Sub CallDeleteAllSheet(control As IRibbonControl)
'    Call SheetsController.DeleteSheets
'End Sub
'
''Callback for buttonUnhideSheet onAction
'Sub CallUnhideSheet(control As IRibbonControl)
'    Call SheetsController.Unhide_All_Sheets
'End Sub
'
''Callback for buttonHideSheet onAction
'Sub CallHideSheet(control As IRibbonControl)
'    Call SheetsController.Hide_All_Sheet
'End Sub
'
''Callback for buttonVeryHideSheet onAction
'Sub CallVeryHideSheet(control As IRibbonControl)
'    Call SheetsController.Very_Hide_All_Sheet
'End Sub
'
''Callback for buttonInvertColors onAction
'Sub CallInvertColors(control As IRibbonControl)
'    Call ColorTools.Invert_Color
'End Sub
'
''Callback for buttonHideShowErrLabels onAction
'Sub CallHideShowErrLabels(control As IRibbonControl)
'    Call ChartTools.hide_show_err_chart_labels
'End Sub
'
'
''OPTION
''Callback for checkBoxShorcut getPressed
'Sub checkBoxShorcut_startup(control As IRibbonControl, ByRef returnedVal)
''PURPOSE: Set the value of the Checkbox when the Ribbon tab is first activated
'End Sub
'
''Callback for checkBoxShorcut onAction
'Sub CallShortcut(control As IRibbonControl, pressed As Boolean)
''PURPOSE:Carryout an action after user clicks the checkbox
'    If pressed Then
'        Call Snipping.shortcut_add
'        Call pasteFormat.shortcut_add
'        Call MoveAndSizeWithCells.shortcut_add
'        Call ColorTools.shortcut_add
'        Call SheetsController.shortcut_add
'    Else
'        Call Snipping.shortcut_delete
'        Call pasteFormat.shortcut_delete
'        Call MoveAndSizeWithCells.shortcut_delete
'        Call ColorTools.shortcut_delete
'        Call SheetsController.shortcut_delete
'    End If
'End Sub
'
''Callback for buttonRemoveAddIn onAction
'Sub CallRemoveAddIn(control As IRibbonControl)
'    YorN = MsgBox("Do you want to delete Add-in: " & _
'    ADD_IN_NAME & _
'    " -" & VERSION & "-?", _
'    vbOKCancel + vbExclamation, _
'    "Confirmation")
'
'    If YorN = vbOK Then
'        Call AddInInstaller.Uninstall_AddIn
'    Else
'        Exit Sub
'    End If
'End Sub
'
'
