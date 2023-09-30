VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SheetsController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Check README.md for more information
Option Explicit
'Declare Variables
Private info As InfoConstants
Private system As SystemUpdate
Private fileSystem As Object
Private userResponse As VbMsgBoxResult
Private Const TITLE_NO As String = "~No.~"
Private Const TITLE_NAME As String = "~SHEET NAMES~"
Private Const TITLE_RENAME As String = "~RENAME~"
Private Const TITLE_NO_CELL As String = "A1"
Private Const TITLE_NAME_CELL As String = "B1"
Private Const TITLE_RENAME_CELL As String = "C1"
Private Const SELECTED_SHEET_COLOR As Long = 3390975 'RGB(255, 189, 51)
Private Const LIST_ISBOLD As Boolean = True
Private Const LIST_FONT_SIZE As Byte = 14
Private Const LIST_FONT_NAME As String = "Calibri"
Private Const LIST_FORMAT_TEXT As String = "@"
Private Const LIST_HEADER_COLOR As Long = 12566463 'Gray 25%
Private Const LIST_BORDER_COLOR As Long = -65281 'Pink
Private Const LIST_BORDER_STYLE As Long = xlDouble
Private Const MAX_SHEETNAME_LENGTH As Byte = 30
Private addSheetRange As Range
Private listSheetRange As Range
Private renameSheetRange As Range
Private preventchangedFlag As String
Private cell As Object
Private wss As Sheets
Private selectedWss As Sheets
Private ws As Object
Private WithEvents wsEvent As Worksheet
Attribute wsEvent.VB_VarHelpID = -1
Private WithEvents appEvent As Application
Attribute appEvent.VB_VarHelpID = -1
Private WithEvents wbEvent As Workbook
Attribute wbEvent.VB_VarHelpID = -1

'Initialize Variables
Private Function hasVariables() As Boolean
On Error GoTo ErrorHandle
    Set info = New InfoConstants
    Set system = New SystemUpdate
    Set fileSystem = system.createFileSystem
    If Not system.hasWorkPlace(hasMsg:=True, workPlaceType:="xlWorksheet") Then
        Let hasVariables = False
        GoTo ExecuteProcedure
    End If
    Set appEvent = system.app
    Set wbEvent = system.wb
'    Set wss = system.wb.Worksheets
    Set wss = system.wb.Sheets
    Set wsEvent = system.ws
    Set selectedWss = system.app.Windows(wbEvent.name).SelectedSheets
    Let hasVariables = True
    GoTo ExecuteProcedure
ErrorHandle:
    Call system.tackleErrors
ExecuteProcedure:
End Function

'Constructor
Private Sub Class_Initialize()
'
End Sub
'Destructor
Private Sub Class_Terminate()
'    Set info = Nothing
'    Set system = Nothing
    Set appEvent = Nothing
    Set wbEvent = Nothing
    Set wsEvent = Nothing
End Sub

'ASSESSORS
'G1.1- Remove special character in sheet name
Private Property Get removeSpecialChars(ByVal cellText As String) As String
    Let removeSpecialChars = Replace(cellText, "\", "")
    Let removeSpecialChars = Replace(removeSpecialChars, "/", "")
    Let removeSpecialChars = Replace(removeSpecialChars, "?", "")
    Let removeSpecialChars = Replace(removeSpecialChars, "*", "")
    Let removeSpecialChars = Replace(removeSpecialChars, "[", "")
    Let removeSpecialChars = Replace(removeSpecialChars, "]", "")
End Property
'G1.2- Trim if too long name
Private Property Get trimSheetName(ByVal cellText As String, Optional ByVal numChar As Byte = MAX_SHEETNAME_LENGTH) As String
    If Len(cellText) >= numChar Then
        Let trimSheetName = Left(cellText, numChar)
    Else
        Let trimSheetName = cellText
    End If
End Property
'G2.0-Check Exist List Sheet
Public Property Get hasListSheet(Optional ByVal onSheet As Worksheet) As Boolean
    If onSheet Is Nothing Then Set onSheet = system.ws
    With onSheet
        If .Range(TITLE_NO_CELL).text = TITLE_NO _
        And .Range(TITLE_NAME_CELL) = TITLE_NAME Then
            Let hasListSheet = True
        Else
            Let hasListSheet = False
        End If
    End With
End Property
'G2.1-Check Exist Rename Sheet
Public Property Get hasRenameSheet(Optional ByVal onSheet As Worksheet) As Boolean
    If onSheet Is Nothing Then Set onSheet = system.ws
    With onSheet
        If .Range(TITLE_RENAME_CELL).text = TITLE_RENAME Then
            Let hasRenameSheet = True
        Else
            Let hasRenameSheet = False
        End If
    End With
End Property
'G2.3.2-Get Raname Range
Private Property Get getRenameRange(Optional ByVal onSheet As Worksheet) As Range
    If onSheet Is Nothing Then Set onSheet = system.ws
    Set getRenameRange = onSheet.Range( _
        Cell1:=Cells( _
            Range(TITLE_RENAME_CELL).row + 1, _
            Range(TITLE_RENAME_CELL).Column), _
        Cell2:=Cells( _
            system.getLastRow(onSheet, Range(TITLE_RENAME_CELL).Column), _
            Range(TITLE_RENAME_CELL).Column))
End Property
'MUTATORS
'METHODS
'G1.1-Add Sheets
Private Sub addSheetsAsRange()
    Dim cellText As String
    'Check Selected Range
    If typeName(Selection) = "Range" Then
        Set addSheetRange = Selection
        'Loop through selected range
        For Each cell In addSheetRange
            'Remove special chars in cell.text
            Let cellText = removeSpecialChars(cell.text)
            'Trim if sheet name too long
            Let cellText = trimSheetName(cellText, MAX_SHEETNAME_LENGTH)
            'Check don't exist and not null
            If Not system.hasSheet(cellText) And cellText <> vbNullString Then
                'Add at the end
                wss.add(After:=wss(wss.Count)).name = cellText
            End If
        Next cell
        'Return to origin sheet
        system.ws.Activate
    Else
        Let userResponse = MsgBox( _
            Prompt:=info.getPrompt & _
                "You have to select range first for adding sheet(s)!", _
            Buttons:=vbOKOnly + vbExclamation, _
            Title:=info.getAuthor)
    End If
End Sub
'G1.2-Delete Sheets
Private Sub deleteAllSheets()
    Let userResponse = MsgBox( _
        Prompt:=info.getPrompt & _
            "All sheets except activative one will be delete." & _
            vbNewLine & "We CAN NOT UNDO and AUTOSAVE." & _
            vbNewLine & "Please check carefully before click OK.", _
        Buttons:=vbOKCancel + vbExclamation, _
        Title:=info.getAuthor)
    If userResponse = vbOK Then
        'Save
        system.wb.Save 'If not save yet it'll save in Documents
        'Loop through all worksheets
        For Each ws In wss
            'Check don't delete current sheet
            If ws.name <> system.ws.name And ws.Visible = xlSheetVisible Then
                ws.Delete
            End If
        Next ws
    End If
End Sub
'G2.2-Insert List Sheets
Private Sub insertListSheetsColumn(Optional ByVal onSheet As Worksheet)
    If onSheet Is Nothing Then Set onSheet = system.ws
    If Not hasListSheet(onSheet) Then
        ' Initialize listSheetRange
        Set listSheetRange = Range( _
            Range(TITLE_NO_CELL), _
            Range(TITLE_NAME_CELL).Offset(wss.Count))
        ' Disble copy/cut before insert
        appEvent.CutCopyMode = False
        ' Add No and Name columns
        With onSheet.Range(TITLE_NO_CELL, TITLE_NAME_CELL).EntireColumn
            .Insert
        End With
        With onSheet.Range(TITLE_NO_CELL)
            ' Add No column headers
            .value = CStr(TITLE_NO)
            ' Add format
            .ClearFormats
            .Interior.Color = LIST_HEADER_COLOR
            .Font.Bold = LIST_ISBOLD
            .Font.size = LIST_FONT_SIZE
            .Font.name = LIST_FONT_NAME
            .Borders.Color = LIST_BORDER_COLOR
            .Borders.LineStyle = LIST_BORDER_STYLE
            .NumberFormat = LIST_FORMAT_TEXT
            .HorizontalAlignment = xlCenter
            .EntireColumn.AutoFit
            ' Add data validation
            .Validation.Delete
            .Validation.add _
                Type:=xlValidateCustom, _
                AlertStyle:=xlValidAlertStop, _
                Operator:=xlBetween, _
                Formula1:="=" & TITLE_NO_CELL & "=" & Chr(34) & TITLE_NO & Chr(34)
            .Validation.IgnoreBlank = False
            .Validation.InputTitle = info.getAuthor
            .Validation.ErrorTitle = info.getAuthor
            .Validation.InputMessage = "List Sheets No"
            .Validation.errorMessage = info.getPrompt & _
                "This cell auto generate by List Sheets, don't manually modify this!" _
                & vbNewLine & "Please trigger List Sheets button on Danh Tools tab to remove this column."
            .Validation.ShowInput = True
            .Validation.ShowError = True
        End With
        With onSheet.Range(TITLE_NAME_CELL)
            ' Add Name column headers
            .value = CStr(TITLE_NAME)
            ' Add format
            .ClearFormats
            .Interior.Color = LIST_HEADER_COLOR
            .Font.Bold = LIST_ISBOLD
            .Font.size = LIST_FONT_SIZE
            .Font.name = LIST_FONT_NAME
            .Borders.Color = LIST_BORDER_COLOR
            .Borders.LineStyle = LIST_BORDER_STYLE
            .NumberFormat = LIST_FORMAT_TEXT
            .HorizontalAlignment = xlCenter
            .EntireColumn.AutoFit
            ' Add data validation
            .Validation.Delete
            .Validation.add _
                Type:=xlValidateCustom, _
                AlertStyle:=xlValidAlertStop, _
                Operator:=xlBetween, _
                Formula1:="=" & TITLE_NAME_CELL & "=" & Chr(34) & TITLE_NAME & Chr(34)
            .Validation.IgnoreBlank = False
            .Validation.InputTitle = info.getAuthor
            .Validation.ErrorTitle = info.getAuthor
            .Validation.InputMessage = "List Sheets Name"
            .Validation.errorMessage = info.getPrompt & _
                "This cell auto generate by List Sheet, don't manually modify this!" _
                & vbNewLine & "Please trigger List Sheet button on Danh Tools tab to remove this column."
            .Validation.ShowInput = True
            .Validation.ShowError = True
        End With
        'Loop through all worksheets
        For Each ws In wss
            ' Add No column data
            With onSheet.Range(TITLE_NO_CELL).Offset(ws.index)
                ' Add format
                .ClearFormats
                .Font.Bold = LIST_ISBOLD
                .Font.size = LIST_FONT_SIZE
                .Font.name = LIST_FONT_NAME
                .Borders.Color = LIST_BORDER_COLOR
                .Borders.LineStyle = LIST_BORDER_STYLE
                .NumberFormat = LIST_FORMAT_TEXT
                .HorizontalAlignment = xlCenter
                 .EntireColumn.AutoFit
                ' Add Numbering to No. column
                .value = CStr(ws.index)
                ' Add data validation
                .Validation.Delete
                .Validation.add _
                    Type:=xlValidateCustom, _
                    AlertStyle:=xlValidAlertStop, _
                    Operator:=xlBetween, _
                    Formula1:="=" & .Address & "=" & Chr(34) & ws.index & Chr(34)
                .Validation.IgnoreBlank = False
                .Validation.ErrorTitle = info.getAuthor
                .Validation.errorMessage = info.getPrompt & _
                    "This cell auto generate by List Sheets, don't manually modify this!" _
                    & vbNewLine & "Please trigger List Sheets button on Danh Tools tab to remove this column."
                .Validation.ShowInput = False
                .Validation.ShowError = True
            End With
            ' Add Name coulmn data
            With onSheet.Range(TITLE_NAME_CELL).Offset(ws.index)
                ' Add Hyperlinks to Name column
                ws.Hyperlinks.add _
                    Anchor:=onSheet.Range(TITLE_NAME_CELL).Offset(ws.index), _
                    Address:="", _
                    SubAddress:="'" & ws.name & "'" & "!" & TITLE_NO_CELL, _
                    screentip:=CStr("Click here to go to " & ws.index & ": " & ws.name)
                ' Add format
                '.ClearFormats
                .Font.Bold = LIST_ISBOLD
                .Font.size = LIST_FONT_SIZE
                .Font.name = LIST_FONT_NAME
                .Borders.Color = LIST_BORDER_COLOR
                .Borders.LineStyle = LIST_BORDER_STYLE
                .NumberFormat = LIST_FORMAT_TEXT
                .HorizontalAlignment = xlLeft
                ' Add SheetName to Name column
                .value = CStr(ws.name)
                ' Add data validation
                .Validation.Delete
                .Validation.add _
                    Type:=xlValidateCustom, _
                    AlertStyle:=xlValidAlertStop, _
                    Operator:=xlBetween, _
                    Formula1:="=" & .Address & "=" & Chr(34) & ws.name & Chr(34)
                .Validation.IgnoreBlank = False
                .Validation.ErrorTitle = info.getAuthor
                .Validation.errorMessage = info.getPrompt & _
                    "This cell auto generate by List Sheets, don't manually modify this!" _
                    & vbNewLine & "Please trigger List Sheets button on Danh Tools tab to remove this column."
                .Validation.ShowInput = False
                .Validation.ShowError = True
                ' Highlight current sheet
                If .text = onSheet.name Then Let .Interior.Color = SELECTED_SHEET_COLOR
            End With
        Next ws
    End If
End Sub
'G2.2.1-Insert Rename Sheets
Private Sub insertRenameColumn(Optional ByVal onSheet As Worksheet)
    If onSheet Is Nothing Then Set onSheet = system.ws
    onSheet.Unprotect
    If Not hasRenameSheet(onSheet) Then
        ' Disble copy/cut before insert
        appEvent.CutCopyMode = False
        ' Add Rename column
        With onSheet.Range(TITLE_RENAME_CELL).EntireColumn
            .Insert Shift:=xlToRight
        End With
        With onSheet.Range(TITLE_RENAME_CELL)
            ' Add Rename column header
            .value = TITLE_RENAME
            ' Add format
            .ClearFormats
            .Interior.Color = LIST_HEADER_COLOR
            .Font.Bold = LIST_ISBOLD
            .Font.size = LIST_FONT_SIZE
            .Font.name = LIST_FONT_NAME
            .Borders.Color = LIST_BORDER_COLOR
            .Borders.LineStyle = LIST_BORDER_STYLE
            .NumberFormat = LIST_FORMAT_TEXT
            .EntireColumn.AutoFit
            .HorizontalAlignment = xlCenter
            ' Add data validation
            .Validation.Delete
            .Validation.add _
                Type:=xlValidateCustom, _
                AlertStyle:=xlValidAlertStop, _
                Operator:=xlBetween, _
                Formula1:="=" & TITLE_RENAME_CELL & "=" & Chr(34) & TITLE_RENAME & Chr(34)
            .Validation.IgnoreBlank = False
            .Validation.InputTitle = info.getAuthor
            .Validation.ErrorTitle = info.getAuthor
            .Validation.InputMessage = "List Sheets No"
            .Validation.errorMessage = info.getPrompt & _
                "This cell auto generate by List Sheets, don't manually modify this!" _
                & vbNewLine & "Please trigger List Sheets button on Danh Tools tab to remove this column."
            .Validation.ShowInput = True
            .Validation.ShowError = True
        End With
        With onSheet.Range( _
            Range(TITLE_RENAME_CELL).Offset(1), _
            Range(TITLE_RENAME_CELL).Offset(wss.Count))
            ' Add Format
            .ClearFormats
            .Font.Bold = LIST_ISBOLD
            .Font.size = LIST_FONT_SIZE
            .Font.name = LIST_FONT_NAME
            .Borders.Color = LIST_BORDER_COLOR
            .Borders.LineStyle = LIST_BORDER_STYLE
            .NumberFormat = LIST_FORMAT_TEXT
            .HorizontalAlignment = xlLeft
            ' Add data validation prevent duplicate input Rename
            .Validation.Delete
            .Validation.add _
                Type:=xlValidateCustom, _
                    AlertStyle:=xlValidAlertStop, _
                    Operator:=xlBetween, _
                    Formula1:= _
                        "=COUNTIF(" _
                            & Range(TITLE_RENAME_CELL).EntireColumn.Address _
                            & "," & Range(TITLE_RENAME_CELL).Offset(1).Address(RowAbsolute:=False, ColumnAbsolute:=False) _
                        & ")= 1"
            .Validation.IgnoreBlank = True
            .Validation.ErrorTitle = info.getAuthor
            .Validation.errorMessage = info.getPrompt & _
                "Your input Sheet name can not be duplicate, please input another Sheet name!"
            .Validation.ShowInput = False
            .Validation.ShowError = True
        End With
    End If
End Sub
'G2.4-Delete List Sheets
Private Sub deleteListColumn(Optional ByVal onSheet As Worksheet)
    If onSheet Is Nothing Then Set onSheet = system.ws
    With onSheet
        .Unprotect
        If hasListSheet(onSheet) And hasRenameSheet(onSheet) Then
            'Delete first 3 columns included rename
            .Range(TITLE_NO_CELL, TITLE_RENAME_CELL).Validation.Delete
            .Range(TITLE_NO_CELL, TITLE_RENAME_CELL).EntireColumn.Delete
        ElseIf hasListSheet(onSheet) Then
            'Delete first 2 column
            .Range(TITLE_NO_CELL, TITLE_NAME_CELL).Validation.Delete
            .Range(TITLE_NO_CELL, TITLE_NAME_CELL).EntireColumn.Delete
        End If
    End With
End Sub
'G2.5-Trigger List Sheets
Private Sub triggerListAllSheets( _
    Optional ByVal onSheet As Worksheet)
    If onSheet Is Nothing Then Set onSheet = system.ws
    If hasListSheet(onSheet) Then
        Call deleteListColumn(onSheet)
        Call Class_Terminate 'Clear Event
    Else
        Call insertListSheetsColumn(onSheet)
    End If
End Sub
'G2.5- Rename Sheets
Private Sub renameSheet( _
    Optional ByVal onSheet As Worksheet)
    Dim newName As String
    If onSheet Is Nothing Then Set onSheet = system.ws
    If hasListSheet(onSheet) Then
        If hasRenameSheet(onSheet) Then
            ' TO-DO: Check data flow order to know why need re-initilize
            ' Add to listSheetRange
            Set listSheetRange = Union( _
                Range( _
                    Range(TITLE_NO_CELL), _
                    Range(TITLE_NAME_CELL).Offset(wss.Count)), _
                Range(TITLE_RENAME_CELL))
            ' Initialize renameSheetRange
            Set renameSheetRange = Range( _
                Range(TITLE_RENAME_CELL).Offset(1), _
                Range(TITLE_RENAME_CELL).Offset(wss.Count))
            ' Change to dump name for prevent duplication
            For Each cell In renameSheetRange
                'Remove special chars in newName
                Let newName = removeSpecialChars(cell.text)
                'Trim if newName too long
                Let newName = trimSheetName(newName, MAX_SHEETNAME_LENGTH)
                'Check don't exist and not null
                If Not system.hasSheet(info.getAuthor & "___" & newName) And newName <> vbNullString Then
                    ' Change to dump name
                    Let wss(cell.row - 1).name = info.getAuthor & "___" & newName
                  End If
            Next cell
            ' Rename Sheets
            For Each cell In renameSheetRange
                'Remove special chars in newName
                Let newName = removeSpecialChars(cell.text)
                'Trim if newName too long
                Let newName = trimSheetName(newName, MAX_SHEETNAME_LENGTH)
                'Check don't exist and not null
                If Not system.hasSheet(newName) And newName <> vbNullString Then
                    ' Change to dump name
                    Let wss(cell.row - 1).name = newName
                    Let onSheet.Range(TITLE_NAME_CELL).Offset(cell.row - 1).value = CStr(newName)
                End If
            Next cell
            ' Update list sheet
            Call deleteListColumn(onSheet)
            Call insertListSheetsColumn(onSheet)
        Else
            Call insertRenameColumn(onSheet)
            ' Add to listSheetRange
            Set listSheetRange = Union( _
                Range( _
                    Range(TITLE_NO_CELL), _
                    Range(TITLE_NAME_CELL).Offset(wss.Count)), _
                Range(TITLE_RENAME_CELL))
            ' Initialize renameSheetRange
            Set renameSheetRange = Range( _
                Range(TITLE_RENAME_CELL).Offset(1), _
                Range(TITLE_RENAME_CELL).Offset(wss.Count))
        End If
    Else
        Let userResponse = MsgBox(Prompt:=info.getPrompt & _
                "You have to enable list sheets first to use rename sheets!", _
            Buttons:=vbOKOnly + vbExclamation, _
            Title:=info.getAuthor)
    End If
End Sub
'G3.1-Show/Hide/Veryhide All Sheets
Private Sub hideAllSheets( _
    ByVal isHide As Boolean, _
    Optional ByVal isVeryHide As Boolean = False)
    Dim visibleStatus As Integer
    'Check condition show / hide / very hide
    If Not isHide Then
        visibleStatus = xlSheetVisible
    ElseIf isHide And isVeryHide Then
        visibleStatus = xlSheetVeryHidden
    ElseIf isHide And Not isVeryHide Then
        visibleStatus = xlSheetHidden
    End If
    'Loop through all worksheets
    For Each ws In wss
        'Check don't hide current sheet
        If ws.name <> system.ws.name Then
            ws.Visible = visibleStatus 'Hide all except current sheet
        End If
    Next ws
    '*'Loop through all selected worksheets
    For Each ws In selectedWss
            ws.Visible = xlSheetVisible 'Show all selected sheets again
    Next ws
End Sub
'MAIN
'M1-1
Public Sub add()
On Error GoTo ErrorHandle
    If Not hasVariables Then GoTo ExecuteProcedure
    Call system.speedOn
    Call addSheetsAsRange
GoTo ExecuteProcedure
ErrorHandle:
    Call system.tackleErrors
ExecuteProcedure:
    Call system.speedOff
End Sub
'M1-2
Public Sub deleteAll()
On Error GoTo ErrorHandle
    If Not hasVariables Then GoTo ExecuteProcedure
    Call system.speedOn
    Call deleteAllSheets
    Call system.speedOff
GoTo ExecuteProcedure
ErrorHandle:
    Call system.tackleErrors
ExecuteProcedure:
End Sub
'M2-1
Public Sub list(Optional ByVal onSheet As Worksheet)
On Error GoTo ErrorHandle
    If Not hasVariables Then GoTo ExecuteProcedure
    Call system.speedOn
    Call triggerListAllSheets(onSheet)
GoTo ExecuteProcedure
ErrorHandle:
    Call system.tackleErrors
ExecuteProcedure:
    Call system.speedOff
End Sub
'M2-2
Public Sub rename(Optional ByVal onSheet As Worksheet)
On Error GoTo ErrorHandle
    If Not hasVariables Then GoTo ExecuteProcedure
    Call system.speedOn
    Call renameSheet(onSheet)
GoTo ExecuteProcedure
ErrorHandle:
    Call system.tackleErrors
ExecuteProcedure:
    Call system.speedOff
End Sub
'M3
Public Sub hide(ByVal isHide As Boolean, Optional ByVal isVeryHide As Boolean = False)
On Error GoTo ErrorHandle
    If Not hasVariables Then GoTo ExecuteProcedure
    Call system.speedOn
    Call hideAllSheets(isHide, isVeryHide)
    'Call hideAllSheets(true)
GoTo ExecuteProcedure
ErrorHandle:
    Call system.tackleErrors
ExecuteProcedure:
    Call system.speedOff
End Sub

Public Sub selectNext()
    If ActiveSheet Is Nothing Then Exit Sub
    If Sheets Is Nothing Then Exit Sub
    If ActiveSheet.index = Sheets.Count Then
        Sheets(1).Visible = xlSheetVisible
        Sheets(1).Select
    Else
        ActiveSheet.Next.Visible = xlSheetVisible
        ActiveSheet.Next.Select
    End If
End Sub

Public Sub selectPrevious()
    If ActiveSheet Is Nothing Then Exit Sub
    If Sheets Is Nothing Then Exit Sub
    If ActiveSheet.index = 1 Then
        Sheets(Sheets.Count).Visible = xlSheetVisible
        Sheets(Sheets.Count).Select
    Else
        ActiveSheet.Previous.Visible = xlSheetVisible
        ActiveSheet.Previous.Select
    End If
End Sub

'EVENTS
Private Sub appEvent_SheetActivate(ByVal sheetObject As Object)
On Error GoTo ErrorHandle
    Call system.speedOn
    '2 levels check previous sheet
    If (typeName(sheetObject) = "Worksheet") Then
        If sheetObject.Type = xlWorksheet Then ' xlWorksheet = -4167
            Call insertListSheetsColumn(sheetObject)
        End If
    End If
GoTo ExecuteProcedure
ErrorHandle:
    Call system.tackleErrors
ExecuteProcedure:
    Call system.speedOff
End Sub

Private Sub appEvent_SheetDeactivate(ByVal sheetObject As Object)
On Error GoTo ErrorHandle
    Call system.speedOn
    '2 levels check previous sheet
    If (typeName(sheetObject) = "Worksheet") Then
        If sheetObject.Type = xlWorksheet Then ' xlWorksheet = -4167
            Call deleteListColumn(sheetObject)
        End If
    End If
GoTo ExecuteProcedure
ErrorHandle:
    Call system.tackleErrors
ExecuteProcedure:
    Call system.speedOff
End Sub

' Conflict with appEvent_SheetActivate (Run before  )
Private Sub appEvent_SheetChange(ByVal sheetObject As Object, ByVal Target As Range)

End Sub

Private Sub wsEvent_Change(ByVal Target As Range)
    If Intersect(Target, listSheetRange) Is Nothing Then Exit Sub
'    sheetObject.CircleInvalid
    If Target.Cells.Count = 1 Then
        If preventchangedFlag = Target.text Then Exit Sub
        Let userResponse = MsgBox( _
            Prompt:=info.getPrompt & _
                "This cell auto generate by List Sheets, don't manually modify this!" _
                & vbNewLine & "Please trigger List Sheets button on Danh Tools tab to remove this column.", _
            Buttons:=vbOKCancel + vbExclamation, _
            Title:=info.getAuthor)
        Let Target.value = preventchangedFlag
    End If
End Sub

Private Sub wsEvent_SelectionChange(ByVal Target As Range)
  '2 levels check previous sheet
    If Intersect(Target, listSheetRange) Is Nothing Then
         wsEvent.Unprotect
        Exit Sub
    Else
'        sheetObject.CircleInvalid
        If Target.Cells.Count = 1 Then
            Let preventchangedFlag = Target.text
        Else
            'Prevent delete, insert, sort, format list sheet
            wsEvent.Protect _
                DrawingObjects:=False, _
                Contents:=True, _
                Scenarios:=False, _
                AllowFormattingCells:=True, _
                AllowFormattingColumns:=True, _
                AllowFormattingRows:=True, _
                AllowInsertingColumns:=False, _
                AllowInsertingRows:=False, _
                AllowDeletingColumns:=False, _
                AllowDeletingRows:=False, _
                AllowInsertingHyperlinks:=True, _
                AllowSorting:=False, _
                AllowFiltering:=False, _
                AllowUsingPivotTables:=True
        End If
    End If
End Sub
