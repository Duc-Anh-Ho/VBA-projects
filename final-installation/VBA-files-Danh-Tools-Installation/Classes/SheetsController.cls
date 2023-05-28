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
Private Const TITLE_NO_CELL As String = "A1"
Private Const TITLE_NAME_CELL As String = "B1"
Private Const SELECTED_SHEET_COLOR As Long = 3390975 'RGB(255, 189, 51)
Private Const LIST_ISBOLD As Boolean = True
Private Const LIST_FONT_SIZE As Byte = 14
Private Const LIST_FONT_NAME As String = "Calibri"
Private targetRange As Range
Private cell As Object
Private wss As Sheets
Private selectedWss As Sheets
Private ws As Object
Private triggerFlag As Boolean
Private WithEvents appEvent As Application
Attribute appEvent.VB_VarHelpID = -1
Private WithEvents wbEvent As Workbook
Attribute wbEvent.VB_VarHelpID = -1
Private WithEvents wsEvent As Worksheet
Attribute wsEvent.VB_VarHelpID = -1

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
    Set selectedWss = system.app.Windows(wbEvent.Name).SelectedSheets
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
'G1.0-Check Exist Sheet 'TODO: Delete - Duplicated with system update hasworkplace
'Private Property Get hasSheet(ByVal sheetName As String) As Boolean
'    'Loop through all worksheets
'    For Each ws In wss
'        If ws.Name = sheetName Then
'            Let hasSheet = True
'            Exit Property
'        End If
'    Next ws
'    Let hasSheet = False
'End Property
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
Private Property Get trimSheetName(ByVal cellText As String, Optional ByVal numChar As Byte = 30) As String
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
'G2.1-Get Trigger
Private Property Get getTriggerFlag() As Boolean
    getTriggerFlag = triggerFlag
End Property
'G2.3.1-Get Sheet Name Range
Private Property Get getNameRange(Optional ByVal onSheet As Worksheet) As Range
    If onSheet Is Nothing Then Set onSheet = system.ws
    Set getNameRange = onSheet.Range( _
         Cell1:=Cells( _
            Range(TITLE_NAME_CELL).row + 1, _
            Range(TITLE_NAME_CELL).Column), _
        Cell2:=Cells( _
            system.getLastRow(onSheet, Range(TITLE_NAME_CELL).Column), _
            Range(TITLE_NAME_CELL).Column))
End Property
'MUTATORS
'G2.2-Set Trigger
Private Property Let letTriggerFlag(ByVal newTriggerFlag As Boolean)
    Let triggerFlag = newTriggerFlag
End Property
'METHODS
'G1.1-Add Sheets
Private Sub addSheetsAsRange()
    Dim cellText As String
    'Check Selected Range
    If TypeName(Selection) = "Range" Then
        Set targetRange = Selection
        'Loop through selected range
        For Each cell In targetRange
            Let cellText = cell.text
            'Remove special chars in cell.text
            Let cellText = removeSpecialChars(cellText)
            'Trim if sheet name too long
            Let cellText = trimSheetName(cellText, 30)
            'Check don't exist and not null
            If Not system.hasSheet(cellText) And cellText <> vbNullString Then
                'Add at the end
                wss.add(After:=wss(wss.Count)).Name = cellText
            End If
         Next cell
        'Return to origin sheet
        system.ws.Activate
    Else
        MsgBox _
            Prompt:=info.getPrompt & _
                "You have to select range first for adding sheet(s)!", _
            Buttons:=vbOKOnly + vbExclamation, _
            Title:=info.getAuthor
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
            If ws.Name <> system.ws.Name And ws.Visible = xlSheetVisible Then
                ws.Delete
            End If
        Next ws
    End If
End Sub
'G2.2-Insert List Sheets
Private Sub insertListAllSheets(Optional ByVal onSheet As Worksheet)
    If onSheet Is Nothing Then Set onSheet = system.ws
    If Not hasListSheet(onSheet) Then
        With onSheet
            'Add 2 columns
            .Range(TITLE_NO_CELL, TITLE_NAME_CELL).EntireColumn.Insert
            'Rename 2 headers
            .Range(TITLE_NO_CELL).value = TITLE_NO
            .Range(TITLE_NAME_CELL) = TITLE_NAME
            'Loop through all worksheets
            For Each ws In wss
                'Add Numbering to No. column
                 .Range(TITLE_NO_CELL).Offset(ws.index) = ws.index
                'Add sheetName to Name column
                 .Range(TITLE_NAME_CELL).Offset(ws.index) = ws.Name
                'Add Hyperlinks to Name column
                ws.Hyperlinks.add _
                    Anchor:=.Range(TITLE_NAME_CELL).Offset(ws.index), _
                    Address:="", _
                    SubAddress:="'" & ws.Name & "'" & "!" & TITLE_NO_CELL, _
                    screentip:=CStr("Click here to go to " & ws.index & ": " & ws.Name)
            Next ws
        End With
    End If
End Sub
'G2.3-Format List Sheets
Private Sub formatListAllSheets(Optional ByVal onSheet As Worksheet)
    If onSheet Is Nothing Then Set onSheet = system.ws
    Dim cellItem As Range
    With onSheet
        'Content Format
        With .Range(TITLE_NO_CELL, TITLE_NAME_CELL).EntireColumn
            .Font.Bold = LIST_ISBOLD
            .Font.size = LIST_FONT_SIZE
            .Font.Name = LIST_FONT_NAME
            .AutoFit
        End With
        'Header Format
        With .Range(TITLE_NO_CELL).EntireColumn
            .HorizontalAlignment = xlCenter
        End With
        With .Range(TITLE_NAME_CELL).EntireColumn
            .HorizontalAlignment = xlLeft
        End With
        'Highlight currentSheet
        For Each cellItem In getNameRange(onSheet)
            If cellItem.text = .Name Then
                Let cellItem.Interior.Color = SELECTED_SHEET_COLOR
            End If
        Next cellItem
    End With
End Sub

'G2.4-Delete List Sheets
Private Sub deleteListAllSheets(Optional ByVal onSheet As Worksheet)
    If onSheet Is Nothing Then Set onSheet = system.ws
    If hasListSheet(onSheet) Then
        With onSheet
            'Delete first 2 column
            .Range(TITLE_NO_CELL, TITLE_NAME_CELL).EntireColumn.Delete
        End With
    End If
End Sub
'G2.5-Trigger List Sheets
Private Sub triggerListAllSheets( _
    Optional ByVal onSheet As Worksheet)
    If onSheet Is Nothing Then Set onSheet = system.ws
    If hasListSheet(onSheet) Then
        Call deleteListAllSheets(onSheet)
        Call Class_Terminate 'Clear Event
    Else
        Call insertListAllSheets(onSheet)
        Call formatListAllSheets(onSheet)
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
        If ws.Name <> system.ws.Name Then
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
'M2
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
    If getTriggerFlag And _
        system.hasWorkPlace(hasMsg:=True, workPlaceType:="xlWorksheet") Then
        Call insertListAllSheets(sheetObject)
        Call formatListAllSheets(sheetObject)
    Else
'        Call Class_Terminate 'Clear Event
    End If
End Sub

Private Sub appEvent_SheetDeactivate(ByVal sheetObject As Object)
    '2 levels check previous sheet
    If (TypeName(sheetObject) = "Worksheet") Then
       If sheetObject.Type = xlWorksheet Then ' xlWorksheet = -4167
            If hasListSheet(sheetObject) Then
                Call deleteListAllSheets(sheetObject)
                Let letTriggerFlag = True
            End If
        End If
    End If
End Sub
