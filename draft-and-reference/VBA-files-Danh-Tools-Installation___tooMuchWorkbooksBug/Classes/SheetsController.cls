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
Private info As Object
Private system As Object
Private fileSystem As Object
Private userResponse As VbMsgBoxResult
Private Const TITLE_NO As String = "~No.~"
Private Const TITLE_NAME As String = "~SHEET NAMES~"
Private Const TITLE_NO_CELL As String = "A1"
Private Const TITLE_NAME_CELL As String = "B1"
Private targetRange As Range
Private cell As Object
Private wss As Sheets
Private ws As Worksheet
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
    Set info = New MyInfo
    Set system = New SystemUpdate
    Set fileSystem = system.createFileSystem
    If Not system.hasWorkplace Then
        hasVariables = False
        GoTo ExecuteProcedure
    End If
     If TypeName(Selection) = "Range" Then
        Set targetRange = Selection
    Else
        MsgBox _
            Prompt:=info.getPrompt & _
                "You have to select range first for adding sheet(s)!", _
            Buttons:=vbOKOnly + vbExclamation, _
            Title:=info.getAuthor
        hasVariables = False
        GoTo ExecuteProcedure
    End If
    Set wss = system.Wb.Worksheets
    Set appEvent = system.app
    Set wbEvent = system.Wb
    Set wsEvent = system.ws
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
    Set info = Nothing
    Set system = Nothing
    Set appEvent = Nothing
    Set wbEvent = Nothing
    Set wsEvent = Nothing
End Sub

'ASSESSORS
'G1.0-Check Exist Sheet
Private Property Get isExistSheet(ByVal sheetName As String) As Boolean
    'Loop through all worksheets
    For Each ws In wss
        If ws.Name = sheetName Then
            Let isExistSheet = True
            Exit Property
        End If
    Next ws
    Let isExistSheet = False
End Property
'G2.0-Check Exist List Sheet
Private Property Get isExistListSheet(Optional ByVal onSheet As Worksheet) As Boolean
    If onSheet Is Nothing Then Set onSheet = system.ws
    With onSheet
        If .Range(TITLE_NO_CELL).Value = TITLE_NO _
        And .Range(TITLE_NAME_CELL) = TITLE_NAME Then
            Let isExistListSheet = True
        Else
            Let isExistListSheet = False
        End If
    End With
End Property

'MUTATORS

'METHODS
'G1.1-Add Sheets
Private Sub addSheetsAsRange()
    'Loop through selected range
    For Each cell In targetRange
        'Check don't exist and not null
        If Not isExistSheet(cell.Value) And Not cell.Value = "" Then
            'Add at the end
            wss.add(After:=wss(wss.Count)).Name = cell.Value
        End If
     Next cell
    'Return to origin sheet
    system.ws.Activate
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
        system.Wb.Save 'If not save yet it'll save in Documents
         'Loop through all worksheets
        For Each ws In wss
            'Check don't delete current sheet
            If ws.Name <> system.ws.Name Then
                ws.delete
            End If
        Next ws
    End If
End Sub
'G2.2-Insert List Sheets
Private Sub insertListAllSheets(Optional ByVal onSheet As Worksheet)
    If onSheet Is Nothing Then Set onSheet = system.ws
    If Not isExistListSheet(onSheet) Then
        With onSheet
            'Add 2 columns
            .Range(TITLE_NO_CELL, TITLE_NAME_CELL).EntireColumn.Insert
            'Rename 2 headers
            .Range(TITLE_NO_CELL).Value = TITLE_NO
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
                    ScreenTip:=CStr("Click here to go to " & ws.index & ": " & ws.Name)
            Next ws
        End With
    End If
End Sub
'G2.3-Format List Sheets
Private Sub formatListAllSheets(Optional ByVal onSheet As Worksheet)
    If onSheet Is Nothing Then Set onSheet = system.ws
    With onSheet
        With .Range(TITLE_NO_CELL, TITLE_NAME_CELL).EntireColumn
            .Font.Bold = True
            .Font.Size = 14
            .Font.Name = "Calibri"
            .AutoFit
        End With
        With .Range(TITLE_NO_CELL).EntireColumn
            .HorizontalAlignment = xlCenter
        End With
        With .Range(TITLE_NAME_CELL).EntireColumn
            .HorizontalAlignment = xlLeft
        End With
    End With
End Sub
'G2.4-Delete List Sheets
Private Sub deleteListAllSheets(Optional ByVal onSheet As Worksheet)
    If onSheet Is Nothing Then Set onSheet = system.ws
    If isExistListSheet(onSheet) Then
        With onSheet
                    'Delete first 2 column
                    .Range(TITLE_NO_CELL, TITLE_NAME_CELL).EntireColumn.delete
        End With
    End If
End Sub
'G2.5-Trigger List Sheets
Private Function triggerListAllSheets(Optional ByVal onSheet As Worksheet)
    If onSheet Is Nothing Then Set onSheet = system.ws
    If isExistListSheet(onSheet) Then
        Call deleteListAllSheets(onSheet)
        Let triggerFlag = False
    Else
        Call insertListAllSheets(onSheet)
        Call formatListAllSheets(onSheet)
        Let triggerFlag = True
    End If
End Function
'G3.1-Show/Hide/Veryhide All Sheets
Private Sub hideAllSheets(ByVal isHide As Boolean, Optional ByVal isVeryHide As Boolean = False)
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
        'Check don't delete current sheet
        If ws.Name <> system.ws.Name Then
            ws.Visible = visibleStatus
        End If
    Next ws
End Sub

'MAIN
'M1-1
Public Sub add()
On Error GoTo ErrorHandle
    If Not hasVariables Then GoTo ExecuteProcedure
    Call system.speedOn
    Call addSheetsAsRange
    Call system.speedOff
GoTo ExecuteProcedure
ErrorHandle:
    Call system.tackleErrors
ExecuteProcedure:
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
    Call system.speedOff
GoTo ExecuteProcedure
ErrorHandle:
    Call system.tackleErrors
ExecuteProcedure:
End Sub
'M3
Public Sub hide(ByVal isHide As Boolean, Optional ByVal isVeryHide As Boolean = False)
On Error GoTo ErrorHandle
    If Not hasVariables Then GoTo ExecuteProcedure
    Call system.speedOn
    Call hideAllSheets(isHide, isVeryHide)
    'Call hideAllSheets(true)
    Call system.speedOff
GoTo ExecuteProcedure
ErrorHandle:
    Call system.tackleErrors
ExecuteProcedure:
End Sub

'EVENT

Private Sub appEvent_SheetActivate(ByVal sheetObject As Object)
    If triggerFlag Then
        Call insertListAllSheets(sheetObject)
        Call formatListAllSheets(sheetObject)
    Else
        Call Class_Terminate
    End If
End Sub

Private Sub appEvent_SheetDeactivate(ByVal sheetObject As Object)
    If isExistListSheet(sheetObject) Then
        Call deleteListAllSheets(sheetObject)
    End If
End Sub
