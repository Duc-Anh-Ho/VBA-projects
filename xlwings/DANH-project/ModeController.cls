VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ModeController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Check README.md for more information
Option Explicit
'Declare Variables
Private info As InfoConstants
Private system As SystemUpdate
Private userResponse As VbMsgBoxResult
Private app As Application
Private wd As Window
Private wb As Workbook
Private ws As Worksheet

'Constructor
Private Sub Class_Initialize()
    Set info = New InfoConstants
    Set system = New SystemUpdate
End Sub

'Destructor
Private Sub Class_Terminate()
    Set info = Nothing
    Set system = Nothing
End Sub

'Initialize Variables
Private Function hasVariables() As Boolean
On Error GoTo ErrorHandle
    Set info = New InfoConstants
    Set system = New SystemUpdate
'    If Not system.hasWorkPlace(hasMsg:=True, workPlaceType:="Window") Then
    If Not system.hasWorkPlace(hasMsg:=True, workPlaceType:="xlWorksheet") Then
        Let hasVariables = False
        GoTo ExecuteProcedure
    End If
    Set app = system.app
    Set wd = system.wd
    Set wb = system.wb
    Set ws = system.ws
    Let hasVariables = True
    GoTo ExecuteProcedure
ErrorHandle:
    Call system.tackleErrors
ExecuteProcedure:
End Function

Public Sub toggleZenMode()
On Error GoTo ErrorHandle
    If Not hasVariables Then GoTo ExecuteProcedure
    Call system.speedOn
    Dim isZen As Boolean:  Let isZen = Not wd.DisplayHorizontalScrollBar
    With wd
        .DisplayHeadings = isZen
        .DisplayWorkbookTabs = isZen
        .DisplayRuler = isZen
'        .DisplayFormulas = isZen
        .DisplayGridlines = isZen
        .DisplayHorizontalScrollBar = isZen
        .DisplayVerticalScrollBar = isZen
    End With
    With app
        .ExecuteExcel4Macro "show.toolbar(""Ribbon""," & isZen & ")"
        .DisplayStatusBar = isZen
        .DisplayFormulaBar = isZen
        .DisplayScrollBars = isZen
        .WindowState = xlMaximized
'        .DisplayFullScreen = isZen
    End With
GoTo ExecuteProcedure
ErrorHandle:
    Call system.tackleErrors
ExecuteProcedure:
    Call system.speedOff
End Sub

Public Sub toggleZoomMode(Optional isMax As Boolean = False)
On Error GoTo ErrorHandle
    If Not hasVariables Then GoTo ExecuteProcedure
    With wd
    If .Zoom <> 100 Then
        Let .Zoom = 100
    ' Zoom To max
    ElseIf isMax Then
        Call zoomMax
    ' Zoom to selection
    Else
        Let .Zoom = True
    End If
    ' .Zoom = IIf(.Zoom = 100, True, 100) ' Shortage way with iif
    End With
GoTo ExecuteProcedure
ErrorHandle:
    Call system.tackleErrors
ExecuteProcedure:
    Call system.speedOff
End Sub

Sub zoomMax()
    Dim lastRow As Long: Let lastRow = system.getLastRow(ws)
    Dim lastColumn As Long: Let lastColumn = system.getLastColumn(ws)
    If lastRow = 0 Or lastColumn = 0 Then Exit Sub 'Empty sheet
    Dim selected As Variant: Set selected = Selection
    Dim rg As Range: Set rg = Range(Cells(1, 1), Cells(lastRow, lastColumn))
    rg.Select
    Let wd.Zoom = True
    selected.Select
End Sub

