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

Public Sub toggleZenMode()
On Error GoTo ErrorHandle
    Set info = New InfoConstants
    Set system = New SystemUpdate
    Call system.speedOn
    Dim isZen As Boolean:  Let isZen = Not ActiveWindow.DisplayHorizontalScrollBar
    With ActiveWindow
        .DisplayHeadings = isZen
        .DisplayWorkbookTabs = isZen
        .DisplayRuler = isZen
'        .DisplayFormulas = isZen
        .DisplayGridlines = isZen
        .DisplayHorizontalScrollBar = isZen
        .DisplayVerticalScrollBar = isZen
    End With
    With Application
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

Public Sub toggleViewMode()
On Error GoTo ErrorHandle
GoTo ExecuteProcedure
    'TODO: toggle zoom view Readme.md > Trello for information
ErrorHandle:
    Call system.tackleErrors
ExecuteProcedure:
    Call system.speedOff
End Sub
