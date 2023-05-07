VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PivotTablesController"
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
    If Not system.hasWorkPlace(hasMsg:=True, sheetType:="Worksheet") Then
        Let hasVariables = False
        GoTo ExecuteProcedure
    End If
    Set appEvent = system.app
    Set wbEvent = system.wb
    Set wsEvent = system.ws
    Let hasVariables = True
    GoTo ExecuteProcedure
ErrorHandle:
    Call system.tackleErrors
ExecuteProcedure:
End Function

'Constructor
Private Sub Class_Initialize()
    Call refreshAllPivotTableCaches
End Sub
'Destructor
Private Sub Class_Terminate()
    Set appEvent = Nothing
    Set wbEvent = Nothing
    Set wsEvent = Nothing
End Sub

'ASSESSORS

'MUTATORS
'METHODS

'MAIN
Public Sub refreshAllPivotTableCaches()
On Error GoTo ErrorHandle
    If Not hasVariables Then GoTo ExecuteProcedure
    Call system.speedOn
    Dim pivotTableItem As PivotCache
    For Each pivotTableItem In system.wb.PivotCaches
        pivotTableItem.Refresh
    Next pivotTableItem
GoTo ExecuteProcedure
ErrorHandle:
    Call system.tackleErrors
ExecuteProcedure:
   Call system.speedOff
End Sub

'EVENT

Private Sub appEvent_SheetChange(ByVal sheetObject As Object, ByVal target As Range)
    Call refreshAllPivotTableCaches
End Sub
