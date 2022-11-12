VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PivotController"
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
'Initialize Variables
Private Function hasVariables() As Boolean
On Error GoTo ErrorHandle
    Set info = New MyInfo
    Set system = New SystemUpdate
    Set fileSystem = system.createFileSystem
    If Not system.hasWorkplace Then
        Let hasVariables = False
        GoTo ExecuteProcedure
    End If
    hasVariables = True
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
'
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
    For Each pivotTableItem In system.Wb.PivotCaches
        pivotTableItem.Refresh
    Next pivotTableItem
GoTo ExecuteProcedure
ErrorHandle:
    Call system.tackleErrors
ExecuteProcedure:
   Call system.speedOff
End Sub

