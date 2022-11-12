Attribute VB_Name = "CustomUi"
' Check README.md for more information

Option Explicit
'Declare Variables
Private system As Object
Private info As Object
Private fileSystem As Object
Private userResponse As VbMsgBoxResult
Private NewAddIn As Object
Private fileC As FilesController
Private shC As SheetsController
Private rangeC As RangesController
Private picC As PicturesController


'Callback for add-sheets getEnabled
Public Sub checkWorkPlace(control As IRibbonControl, ByRef returnedVal)
    On Error GoTo ErrorHandle
    Set info = New MyInfo
    Set system = New SystemUpdate
    Set fileSystem = system.createFileSystem
    If Not system.hasWorkplace Then
        Let returnedVal = False
        GoTo ExecuteProcedure
    End If
    returnedVal = True
    GoTo ExecuteProcedure
ErrorHandle:
    Call system.tackleErrors
ExecuteProcedure:
End Sub

'Callback for add-sheets onAction
Public Sub addSheet(control As IRibbonControl)
End Sub

'TEST
Public Sub testUI()
    'Install add-in
'    Set NewAddIn = New AutoAddin
'    Set NewAddIn = Nothing
    Set fileC = New FilesController
    Call fileC.exportAllVBAfiles
'    Call fileC.importAllVBAfiles
End Sub

