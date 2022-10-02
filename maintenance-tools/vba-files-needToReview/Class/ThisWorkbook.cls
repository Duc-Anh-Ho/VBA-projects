VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

'Old code not for addin -> Move to Class
'Private Sub Workbook_SheetActivate(ByVal Sh As Object)
'    Call SheetsController.ListSheets_delete(SheetsController.previous_sheet) 'xoa listsheet khi chuyen qua sheet khac
'End Sub
'
'Private Sub Workbook_SheetDeactivate(ByVal Sh As Object)
'    On Error Resume Next 'Case Chart sheet
'    Set SheetsController.previous_sheet = Sh 'take the name of last sheet opening
'    Call SheetsController.ListSheets_delete(SheetsController.previous_sheet) 'xoa listsheet khi chuyen qua sheet khac
''    'Tu dong them list sheet khi chuyen sheet (Optional)
''    Call SheetsController.ListSheets
'End Sub

'' Prioritize run default event
Private Sub Workbook_Open()
    If Not ThisWorkbook.Name = ADD_IN_NAME & ".xlam" Then 'case file is install
        Application.Windows(ThisWorkbook.Name).Visible = False
        Application.ScreenUpdating = False
        AddInInstaller.Install_AddIn
    Else 'case file is add-in
        Application.ScreenUpdating = False
        ThisWorkbook.ChangeFileAccess xlReadOnly 'change to readonly in order to delete when it opened
        Call mAddInAutoEnvent.Auto_open 'because workbook_open event will run first so it'll not run Auto_open
    End If
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    If Not ThisWorkbook.Name = ADD_IN_NAME & ".xlam" Then 'case file is install
        Application.ScreenUpdating = True
    Else 'case file is add-in
        If Not Len(AddInInstaller.targetPathAndFile) = 0 Then 'Case close file only
            Kill (AddInInstaller.targetPathAndFile) 'delete the old file || This step have to change file to read only inorder to delete
        End If
        Application.ScreenUpdating = True
    End If
End Sub

'' Overwrite event of application by new Class



