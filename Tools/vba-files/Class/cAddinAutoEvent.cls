VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cAddinAutoEvent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' Make auto run event in add-in

'Initilize new app for event
Public WithEvents AddInApplication As Excel.Application
Attribute AddInApplication.VB_VarHelpID = -1

Private Sub Class_Initialize() 'Class_initalize : occurs when an instance of the associated class is created.
    Set AddInApplication = Application
End Sub

Private Sub AddInApplication_SheetActivate(ByVal Sh As Object)
    Call SheetsController.ListSheets_delete(SheetsController.previous_sheet) 'xoa listsheet khi chuyen qua sheet khac
End Sub

Private Sub AddInApplication_SheetDeactivate(ByVal Sh As Object)
    On Error Resume Next 'Case Chart sheet
    Set SheetsController.previous_sheet = Sh 'take the name of last sheet opening
    Call SheetsController.ListSheets_delete(SheetsController.previous_sheet) 'xoa listsheet khi chuyen qua sheet khac
'    'Tu dong them list sheet khi chuyen sheet (Optional)
'    Call SheetsController.ListSheets
End Sub


