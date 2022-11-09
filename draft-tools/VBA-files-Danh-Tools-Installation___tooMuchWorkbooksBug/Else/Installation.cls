VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Installation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Private NewAddIn As Object

Private Sub Workbook_Open()
    'Auto install DANH AddIn
    Set NewAddIn = New AutoAddin
    Set NewAddIn = Nothing
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    '
End Sub


