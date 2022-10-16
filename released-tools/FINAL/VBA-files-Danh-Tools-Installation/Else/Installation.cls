VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Installation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' AUTO ADD ADDIN TOOLS
' Author: DANH
' Version: 1.0.0
' Update: 2022/10/07
' Check README.md for more information

Option Explicit
Private Auto As Object

Private Sub Workbook_Open()
    Set Auto = New AutoAddIn
    Set Auto = Nothing
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    '
End Sub

