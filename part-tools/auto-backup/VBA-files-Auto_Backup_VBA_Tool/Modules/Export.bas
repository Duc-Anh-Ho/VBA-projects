Attribute VB_Name = "Export"
Option Explicit

Private Auto As New AutoVBE

Private Sub testOnly()
        Auto.exportAllVBAfiles
        Set Auto = Nothing
End 'kill Project / delete all variables
End Sub

