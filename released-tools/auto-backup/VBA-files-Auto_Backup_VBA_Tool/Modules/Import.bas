Attribute VB_Name = "Import"
Option Explicit

Private Auto As New AutoVBE

Private Sub testOnly()
        Auto.importSelectedVBAfiles
        Set Auto = Nothing
End 'kill Project / delete all variables
End Sub

