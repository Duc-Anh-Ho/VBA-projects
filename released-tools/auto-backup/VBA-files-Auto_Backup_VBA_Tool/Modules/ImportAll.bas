Attribute VB_Name = "ImportAll"
Option Explicit

Private Auto As New AutoVBE

Private Sub testOnly()
        Auto.importAllVBAfiles
        Set Auto = Nothing
End 'kill Project / delete all variables
End Sub


