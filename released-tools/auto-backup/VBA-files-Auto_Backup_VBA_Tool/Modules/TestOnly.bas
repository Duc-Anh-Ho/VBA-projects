Attribute VB_Name = "TestOnly"
Option Explicit

Private Auto As New AutoVBE

Public Sub testOnly()
        Auto.exportAllVBAfiles
        'Auto.importSelectedVBAfiles
        Auto.importAllVBAfiles
        Set Auto = Nothing
End 'kill Project / delete all variables
End Sub

