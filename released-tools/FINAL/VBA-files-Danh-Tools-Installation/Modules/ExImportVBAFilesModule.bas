Attribute VB_Name = "ExImportVBAFilesModule"
Option Explicit

Private Auto As New AutoVBE

Public Sub exportAndImportAll()
        Auto.exportAllVBAfiles
        'Auto.importSelectedVBAfiles
        Auto.importAllVBAfiles
        Set Auto = Nothing
End 'kill Project / delete all variables
End Sub

