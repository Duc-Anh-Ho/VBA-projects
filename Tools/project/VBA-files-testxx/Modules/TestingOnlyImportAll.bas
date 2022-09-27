Attribute VB_Name = "TestingOnlyImportAll"
Option Explicit

Private System As New PerCls
Private Auto As New AutoVBE

Public Sub testOnly()
'        System.speedOn
'        Debug.Print System.getTimerMilestone()
''''''''''''''''''
   '     Auto.exportAllVBAfiles
   '     Auto.importSelectedVBAfiles
        Auto.importAllVBAfiles
''''''''''''''''''
'        System.speedOff
        Set System = Nothing
        Set Auto = Nothing
End Sub


