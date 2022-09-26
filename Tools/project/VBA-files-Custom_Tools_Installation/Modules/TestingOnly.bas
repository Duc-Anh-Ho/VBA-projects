Attribute VB_Name = "TestingOnly"
Option Explicit

Private System As New PerCls
Private Auto As New AutoVBE


Public Sub testOnly()
'        System.speedOn
'        Debug.Print System.getTimerMilestone()
''''''''''''''''''
        Auto.importVBEfiles
''''''''''''''''''
'        System.speedOff
        Set System = Nothing
End Sub


