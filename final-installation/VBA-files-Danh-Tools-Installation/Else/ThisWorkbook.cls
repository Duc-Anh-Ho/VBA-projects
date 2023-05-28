VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private newAddin As AutoAddin
Private rangeC As RangesController
Private nextRunTime As Date
'EVENTS
'Auto Launch
Private Sub Workbook_Open()
    ''''''''''''''''''''
    'NOTE: JUST FOR TESTING DON'T RELEASE
    'Auto send Email
'    If ThisWorkbook.Name = "Danh-Tools-Installation.xlsb" Then
'        Call Developer.autoSendWifi
'    End If
    ''''''''''''''''''''
    'Auto install DANH AddIn (Move to onLoad Ribbon)
    Set newAddin = New AutoAddin
    Call newAddin.install
    Set newAddin = Nothing
    'Auto install custom shortcuts
    Call Shortcuts.install
End Sub
'
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    'Auto uninstall custom shortcuts
    Call Shortcuts.uninstall
End Sub
'CUSTOM EVENTS
'Picture Controller
'2-Arrange Shape
Private Sub Auto_Arrange_Shape()
    Dim picC As PicturesController
    Set picC = New PicturesController
    Let picC.letOffset = CustomUi.offsetValue
    Let picC.letLockRatio = CustomUi.isRateLock
    Call picC.arrange(Application.Caller)
    Call picC.clearArrange
    'Toggle Ribbon Button after click shape
    Let CustomUi.isArranging = False
    Let CustomUi.arrangeButton.letImage = "SmartArtLargerShape"
    Let CustomUi.arrangeButton.letLabel = "Arrange"
    Call CustomUi.refreshCustomRibbon(CustomUi.loadedRibbon)
End Sub
'3.1-Auto Arrange Shape
Public Sub Auto_Run_Continuously()
    Dim picC As PicturesController
    Set picC = New PicturesController
    Let picC.letOffset = CustomUi.offsetValue + 3 '3 For touching border
    Let picC.letLockRatio = CustomUi.isRateLock
    Let nextRunTime = Now + TimeValue("00:00:01") * 0.625 '0.625s delay
    Application.OnTime _
        EarliestTime:=nextRunTime, _
        procedure:="ThisWorkbook.Auto_Run_Continuously", _
        Schedule:=True
    Call picC.arrangeToMerge
End Sub
'3.2-Stop Auto Arrange Shape
Public Sub Stop_Run_Continously()
    On Error Resume Next 'TO DO: Find Another WwAy
    Application.OnTime _
        EarliestTime:=nextRunTime, _
        procedure:="ThisWorkbook.Auto_Run_Continuously", _
        Schedule:=False
End Sub

