VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "customEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Check README.md for more information
Option Explicit
'Declare Variables
Private info As InfoConstants
Private system As SystemUpdate
Private fileSystem As Object
Private userResponse As VbMsgBoxResult
Private WithEvents appEvent As Application
Attribute appEvent.VB_VarHelpID = -1
Private WithEvents wbEvent As Workbook
Attribute wbEvent.VB_VarHelpID = -1
Private WithEvents wsEvent As Worksheet
Attribute wsEvent.VB_VarHelpID = -1

'Constructor
Private Sub Class_Initialize()
    Set appEvent = Excel.Application
    Set wbEvent = appEvent.ThisWorkbook
    Set wsEvent = wbEvent.ActiveSheet
End Sub
'Destructor
Private Sub Class_Terminate()
    Set info = Nothing
    Set system = Nothing
    Set appEvent = Nothing
    Set wbEvent = Nothing
    Set wsEvent = Nothing
End Sub

Private Sub appEvent_SheetActivate(ByVal sheetObject As Object)
    'Refesh custom UI when Change Sheets
    Call CustomUi.refeshCustomRibbon(CustomUi.loadedRibbon)
End Sub

Private Sub appEvent_SheetDeactivate(ByVal sheetObject As Object)
    '
End Sub

Private Sub wbEvent_Open()
    '
End Sub

Private Sub wbEvent_BeforeClose(Cancel As Boolean)
    '
End Sub
