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
Private isOnlySh As Boolean
Private previousWb As Workbook

'Private nextWb As Workbook
'Constructor
Private Sub Class_Initialize()
    Set system = New SystemUpdate
    Set info = New InfoConstants
    Set appEvent = Excel.Application
    Set wbEvent = appEvent.ThisWorkbook
    Set wsEvent = wbEvent.ActiveSheet
    Let isOnlySh = True
End Sub
'Destructor
Private Sub Class_Terminate()
    Set info = Nothing
    Set system = Nothing
    Set appEvent = Nothing
    Set wbEvent = Nothing
    Set wsEvent = Nothing
End Sub

Private Sub appEvent_SheetActivate(ByVal ws As Object)
    ' Refresh Custom UI when Change Sheets / And New sheet
    ' NOTE: SheetActivate Event DO NOT work with special Worksheet Type
    If appEvent.Workbooks.Count > 0 Then
        Let CustomUi.hasWorksheet = system.hasWorkPlace(False, "xlWorksheet")
        Let CustomUi.hasWorkChart = system.hasWorkPlace(False, "Chart")
        Let CustomUi.hasWorkDialog = system.hasWorkPlace(False, "DialogSheet")
        Call CustomUi.refreshCustomRibbon(CustomUi.loadedRibbon)
    End If
End Sub

Private Sub appEvent_SheetDeactivate(ByVal ws As Object)
    ' *Case Deactivate normal Worksheet type and active a special one (xlChart, xlExcel4MacroSheet...)
    If Not system.hasWorkPlace(False, "xlWorksheet") Then
        Call appEvent_SheetActivate(Application.ActiveSheet)
    End If
End Sub

Private Sub appEvent_WorkbookBeforeClose(ByVal wb As Workbook, Cancel As Boolean)
    'Refresh Custom UI  when close all workbooks
    If appEvent.Workbooks.Count = 1 Then
        Let CustomUi.hasWorksheet = False
        Let CustomUi.hasWorkChart = False
        Let CustomUi.hasWorkDialog = False
        Call CustomUi.refreshCustomRibbon(CustomUi.loadedRibbon)
    End If
End Sub

Private Sub appEvent_WorkbookActivate(ByVal wb As Workbook)
    ' Reset Custom UI when open first workbook
    If appEvent.Workbooks.Count = 1 And isOnlySh Then
        Set CustomUi.loadedRibbon = Nothing
        Call CustomUi.refreshCustomRibbon
    End If
    
    'Check case deleted workbook to reset  Custom UI
    ' *Can't check Is Nothing because after remove previousWb, somehow, it's become object type not workbook or nothing
    If typeName(previousWb) <> "Workbook" Then
        Set CustomUi.loadedRibbon = Nothing
        Call CustomUi.refreshCustomRibbon
    Else
        ' Check case event enable to prevent change Workbook
        If CustomUi.hasListSheet Then
            Call popup("List Sheets", previousWb)
        ElseIf CustomUi.hasSYNCPivot Then
            Call popup("SYNC Pivot", previousWb)
        ElseIf CustomUi.hasHighlight Then
            Call popup("Highlight Range", previousWb)
        ElseIf CustomUi.isArranging Then
            Call popup("Arrange", previousWb)
        ElseIf CustomUi.isAutoArrange Then
            Call popup("Auto Arrange", previousWb)
        End If
    End If
End Sub

Private Sub appEvent_WorkbookDeactivate(ByVal wb As Workbook)
    'Check case rest sheet is delete from 2 sheet to 1 not open new
    If appEvent.Workbooks.Count = 1 Then
        Let isOnlySh = True
    Else
        Let isOnlySh = False
    End If
   Set previousWb = wb
End Sub

Private Sub popup(ByVal buttonName As String, ByRef wb As Workbook)
    Let userResponse = MsgBox( _
         Prompt:= _
             "Please turn off " & _
             Chr(34) & buttonName & Chr(34) & _
             " on DANH Tools Tab before changing Workbook!", _
         Buttons:=vbOKOnly + vbExclamation, _
         Title:=info.getAddinShortName)
    ' Disable Event prevent Stack Overflow
    Application.EnableEvents = False
    wb.Activate
    Application.EnableEvents = True
End Sub



