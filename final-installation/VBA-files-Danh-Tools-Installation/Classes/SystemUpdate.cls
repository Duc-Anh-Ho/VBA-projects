VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SystemUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Check README.md for more information
Option Explicit
'Declare Variables
'Local
Private Const FILE_SYSTEM_SCRIPT As String = "Scripting.FileSystemObject"
'Reference: ss64.com/vb/shell.html -- 2 diffirence COM classes
Private Const POWER_SHELL_SCRIPT As String = "Wscript.Shell" '  Windows Script Host Shell object
'Private Const POWER_SHELL_SCRIPT As String = "Shell.Application" '  Shell Application object
Private Const FILE_STREAM_SCRIPT As String = "ADODB.Stream"
Private Const CDO_CONFIG_SCRIPT As String = "CDO.Configuration"
Private Const CDO_MESS_SCRIPT As String = "CDO.Message"
Private Const HTML_FILE_SCRIPT As String = "HtmlFile"
Private Const CLIPBOARD_SCRIPT As String = "New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}"
Private Const EXCEL_TYPE As String = "Microsoft Excel"
Private Const WD_TYPE As String = "Window"
Private Const WB_TYPE As String = "Workbook"
Private Const WS_TYPE As String = "Worksheet"
Private Const WS_TYPE_DETAIL_SHEET As String = "xlWorksheet"
Private Const WS_TYPE_DETAIL_CHART As String = "xlChart"
Private Const WS_TYPE_DETAIL_MACRO As String = "xlExcel4MacroSheet"
Private Const WS_TYPE_DETAIL_INT_MACRO As String = "xlExcel4IntMacroSheet"
Private Const CHART_TYPE As String = "Chart"
Private Const DIALOG_TYPE As String = "DialogSheet"
Private userResponse As VbMsgBoxResult
Private info As InfoConstants
Private systemTimer As Single
Public WithEvents app As Application
Attribute app.VB_VarHelpID = -1
Public WithEvents wb As Workbook
Attribute wb.VB_VarHelpID = -1
Public WithEvents ws As Worksheet
Attribute ws.VB_VarHelpID = -1
Public wd As Window
Attribute wd.VB_VarHelpID = -1
'Public wss As Worksheets 'TODO: Refactor as global accessor get wss for sheets controller
'Public wsObs As Sheets 'wsObs could be Charts || DialogSheets || Worksheets (ws regular || ws marcro)
Public sheetOb As Object 'SheetOb could be Chart || DialogSheet || Worksheet (ws regular || ws marcro)
Public chartSh As Chart
Public dialogSh As DialogSheet

'Initialize Variables
Private Function hasVariables() As Boolean
On Error GoTo ErrorHandle
    Set info = New InfoConstants
    Let systemTimer = Timer ' start timer
    Let hasVariables = True
    GoTo ExecuteProcedure
ErrorHandle:
    Call tackleErrors
ExecuteProcedure:
End Function

'Constructor
Private Sub Class_Initialize()
On Error GoTo ErrorHandle
    If Not hasVariables Then GoTo ExecuteProcedure
    GoTo ExecuteProcedure
ErrorHandle:
    Call tackleErrors
ExecuteProcedure:
End Sub
'Destructor
Private Sub Class_Terminate()
    '
End Sub

'ASSESSORS

'CHECK
Private Function wsDetailType() As String
    Select Case Application.ActiveSheet.Type
        Case xlWorksheet
            Let wsDetailType = WS_TYPE_DETAIL_SHEET
        Case xlChart
            Let wsDetailType = WS_TYPE_DETAIL_CHART
        Case xlExcel4MacroSheet
            Let wsDetailType = WS_TYPE_DETAIL_MACRO
        Case xlExcel4IntlMacroSheet
            Let wsDetailType = WS_TYPE_DETAIL_INT_MACRO
    End Select
End Function

Public Property Get hasWorkPlace( _
    Optional ByVal hasMsg As Boolean = False, _
    Optional ByVal workPlaceType As String _
) As Variant
On Error GoTo ErrorHandle
    Dim mess As String
    Dim resultStr As String
'    '*Case close app but used storage procedure || application.ontime run || mutilple thread
'   ' TODO
'    If not hasAppllication()  Then
'          Let mess = "Can not find any application!"
'          Let hasWorkPlace = False
'          GoTo ExecuteProcedure
'    End If
    'Check each possible cases
    Select Case workPlaceType
        'Return *STRING Cases
        Case vbNullString
            'LV-1
            'TO-DO: Make cases for MS: Work, PPT, Access, Outlook...
            If Not hasApplication(EXCEL_TYPE) Then
                Let resultStr = "NOTHING"
                GoTo ExecuteProcedure
            Else
                Select Case Application
                    Case EXCEL_TYPE
                        Set app = Application
                        Let resultStr = EXCEL_TYPE
                    Case Else
                        ' Set app = Application
                        Let resultStr = Application.name & "_Un-handle"
                        GoTo ExecuteProcedure
                End Select
            End If
            'LV-1.5
            If Not app.ActiveWindow Is Nothing Then
                GoTo ExecuteProcedure
            Else
                Select Case typeName(app.ActiveWindow)
                    Case WD_TYPE
                        Set wd = app.ActiveWindow
                        Let resultStr = resultStr & "_" & WD_TYPE
                    Case Else
                        ' Set app = Application
                        Let resultStr = resultStr & "_" & typeName(app.ActiveWindow) & "_Un-handle"
                        GoTo ExecuteProcedure
                End Select
            End If
            'LV-2
            'TO-DO: Make cases for else types: document..
            If app.ActiveWorkbook Is Nothing Then
                GoTo ExecuteProcedure
            Else
                Select Case typeName(app.ActiveWorkbook)
                    Case WB_TYPE
                        Set wb = app.ActiveWorkbook
                        Let resultStr = resultStr & "_" & WB_TYPE
                    Case Else
                        ' Set wb = app.ActiveWorkbook
                        Let resultStr = resultStr & "_" & typeName(app.ActiveWorkbook) & "_Un-Handle"
                        GoTo ExecuteProcedure
                End Select
            End If
            'LV-3
            If wb.ActiveSheet Is Nothing Then
                GoTo ExecuteProcedure
            Else
                Set sheetOb = wb.ActiveSheet
                Select Case typeName(wb.ActiveSheet)
                    Case WS_TYPE
                        Set ws = wb.ActiveSheet
                        Set chartSh = Nothing
                        Set dialogSh = Nothing
                        Let resultStr = resultStr & "_" & WS_TYPE
                    Case CHART_TYPE
                        Set ws = Nothing
                        Set chartSh = wb.ActiveSheet
                        Set dialogSh = Nothing
                        Let resultStr = resultStr & "_" & CHART_TYPE
                    Case DIALOG_TYPE
                        Set ws = Nothing
                        Set chartSh = Nothing
                        Set dialogSh = wb.ActiveSheet
                        Let resultStr = resultStr & "_" & DIALOG_TYPE
                        GoTo ExecuteProcedure ' DialogSheet DON'T have type attribute
                    Case Else
                        Let resultStr = resultStr & "_" & typeName(wb.ActiveSheet) & "_Un-Handle"
                        GoTo ExecuteProcedure
                End Select
            End If
            'LV-4
            Select Case sheetOb.Type 'Type will return number, NOT string.
                Case xlWorksheet
                    Let resultStr = resultStr & "_" & WS_TYPE_DETAIL_SHEET
                Case xlChart
                    Let resultStr = resultStr & "_" & WS_TYPE_DETAIL_CHART
                Case xlExcel4MacroSheet
                    Let resultStr = resultStr & "_" & WS_TYPE_DETAIL_MACRO
                Case xlExcel4IntlMacroSheet
                    Let resultStr = resultStr & "_" & WS_TYPE_DETAIL_INT_MACRO
                Case Else
                Let resultStr = resultStr & "_Unknow Type"
            End Select
            GoTo ExecuteProcedure
        'Return *BOOLEAN Cases
        Case EXCEL_TYPE ' LV-1 'TO-DO: Make cases for MS: Work, PPT, Access, Outlook...
            If Not hasApplication(EXCEL_TYPE) Then
                Let mess = "Sorry, this tool currently is just working with MS-EXCEL!"
                GoTo ExecuteProcedure
            Else
                Set app = Application
                GoTo ExecuteProcedure
            End If
        Case WD_TYPE  ' LV-1.5
            If Not hasApplication(EXCEL_TYPE) Then
                   Let mess = "Sorry, this tool currently is just working with MS-EXCEL!"
                   GoTo ExecuteProcedure
            ElseIf Application.ActiveWindow Is Nothing Then
                Let mess = "Don't exist any window yet, please open one!"
                GoTo ExecuteProcedure
            Else
                Set app = Application
                Set wd = app.ActiveWindow
                GoTo ExecuteProcedure
            End If
        Case WB_TYPE ' LV-2
            If Not hasApplication(EXCEL_TYPE) Then
                Let mess = "Sorry, this tool currently is just working with MS-EXCEL!"
                GoTo ExecuteProcedure
            ElseIf Application.ActiveWorkbook Is Nothing Then
                Let mess = "Don't exist any workbook yet, please open one!"
                GoTo ExecuteProcedure
            Else
                Let app = Application
                Let wb = app.ActiveWorkbook
                GoTo ExecuteProcedure
            End If
        Case WS_TYPE ' LV-3
            If Not hasApplication(EXCEL_TYPE) Then
                Let mess = "Sorry, this tool currently is just working with MS-EXCEL!"
                GoTo ExecuteProcedure
            ElseIf Application.ActiveWorkbook Is Nothing Then
                Let mess = "Don't exist any workbook yet, please open one!"
                GoTo ExecuteProcedure
            ElseIf Application.ActiveSheet Is Nothing Then
                Let mess = "Don't active any sheet yet, please select one!"
                GoTo ExecuteProcedure
            ElseIf typeName(Application.ActiveSheet) <> WS_TYPE Then
                Let mess = _
                    "You've selected a " & _
                    typeName(Application.ActiveSheet) & _
                    ", need to select a " & _
                    workPlaceType
                GoTo ExecuteProcedure
            Else
                Set app = Application
                Set wb = app.ActiveWorkbook
                Set sheetOb = wb.ActiveSheet
                'Set ws following sheetOb
                If wsDetailType() <> WS_TYPE_DETAIL_SHEET Then
                    Set ws = sheetOb
                Else
                    Set ws = Nothing
                End If
                Set chartSh = Nothing
                Set dialogSh = Nothing
                GoTo ExecuteProcedure
            End If
        Case CHART_TYPE ' LV-3
             If Not hasApplication(EXCEL_TYPE) Then
                Let mess = "Sorry, this tool currently is just working with MS-EXCEL!"
                GoTo ExecuteProcedure
            ElseIf Application.ActiveWorkbook Is Nothing Then
                Let mess = "Don't exist any workbook yet, please open one!"
                GoTo ExecuteProcedure
            ElseIf Application.ActiveSheet Is Nothing Then
                Let mess = "Don't active any sheet yet, please select one!"
                GoTo ExecuteProcedure
            ElseIf typeName(Application.ActiveSheet) <> CHART_TYPE Then
                Let mess = _
                    "You've selected a " & _
                    typeName(Application.ActiveSheet) & _
                    ", need to select a " & _
                    workPlaceType
                GoTo ExecuteProcedure
            Else
                Set app = Application
                Set wb = app.ActiveWorkbook
                Set sheetOb = Nothing
                Set chartSh = wb.ActiveSheet
                Set dialogSh = Nothing
                GoTo ExecuteProcedure
            End If
        Case DIALOG_TYPE 'LV-3
            If Not hasApplication(EXCEL_TYPE) Then
                Let mess = "Sorry, this tool currently is just working with MS-EXCEL!"
                GoTo ExecuteProcedure
            ElseIf Application.ActiveWorkbook Is Nothing Then
                Let mess = "Don't exist any workbook yet, please open one!"
                GoTo ExecuteProcedure
            ElseIf Application.ActiveSheet Is Nothing Then
                Let mess = "Don't active any sheet yet, please select one!"
                GoTo ExecuteProcedure
            ElseIf typeName(Application.ActiveSheet) <> DIALOG_TYPE Then
               Let mess = _
                    "You've selected a " & _
                    typeName(Application.ActiveSheet) & _
                    ", need to select a " & _
                    workPlaceType
                GoTo ExecuteProcedure
            Else
                Set app = Application
                Set wb = app.ActiveWorkbook
                Set sheetOb = Nothing
                Set chartSh = Nothing
                Set dialogSh = wb.ActiveSheet
                GoTo ExecuteProcedure
            End If
        Case WS_TYPE_DETAIL_SHEET 'LV-4
            If Not hasApplication(EXCEL_TYPE) Then
                Let mess = "Sorry, this tool currently is just working with MS-EXCEL!"
                GoTo ExecuteProcedure
            ElseIf Application.ActiveWorkbook Is Nothing Then
                Let mess = "Don't exist any workbook yet, please open one!"
                GoTo ExecuteProcedure
            ElseIf Application.ActiveSheet Is Nothing Then
                Let mess = "Don't active any sheet yet, please select one!"
                GoTo ExecuteProcedure
            ElseIf typeName(Application.ActiveSheet) <> WS_TYPE Then
                Let mess = _
                    "You've selected a " & _
                    typeName(Application.ActiveSheet) & _
                    ", need to select a " & _
                    workPlaceType
                GoTo ExecuteProcedure
            ElseIf wsDetailType() <> WS_TYPE_DETAIL_SHEET Then
                Let mess = _
                    "You've selected a special worksheet -" & _
                    wsDetailType() & _
                    "- you need to select a " & _
                    WS_TYPE_DETAIL_CHART & " one!"
                GoTo ExecuteProcedure
            Else
                Set app = Application
                Set wb = app.ActiveWorkbook
                Set sheetOb = wb.ActiveSheet
                Set ws = sheetOb
                Set chartSh = Nothing
                Set dialogSh = Nothing
                GoTo ExecuteProcedure
            End If
        Case WS_TYPE_DETAIL_CHART 'LV-4
            If Not hasApplication(EXCEL_TYPE) Then
                Let mess = "Sorry, this tool currently is just working with MS-EXCEL!"
                GoTo ExecuteProcedure
            ElseIf Application.ActiveWorkbook Is Nothing Then
                Let mess = "Don't exist any workbook yet, please open one!"
                GoTo ExecuteProcedure
            ElseIf Application.ActiveSheet Is Nothing Then
                Let mess = "Don't active any sheet yet, please select one!"
                GoTo ExecuteProcedure
            ElseIf typeName(Application.ActiveSheet) <> CHART_TYPE Then
                Let mess = _
                    "You've selected a " & _
                    typeName(Application.ActiveSheet) & _
                    ", need to select a " & _
                    workPlaceType
                GoTo ExecuteProcedure
            ElseIf wsDetailType() <> WS_TYPE_DETAIL_CHART Then
                Let mess = _
                    "You've selected a special worksheet -" & _
                    wsDetailType() & _
                    "- you need to select a " & _
                    WS_TYPE_DETAIL_CHART & " one!"
                GoTo ExecuteProcedure
            Else
                Set app = Application
                Set wb = app.ActiveWorkbook
                Set sheetOb = wb.ActiveSheet
                Set ws = Nothing
                Set chartSh = sheetOb
                Set dialogSh = Nothing
                GoTo ExecuteProcedure
            End If
        Case WS_TYPE_DETAIL_MACRO 'LV-4
            If Not hasApplication(EXCEL_TYPE) Then
                Let mess = "Sorry, this tool currently is just working with MS-EXCEL!"
                GoTo ExecuteProcedure
            ElseIf Application.ActiveWorkbook Is Nothing Then
                Let mess = "Don't exist any workbook yet, please open one!"
                GoTo ExecuteProcedure
            ElseIf Application.ActiveSheet Is Nothing Then
                Let mess = "Don't active any sheet yet, please select one!"
                GoTo ExecuteProcedure
            ElseIf typeName(Application.ActiveSheet) <> WS_TYPE Then
                Let mess = _
                    "You've selected a " & _
                    typeName(Application.ActiveSheet) & _
                    ", need to select a " & _
                    workPlaceType
                GoTo ExecuteProcedure
            ElseIf wsDetailType() <> WS_TYPE_DETAIL_MACRO Then
                Let mess = _
                    "You've selected a special worksheet -" & _
                    wsDetailType() & _
                    "- you need to select a " & _
                    WS_TYPE_DETAIL_MACRO & " one!"
                GoTo ExecuteProcedure
            Else
                Set app = Application
                Set wb = app.ActiveWorkbook
                Set sheetOb = wb.ActiveSheet
                Set ws = Nothing
                Set chartSh = Nothing
                Set dialogSh = Nothing
                GoTo ExecuteProcedure
            End If
        Case WS_TYPE_DETAIL_INT_MACRO 'LV-4
            If Not hasApplication(EXCEL_TYPE) Then
                Let mess = "Sorry, this tool currently is just working with MS-EXCEL!"
                GoTo ExecuteProcedure
            ElseIf Application.ActiveWorkbook Is Nothing Then
                Let mess = "Don't exist any workbook yet, please open one!"
                GoTo ExecuteProcedure
            ElseIf Application.ActiveSheet Is Nothing Then
                Let mess = "Don't active any sheet yet, please select one!"
                GoTo ExecuteProcedure
            ElseIf typeName(Application.ActiveSheet) <> WS_TYPE Then
                Let mess = _
                    "You've selected a " & _
                    typeName(Application.ActiveSheet) & _
                    ", need to select a " & _
                    workPlaceType
                GoTo ExecuteProcedure
            ElseIf wsDetailType() <> WS_TYPE_DETAIL_INT_MACRO Then
                Let mess = _
                    "You've selected a special worksheet -" & _
                    wsDetailType() & _
                    "- you need to select a " & _
                    WS_TYPE_DETAIL_INT_MACRO & " one!"
                GoTo ExecuteProcedure
            Else
                Set app = Application
                Set wb = app.ActiveWorkbook
                Set sheetOb = wb.ActiveSheet
                Set ws = Nothing
                Set chartSh = Nothing
                Set dialogSh = Nothing
                GoTo ExecuteProcedure
            End If
        Case Else
            Let resultStr = _
                "-hasWorkPlace- only accept below options: " & _
                vbNewLine & EXCEL_TYPE & _
                vbNewLine & WB_TYPE & _
                vbNewLine & WD_TYPE & _
                vbNewLine & WS_TYPE & _
                vbNewLine & CHART_TYPE & _
                vbNewLine & DIALOG_TYPE & _
                vbNewLine & WS_TYPE_DETAIL_SHEET & _
                vbNewLine & WS_TYPE_DETAIL_CHART & _
                vbNewLine & WS_TYPE_DETAIL_MACRO & _
                vbNewLine & WS_TYPE_DETAIL_INT_MACRO
            Let mess = resultStr
            GoTo ExecuteProcedure
    End Select
ErrorHandle:
'    Call tackleErrors
ExecuteProcedure:
    'Check return String Cases
    If workPlaceType = vbNullString Then
        Let hasWorkPlace = resultStr
    'Check if don't have mess mean True case
    ElseIf mess = vbNullString Then
        Let hasWorkPlace = True
    Else
        'Check display mess option
        If hasMsg Then
            Let userResponse = MsgBox( _
                Prompt:=info.getPrompt & mess, _
                Buttons:=vbOKOnly + vbExclamation, _
                Title:=info.getAuthor)
        End If
        Let hasWorkPlace = False
    End If
End Property
'APPLICATION
Public Property Get hasApplication(Optional appName As String = EXCEL_TYPE) As Boolean
    If Application.name = appName Then
        Let hasApplication = True
    Else
        Let hasApplication = False
    End If
End Property
'WORKBOOK
Public Property Get hasWorkbook(wbName As String) As Boolean
    Dim wo As Workbook
    For Each wo In app.Workbooks
        If wo.name = wbName Then
            Let hasWorkbook = True
            Exit Property
        End If
    Next wo
    Let hasWorkbook = False
End Property
'SHEET
Public Property Get hasSheet(shName As String) As Boolean
    Dim Sh As Object ' Object for Special Worksheets
    For Each Sh In wb.Sheets
        If Sh.name = shName Then
            Let hasSheet = True
            Exit Property
        End If
    Next Sh
    Let hasSheet = False
End Property
'//OBJECT CREATE BY SCRIPTS
'FILE
Public Property Get createFileSystem() As Object
    Set createFileSystem = CreateObject(FILE_SYSTEM_SCRIPT)
End Property
'SHELL/CMD
Public Property Get createPowerShell() As Object
    Set createPowerShell = CreateObject(POWER_SHELL_SCRIPT)
End Property
'STREAM
Public Property Get createFileStream() As Object
    Set createFileStream = CreateObject(FILE_STREAM_SCRIPT)
End Property
'CDO CONFIG
Public Property Get createCDOConfig() As Object
    Set createCDOConfig = CreateObject(CDO_CONFIG_SCRIPT)
End Property
'CDO MESS
Public Property Get createCDOMess() As Object
    Set createCDOMess = CreateObject(CDO_MESS_SCRIPT)
End Property
'HTMLFILE
Public Property Get createHTMLFile() As Object
    Set createHTMLFile = CreateObject(HTML_FILE_SCRIPT)
End Property
'CLIPBOARD
Public Property Get createClipboard() As Object
    Set createClipboard = CreateObject(CLIPBOARD_SCRIPT)
End Property
'TIMER
Public Property Get getTimerMilestone() As String
On Error GoTo ErrorHandle
    Let getTimerMilestone = Timer - systemTimer & " secs"
GoTo ExecuteProcedure
ErrorHandle:
    Call tackleErrors
ExecuteProcedure:
End Property
'FILE
Public Property Get getExcelPath(Optional ByVal FileFilter As String) As String
On Error GoTo ErrorHandle
    With Application
        If FileFilter = vbNullString Then
                Let FileFilter = "Excel Files (*.xls*), *xls*"
        End If
        Let getExcelPath = .GetOpenFilename( _
                Title:="Select file", _
                FileFilter:=FileFilter)
    End With
GoTo ExecuteProcedure
ErrorHandle:
    Call tackleErrors
ExecuteProcedure:
End Property
'FILE
Public Property Get getExcelFile(ByVal path As Variant) As Workbook
On Error GoTo ErrorHandle
    With Application
        If path <> False Then Set getExcelFile = .Workbooks.Open( _
                fileName:=path, _
                ReadOnly:=True)
    End With
GoTo ExecuteProcedure
ErrorHandle:
    Call tackleErrors
ExecuteProcedure:
End Property
'FILE
Public Property Get getFileName(Optional ByRef typeName As String = "name") As String
On Error GoTo ErrorHandle
    If Application.ActiveWorkbook Is Nothing Then GoTo ExecuteProcedure
    Dim fileSystem As Object
    Set fileSystem = createFileSystem()
    With Application.ActiveWorkbook
        Select Case typeName
            Case "name"
                Let getFileName = .name
            Case "path"
                Let getFileName = .path
            Case "fullName"
                Let getFileName = .FullName
            Case "shortName"
                Let getFileName = fileSystem.GetBaseName(.name)
            Case "extension"
                Let getFileName = fileSystem.GetExtensionName(.name)
    End Select
    Set fileSystem = Nothing
End With
GoTo ExecuteProcedure
ErrorHandle:
    Call tackleErrors
ExecuteProcedure:
End Property
'FOLDER
Public Property Get getFolder() As String
On Error GoTo ErrorHandle
    Dim diaFolder As FileDialog
    Dim isSelected As Boolean
    Set diaFolder = Application.FileDialog(msoFileDialogFolderPicker)
    Let diaFolder.AllowMultiSelect = False
    Let diaFolder.Title = "Select a folder"
    Let isSelected = diaFolder.Show
    If isSelected Then
         Let getFolder = diaFolder.SelectedItems(1)
    Else
        Let getFolder = vbNullString
    End If
    Set diaFolder = Nothing
GoTo ExecuteProcedure
ErrorHandle:
    Call tackleErrors
ExecuteProcedure:
End Property
'CUSTOMIZE
Public Property Get getFomulaSeparator() As String
    Let getFomulaSeparator = app.International(xlListSeparator)
End Property
'POSITION
Public Property Get getLastRow(ByVal ws As Worksheet, Optional ByVal atColumn As Long) As Long
On Error GoTo ErrorHandle
    With ws
        If atColumn = False Then
            Let getLastRow = .Cells.Find( _
                "*", _
                SearchOrder:=xlByRows, _
                SearchDirection:=xlPrevious).row
        Else
            If .Cells(.Rows.Count, atColumn).End(xlUp).value = False Then
                    getLastRow = 0
            Else
                    Let getLastRow = .Cells(.Rows.Count, atColumn).End(xlUp).row
            End If
        End If
    End With
GoTo ExecuteProcedure
ErrorHandle:
    Call tackleErrors
ExecuteProcedure:
End Property
'POSITION
Public Property Get getLastColumn(ByVal ws As Worksheet, Optional ByVal atRow As Long) As Long
On Error GoTo ErrorHandle
    With ws
        If atRow = False Then
            Let getLastColumn = ws.Cells.Find( _
                "*", _
                SearchOrder:=xlByColumns, _
                SearchDirection:=xlPrevious).Column
        Else
            If .Cells(atRow, .Columns.Count).End(xlToLeft).value = False Then
                getLastColumn = 0
            Else
                Let getLastColumn = .Cells(atRow, .Columns.Count).End(xlToLeft).Column
            End If
        End If
    End With
GoTo ExecuteProcedure
ErrorHandle:
    Call tackleErrors
ExecuteProcedure:
End Property
'ARRAY
Public Property Get getArrayLength(ByRef arr As Variant, Optional ByVal dimension As Long = -1) As Long
On Error GoTo ErrorHandle
        If IsEmpty(arr) Then
            Let getArrayLength = 0
        Else
            If dimension = -1 Then
                Let getArrayLength = UBound(arr) - LBound(arr) + 1
            Else
                Let getArrayLength = UBound(arr, dimension) - LBound(arr, dimension) + 1
            End If
        End If
GoTo ExecuteProcedure
ErrorHandle:
    Call tackleErrors
ExecuteProcedure:
End Property
'ARRAY
Public Property Get getArrayDimension(ByRef arr As Variant) As Long
    On Error GoTo OverDimensionErr
    Dim i As Long
    Dim tmp As Long
    Let i = 0
    Do While True
        Let i = i + 1
        Let tmp = UBound(arr, i)
    Loop
OverDimensionErr:
    Let getArrayDimension = i - 1
End Property
'ARRAY
Public Property Get mergeTwoArrays(ByRef arr1 As Variant, ByRef arr2 As Variant) As Variant
On Error GoTo ErrorHandle
    If IsEmpty(arr1) And IsEmpty(arr2) Then
        ' Let mergeTwoArrays = Empty
        Exit Property
    End If
    If IsEmpty(arr1) Then
        Let mergeTwoArrays = arr2
        Exit Property
    End If
    If IsEmpty(arr2) Then
        Let mergeTwoArrays = arr1
        Exit Property
    End If
OneDimension:
    If (getArrayDimension(arr1) < 2 Or getArrayDimension(arr2) < 2) Then
        Let mergeTwoArrays = Split(Join(arr1, Chr(1)) & Chr(1) & Join(arr2, Chr(1)), Chr(1))
        Exit Property
    End If
ManyDimensions:
    If getArrayDimension(arr1) = getArrayDimension(arr2) Then
        Dim arr() As Variant
        Dim dimension As Long: Let dimension = getArrayDimension(arr1) - 1 ' Just code for 2D only ***
        Dim arr1Length As Long: Let arr1Length = getArrayLength(arr1)
        Dim arr2Length As Long: Let arr2Length = getArrayLength(arr2)
        ReDim arr(LBound(arr1) To arr1Length + arr2Length, dimension To dimension)
        Dim i, j As Long
        For i = LBound(arr) To UBound(arr1)
            arr(i, dimension) = arr1(i, dimension)
        Next i
        Let i = 0
        For j = UBound(arr1) + 1 To UBound(arr)
            i = i + 1
            arr(j, dimension) = arr2(i, dimension)
        Next j
        Let mergeTwoArrays = arr
    End If
GoTo ExecuteProcedure
ErrorHandle:
    Call tackleErrors
ExecuteProcedure:
End Property
'
'Public Property Get setArraySameValue(ByRef arr As Variant, ByRef value As Variant) As Variant
'    Dim dimension As Long: Let dimension = getArrayDimension(arr)
'End Property

' TO-DO: Incomplete
Public Property Get mergeMulArrays(ParamArray Arrays() As Variant) As Variant
    Const dimension = 1 ' for Range to Cells 2D
    Dim i As Long, j As Long, cnter As Long, UBoundArr As Long, OldUBoundArray As Long
    Dim arr() As Variant
    For j = LBound(Arrays) To UBound(Arrays)
        Let UBoundArr = UBoundArr + UBound(Arrays(j), dimension)
    Next j
    ReDim arr(1 To UBoundArr, dimension To dimension) ' magic num 1 is for Range to Cells also
    For j = LBound(Arrays) To UBound(Arrays)
        For i = LBound(Arrays(j)) To UBound(Arrays(j))
            Let arr(i + OldUBoundArray, dimension) = Arrays(j)(i, dimension)
        Next i
        Let OldUBoundArray = OldUBoundArray + UBound(Arrays(j), dimension)
    Next j
    Let mergeMulArrays = arr
End Property
'CLIPBOARD
Public Property Get getClipboard() As Variant
    Dim clipboardObj As Object
     Set clipboardObj = createHTMLFile()
    If Not clipboardObj Is Nothing Then
        getClipboard = clipboardObj.ParentWindow.ClipboardData.GetData("text")
    Else
        Set clipboardObj = createClipboard()
        clipboardObj.GetFromClipboard
        getClipboard = clipboardObj.getText()
    End If
    Set clipboardObj = Nothing
End Property

'MUTATORS

'STATUS BAR
Public Property Let setStatusBar(ByVal content As Variant)
    If Not hasWorkPlace(True, EXCEL_TYPE) Then Exit Property
    Let app.DisplayStatusBar = False
    If typeName(content) = "String" Then
        Let app.statusBar = content
    Else
        Let app.statusBar = vbNullString
    End If
    Let app.DisplayStatusBar = True
End Property
'RIBBON
Public Property Let isHideRibbon(ByRef isHide As Boolean)
    If isHide Then
        app.ExecuteExcel4Macro "show.toolbar(""Ribbon"",False)"
    Else
        app.ExecuteExcel4Macro "show.toolbar(""Ribbon"",True)"
    End If
End Property
'CLIPBOARD
' NOTE: use text as Variant for 64-bit VBA support
Public Property Let setClipboard(ByRef text As Variant)
    Dim clipboardObj As Object
    Set clipboardObj = createHTMLFile()
    If Not clipboardObj Is Nothing Then
        clipboardObj.ParentWindow.ClipboardData.SetData "text", text
    ' NOTE: below some case can cause encoding bug
    Else
        Set clipboardObj = createClipboard()
        clipboardObj.SetText (text)
        clipboardObj.PutInClipboard
    End If
    Set clipboardObj = Nothing
End Property

'METHODS

'TIMER
Public Sub restartTimer()
    Let systemTimer = Timer ' Restart timer
End Sub

'SHORTCUT
Public Sub createShortcut(ByVal key As String, ByVal functionName As String, Optional ByVal isActive As Boolean = True, Optional ByVal isRemove As Boolean = False)
    ' Add Shortcut
    If isActive Then Application.OnKey key, functionName
    ' Delete Shortcut
    If Not isActive Then Application.OnKey key
    ' Delete Shortcut
    If Not isRemove Then Application.OnKey key, vbNullString
End Sub

'SPEED
Public Sub speedOn( _
    Optional ByVal screenStt As Boolean = False, _
    Optional ByVal eventsStt As Boolean = False, _
    Optional ByVal linksStt As Boolean = False, _
    Optional ByVal alertsStt As Boolean = False)
    With Application
        If .ScreenUpdating <> screenStt Then .ScreenUpdating = screenStt
        If .EnableEvents <> eventsStt Then .EnableEvents = eventsStt
'        If .AskToUpdateLinks <> linksStt Then .AskToUpdateLinks = linksStt
        If .DisplayAlerts <> alertsStt Then .DisplayAlerts = alertsStt
        If Not ActiveWindow Is Nothing Then
            If .Calculation <> xlAutomatic Then .Calculation = xlAutomatic
        End If
    End With
End Sub

'SPEED
Public Sub speedOff( _
    Optional ByVal screenStt As Boolean = True, _
    Optional ByVal eventsStt As Boolean = True, _
    Optional ByVal linksStt As Boolean = True, _
    Optional ByVal alertsStt As Boolean = True)
    With Application
        If .ScreenUpdating <> screenStt Then .ScreenUpdating = screenStt
        If .EnableEvents <> eventsStt Then .EnableEvents = eventsStt
        If .AskToUpdateLinks <> linksStt Then .AskToUpdateLinks = linksStt
        If .DisplayAlerts <> alertsStt Then .DisplayAlerts = alertsStt
        If Not ActiveWindow Is Nothing Then
            If .Calculation <> xlAutomatic Then .Calculation = xlAutomatic
        End If
    End With
End Sub

'FORMAT
Public Sub storeSheetFormat(fromSheet As Worksheet, Optional toSheetName As String = "formatStored")
    'Create format sheet if not exist
    If toSheetName <> "formatStored" Then
        If Not hasSheet(toSheetName) Then ThisWorkbook.Worksheets.add.name = toSheetName
    End If
    'Copy all from object sheet
    fromSheet.Cells.copy
    'Paste format to stored sheet
    ThisWorkbook.Worksheets(toSheetName).Cells(1, 1).PasteSpecial Paste:=xlPasteFormats
End Sub

'FORMAT
Public Sub pasteSheetFormat(toSheetName As Worksheet, Optional fromSheet As String = "formatStored")
    'Create format sheet if not exist
    If fromSheet <> "formatStored" Then
        If Not hasSheet(fromSheet) Then ThisWorkbook.Worksheets.add.name = toSheetName
    End If
'    Copy all from object sheet
    ThisWorkbook.Worksheets(fromSheet).Cells.copy
    'Paste format to stored sheet
    toSheetName.Cells(1, 1).PasteSpecial Paste:=xlPasteFormats
End Sub

'Error Handler
Public Sub tackleErrors()
    Select Case Err.Number
        Case 0
        'Can't enable Addin
        Case 1004
        'Can not import VBA file
        Case 50057
            Let userResponse = MsgBox( _
                Err.Description & _
                vbNewLine & _
                info.getPrompt & _
                "Can not import this file!")
        'FIle import is not VBA file
        Case 50021
            Let userResponse = MsgBox( _
                Err.Description & _
                vbNewLine & _
                info.getPrompt & _
                "This file is not VBA file!")
        'VBA file have password
        Case 50289
            Let userResponse = MsgBox( _
                Prompt:=Err.Description & _
                    vbNewLine & _
                    info.getPrompt & _
                    "Can not access this VBA file because of been protected by a password!", _
                Buttons:=vbCritical + vbOKOnly, _
                Title:=info.getAuthor)
        'Un-handled Error
        Case Else
            Call errorDisplay
    End Select
    Call speedOff
    On Error GoTo 0
End Sub

'Error Handler
Private Sub errorDisplay()
    Dim errorMessage As String
    Let errorMessage = _
        "Error # " & Str(Err.Number) & _
        " was generated by " & Err.Source & _
        vbNewLine & "Error Line: " & Erl & _
        vbNewLine & Err.Description
    Let userResponse = MsgBox( _
        Prompt:=errorMessage, _
        Title:="Error", _
        HelpFile:=Err.HelpFile, _
        Context:=Err.HelpContext)
End Sub

'VBE - DEBUGGER
Public Sub debugAssert(condition As Boolean)
#If debugger Then
    Debug.Assert condition
#End If
End Sub

Public Sub clearImmediateWindow()
    Debug.Print String(65535, vbCr)
End Sub


