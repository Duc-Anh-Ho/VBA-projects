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
Private userResponse As VbMsgBoxResult
Private info As Object
Private systemTimer As Single
'Global
Public WithEvents app As Application
Attribute app.VB_VarHelpID = -1
Public WithEvents Wb As Workbook
Attribute Wb.VB_VarHelpID = -1
Public WithEvents ws As Worksheet
Attribute ws.VB_VarHelpID = -1
'Initialize Variables
Private Function hasVariables() As Boolean
On Error GoTo ErrorHandle
    Set info = New MyInfo
    Let systemTimer = Timer ' start timer
    hasVariables = True
    GoTo ExecuteProcedure
ErrorHandle:
    Call tackleErrors
ExecuteProcedure:
End Function

'Constructor
Private Sub Class_Initialize()
On Error GoTo ErrorHandle
    If Not hasVariables Then GoTo ExecuteProcedure
    '
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
Public Property Get hasWorkplace() As Boolean
On Error GoTo ErrorHandle
    'Check EXCEL app
    If Application.Name <> "Microsoft Excel" Then
        MsgBox _
            Prompt:=info.getPrompt & _
                "Sorry, this tool currently is just working with MS-EXCEL", _
            Buttons:=vbOKOnly + vbExclamation, _
            Title:=info.getAuthor
            hasWorkplace = False
            GoTo ExecuteProcedure
    Else
        Set app = Application
    End If
    'Check ActiveWorkbook exist
    If app.ActiveWorkbook Is Nothing Then
        MsgBox _
            Prompt:=info.getPrompt & _
                "Can not access file, please re-open!", _
            Buttons:=vbOKOnly + vbExclamation, _
            Title:=info.getAuthor
        hasWorkplace = False
        GoTo ExecuteProcedure
    Else
        Set Wb = app.ActiveWorkbook
    End If
    'Check ActiveSheet exist
    If app.ActiveSheet Is Nothing Then
        MsgBox _
            Prompt:=info.getPrompt & _
                "Please, Select a target first!", _
            Buttons:=vbOKOnly + vbExclamation, _
            Title:=info.getAuthor
        hasWorkplace = False
        GoTo ExecuteProcedure
    Else
        Set ws = Wb.ActiveSheet
    End If
    'All right
    hasWorkplace = True
    GoTo ExecuteProcedure
ErrorHandle:
    Call tackleErrors
ExecuteProcedure:
End Property
'FILE
Public Property Get createFileSystem() As Object
    Set createFileSystem = CreateObject("Scripting.FileSystemObject")
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
                Filename:=path, _
                ReadOnly:=True)
    End With
GoTo ExecuteProcedure
ErrorHandle:
    Call tackleErrors
ExecuteProcedure:
End Property
'POSITION
Public Property Get getLastRow(ByVal ws As Worksheet, Optional ByVal atColumn As Long) As Long
On Error GoTo ErrorHandle
    With ws
        If atColumn = False Then
                Let getLastRow = .Cells.Find( _
                        "*", _
                        SearchOrder:=xlByRows, _
                        SearchDirection:=xlPrevious).Row
        Else
                If .Cells(.Rows.Count, atColumn).End(xlUp).Value = False Then
                        getLastRow = 0
                Else
                        Let getLastRow = .Cells(.Rows.Count, atColumn).End(xlUp).Row
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
            If .Cells(atRow, .Columns.Count).End(xlToLeft).Value = False Then
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
    i = 0
    Do While True
        i = i + 1
        tmp = UBound(arr, i)
    Loop
OverDimensionErr:
    getArrayDimension = i - 1
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
'VBE
Public Sub clearImmediateWindow()
    Debug.Print String(65535, vbCr)
End Sub
'SPEED
Public Sub speedOn( _
    Optional ByVal screenStt As Boolean = False, _
    Optional ByVal eventsStt As Boolean = False, _
    Optional ByVal linksStt As Boolean = False, _
    Optional ByVal alertsStt As Boolean = False)
    With Application
        .ScreenUpdating = screenStt
        .EnableEvents = eventsStt
        .AskToUpdateLinks = linksStt
        .DisplayAlerts = alertsStt
         If Not ActiveWindow Is Nothing Then .Calculation = xlManual
    End With
End Sub
'SPEED
Public Sub speedOff( _
    Optional ByVal screenStt As Boolean = True, _
    Optional ByVal eventsStt As Boolean = True, _
    Optional ByVal linksStt As Boolean = True, _
    Optional ByVal alertsStt As Boolean = True)
    With Application
        .ScreenUpdating = screenStt
        .EnableEvents = eventsStt
        .AskToUpdateLinks = linksStt
        .DisplayAlerts = alertsStt
        If Not ActiveWindow Is Nothing Then .Calculation = xlAutomatic
    End With
End Sub
'Error Handler
Public Sub tackleErrors()
    Select Case Err.Number
        Case 0
        'Can't enable Addin
        Case 1004
        'Can not import VBA file
        Case 50057
            MsgBox Err.Description & _
                vbNewLine & _
                info.getPrompt & _
                "Can not import this file!"
        'FIle import is not VBA file
        Case 50021
            MsgBox Err.Description & _
                vbNewLine & _
                info.getPrompt & _
                "This file is not VBA file!"
        'VBA file have password
        Case 50289
            MsgBox _
                Prompt:=Err.Description & _
                    vbNewLine & _
                    info.getPrompt & _
                    "Can not access this VBA file because of been protected by a password!", _
                Buttons:=vbCritical + vbOKOnly, _
                Title:=info.getAuthor
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
    MsgBox _
        Prompt:=errorMessage, _
        Title:="Error", _
        HelpFile:=Err.HelpFile, _
        Context:=Err.HelpContext
End Sub
'EVENTS
'Update actived workbook and actived sheet when changed
Private Sub app_SheetActivate(ByVal sh As Object)
    If hasWorkplace Then
    End If
End Sub
'Private Sub app_WorkbookActivate(ByVal Wb As Workbook)
'    Call hasWorkplace
'End Sub
