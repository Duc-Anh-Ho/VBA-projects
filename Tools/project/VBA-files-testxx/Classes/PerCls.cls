VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PerCls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' UPDATE SYSTEM TOOLS
' Author: DANH
' Version: 1.0.2
' Update: 2022/09/25

Option Explicit

'Private excelPath As String
'Private excelfile As Workbook
Private Const MAIN_SHEET_NAME = "MainSheet"
Private systemTimer As Single

'Constructor
Private Sub Class_Initialize()
        Let systemTimer = Timer ' start timer
        Debug.Print "--PerCLs Initilized--"
End Sub

'Destructor
Private Sub Class_Terminate()
        Debug.Print "--PerCLs Terminate-- " & getTimerMilestone()
End Sub

'New update
'Public Sub startTimer( )
Public Sub restartTimer()
        On Error Resume Next
        Let systemTimer = Timer ' start timer
        On Error GoTo 0
End Sub

Public Property Get getTimerMilestone() As String
        On Error Resume Next
        Let getTimerMilestone = Timer - systemTimer & " secs"
        On Error GoTo 0
End Property

'Assessor

Public Property Get getExcelPath(Optional ByVal fileFilter As String) As String
        On Error Resume Next
        With Application
                If fileFilter = vbNullString Then fileFilter = "Excel Files (*.xls*), *xls*"
                Let getExcelPath = .GetOpenFilename( _
                        Title:="Select file", _
                        fileFilter:=fileFilter)
       End With
       On Error GoTo 0
End Property

Public Property Get getExcelFile(ByVal Path As Variant) As Workbook
        On Error Resume Next
        With Application
                If Path <> False Then Set getExcelFile = .Workbooks.Open( _
                                fileName:=Path, _
                                ReadOnly:=True)
        End With
        On Error GoTo 0
End Property

Public Property Get getMainSheetName() As String
        Let getMainSheetName = MAIN_SHEET_NAME
End Property

Public Property Get getLastRow(ByVal ws As Worksheet, Optional ByVal atColumn As Long) As Long
        On Error Resume Next
        With ws
                If atColumn = False Then
                        Let getLastRow = .Cells.Find( _
                                "*", _
                                SearchOrder:=xlByRows, _
                                SearchDirection:=xlPrevious).Row
                Else
                        If .Cells(.Rows.count, atColumn).End(xlUp).Value = False Then
                                getLastRow = 0
                        Else
                                Let getLastRow = .Cells(.Rows.count, atColumn).End(xlUp).Row
                        End If
                End If
        End With
        On Error GoTo 0
End Property

Public Property Get getLastColumn(ByVal ws As Worksheet, Optional ByVal atRow As Long) As Long
        On Error Resume Next
        With ws
                If atRow = False Then
                        Let getLastColumn = ws.Cells.Find( _
                                "*", _
                                SearchOrder:=xlByColumns, _
                                SearchDirection:=xlPrevious).Column
                Else
                        If .Cells(atRow, .Columns.count).End(xlToLeft).Value = False Then
                                getLastColumn = 0
                        Else
                                Let getLastColumn = .Cells(atRow, .Columns.count).End(xlToLeft).Column
                        End If
                End If
        End With
        On Error GoTo 0
End Property

Public Property Get getArrayLength(ByRef arr As Variant, Optional ByVal dimension As Long = -1) As Long
       On Error Resume Next
                If IsEmpty(arr) Then
                        Let getArrayLength = 0
                Else
                        If dimension = -1 Then
                                Let getArrayLength = UBound(arr) - LBound(arr) + 1
                        Else
                                Let getArrayLength = UBound(arr, dimension) - LBound(arr, dimension) + 1
                        End If
                End If
       On Error GoTo 0
End Property

Public Property Get getArrayDimension(ByRef arr As Variant) As Long
      On Error GoTo Err
      Dim i As Long
      Dim tmp As Long
      i = 0
      Do While True
            i = i + 1
            tmp = UBound(arr, i)
      Loop
Err:
    getArrayDimension = i - 1
End Property

Public Property Get mergeTwoArrays(ByRef arr1 As Variant, ByRef arr2 As Variant) As Variant
        On Error Resume Next
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
        On Error GoTo 0
End Property
'
'Public Property Get setArraySameValue(ByRef arr As Variant, ByRef value As Variant) As Variant
'    Dim dimension As Long: Let dimension = getArrayDimension(arr)
'End Property

' Incomplete
Public Property Get mergeMulArrays(ParamArray Arrays() As Variant) As Variant
    Const dimension = 1 ' for Range to Cells 2D
    Dim i As Long, j As Long, cnter As Long, UBoundArr As Long, OldUBoundArray As Long
    Dim arr() As Variant
    For j = LBound(Arrays) To UBound(Arrays)
        Let UBoundArr = UBoundArr + UBound(Arrays(j), dimension)
    Next j
    ReDim arr(1 To UBoundArr, dimension To dimension) ' magic 1 is for Range to Cells also
    For j = LBound(Arrays) To UBound(Arrays)
        For i = LBound(Arrays(j)) To UBound(Arrays(j))
            Let arr(i + OldUBoundArray, dimension) = Arrays(j)(i, dimension)
        Next i
        Let OldUBoundArray = OldUBoundArray + UBound(Arrays(j), dimension)
    Next j
    Let mergeMulArrays = arr
End Property

' Speed up project: turn off nnecessary items
Public Sub speedOn()
        With Application
                .ScreenUpdating = False
                .EnableEvents = False
                .AskToUpdateLinks = False
                .DisplayAlerts = False
                .Calculation = xlAutomatic
        End With
        ActiveWorkbook.Date1904 = False
        ActiveWindow.View = xlNormalView
        ' Debug.Print "--Speed ON--"
End Sub

' Turn on items
Public Sub speedOff()
        With Application
                .ScreenUpdating = True
                .EnableEvents = True
                .AskToUpdateLinks = True
                .DisplayAlerts = True
                .Calculation = xlAutomatic
                .StatusBar = False
        End With
        ActiveWorkbook.Date1904 = False
        ' Debug.Print "--Speed OFF--"
End Sub

Public Sub errorDisplay()
        Dim errorMessage As String: Let errorMessage = _
                "Error # " & Str(Err.Number) & _
                " was generated by " & Err.Source & _
                Chr(13) & "Error Line: " & Erl & _
                Chr(13) & Err.Description
        MsgBox errorMessage, , "Error", Err.HelpFile, Err.HelpContext
End Sub

'ELSA
Public Sub addShortcut(ByVal Key As String, ByVal functionName As String)
        Application.OnKey Key, functionName
End Sub

Public Sub removeShortcut(ByVal Key)
        Application.OnKey Key, vbNullString
End Sub

Public Sub clearImmediateWindow()
        Debug.Print String(65535, vbCr)
End Sub

Public Sub checkName(ByVal name As String, ByVal wb As Workbook, ByVal projectName As String)
        On Error Resume Next
        With wb
                .VBProject.name = projectName
                If .CodeName <> name Then
                        .VBProject.VBComponents(.CodeName).name = name
                End If
                If .Worksheets(1).name <> MAIN_SHEET_NAME Then
                       .Worksheets(1).name = MAIN_SHEET_NAME
                End If
        End With
        On Error GoTo 0
End Sub






