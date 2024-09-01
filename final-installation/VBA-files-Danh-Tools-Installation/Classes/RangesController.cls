VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RangesController"
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
Private targetRange As Range ' Object
Private userResponse As VbMsgBoxResult
Private highlightColor As Long
Private isBold As Boolean
Private blurRate As Byte
Private addSize As Byte
Private WithEvents appEvent As Application
Attribute appEvent.VB_VarHelpID = -1
Private WithEvents wbEvent As Workbook
Attribute wbEvent.VB_VarHelpID = -1
Private WithEvents wsEvent As Worksheet
Attribute wsEvent.VB_VarHelpID = -1
'Private Const INPUT_BOX_RANGE_TYPE As Byte = 8
Private Const START_CHAR As Byte = 1
Private Const SHEET_TYPE As String = "Worksheet"
Private Const WS_TYPE_DETAIL_SHEET As String = "xlWorksheet"
Private Const MULTI_REPLACE_TMP_SHEET_NAME As String = "mult-replace-tmp"

'Initialize Variables
Private Function hasVariables() As Boolean
On Error GoTo ErrorHandle
    Set info = New InfoConstants
    Set system = New SystemUpdate
    Set fileSystem = system.createFileSystem
    If Not system.hasWorkPlace( _
        hasMsg:=True, _
        workPlaceType:=WS_TYPE_DETAIL_SHEET _
    ) Then
        Let hasVariables = False
        GoTo ExecuteProcedure
    End If
    Set appEvent = system.app
    Set wbEvent = system.wb
    Set wsEvent = system.ws
    Let hasVariables = True
    GoTo ExecuteProcedure
ErrorHandle:
    Call system.tackleErrors
ExecuteProcedure:
End Function

'Constructor
Private Sub Class_Initialize()
'
End Sub
'Destructor
Private Sub Class_Terminate()
    Set appEvent = Nothing
    Set wbEvent = Nothing
    Set wsEvent = Nothing
End Sub

'ASSESSORS
'0- Check targetRange
Private Property Get isRange(selectedRange As Variant, Optional hasMessage As Boolean = True) As Boolean
    If typeName(selectedRange) = "Range" Then
'    If TypeName(Selection) = "ChartArea" Then
        Set targetRange = selectedRange
        Let isRange = True
    Else
        If hasMessage Then
            MsgBox _
                Prompt:=info.getPrompt & _
                    "You you have to select range!", _
                Buttons:=vbOKOnly + vbExclamation, _
                Title:=info.getAuthor
         End If
         Let isRange = False
    End If
End Property
'1-0 Inver Color
Private Property Get getInvertColor(ByRef RGBColor As Long) As Long
    Dim redColor As Integer
    Dim greenColor As Integer
    Dim blueColor As Integer
     'Algorithm color
    Let redColor = 255 - RGBColor Mod 256
    Let greenColor = 255 - RGBColor \ 256 Mod 256
    Let blueColor = 255 - RGBColor \ 65536 Mod 256
    Let getInvertColor = RGB(redColor, greenColor, blueColor)
End Property
'2-0 Bold First Line 'Check newline special character
Private Property Get getNewLine() As String
    If system.app.OperatingSystem Like "Windows*" Then
        Let getNewLine = Chr(10)
    Else
        Let getNewLine = vbNewLine
    End If
End Property
'3-0 Get Highlight Color
Public Property Get getHighlightColor() As Long
    Let getHighlightColor = highlightColor
End Property
'3-2 Get Bold Status
Public Property Get checkBold() As Boolean
    Let checkBold = isBold
End Property
'3-4 Get Blur Rate
Public Property Get getBlurRate() As Integer
    If blurRate >= 100 Then blurRate = 100
    Let getBlurRate = blurRate
End Property
'3-5 Get Add Size
Public Property Get getAddSize() As Integer
    If addSize >= 10 Then addSize = 10
    Let getAddSize = addSize
End Property
'MUTATORS
'3-1 Set up HIghlight Color
Public Property Let letHighlightColor(ByRef newHighlightColor As Long)
    Let highlightColor = newHighlightColor
End Property
'3-3 Set Up Bold Status
Public Property Let letBold(ByRef newBoldStt As Boolean)
    Let isBold = newBoldStt
End Property
'3-5 Set Up Blur Rate
Public Property Let letBlurRate(ByRef newBlurRate As Byte)
    Let blurRate = newBlurRate
End Property
'3-7 Set up Add Size
Public Property Let letAddSize(ByRef newAddSize As Byte)
    Let addSize = newAddSize
End Property
'METHODS
'2-1 Trigger Bold First Line
Private Sub triggerBoldFirstLine(ByRef cell As Range)
    Let cell.Characters( _
            Start:=START_CHAR, _
            Length:=InStr( _
                cell.value, _
                getNewLine)) _
        .Font.Bold = Not cell.Characters( _
            Start:=START_CHAR, _
            Length:=InStr( _
                cell.value, _
                getNewLine)) _
        .Font.Bold
End Sub
'3-0 Highlight Range
Private Sub HighlightRange(ByRef Target As Range)
    With Target
        .EntireRow.Interior.Color = getHighlightColor
        .EntireColumn.Interior.Color = getHighlightColor
        .Interior.TintAndShade = getBlurRate / 100
        .Font.size = .Font.size + getAddSize
        .Font.Bold = checkBold
        .EntireRow.AutoFit
        .EntireColumn.AutoFit
        .Select
    End With
End Sub
'4-0 Toggle Page Brake
Private Sub togglePageBreaks(ByRef isDisplay As Boolean, Optional isApplyAll As Boolean = False)
    Dim wsItem As Worksheet
    If isApplyAll Then
        For Each wsItem In system.wb.Worksheets
            Let wsItem.DisplayPageBreaks = isDisplay
        Next wsItem
    Else
        Let system.ws.DisplayPageBreaks = isDisplay
    End If
End Sub
'MAIN
'1. Inver Color
Public Sub invertColor()
On Error GoTo ErrorHandle
    If Not hasVariables Then GoTo ExecuteProcedure
    If Not isRange(Selection) Then GoTo ExecuteProcedure
    Call system.speedOn
    Dim cell As Object
    'Loop through range areas
    For Each cell In targetRange
        Let cell.Interior.Color = getInvertColor(cell.Interior.Color)
        Let cell.Font.Color = getInvertColor(cell.Font.Color)
        Let cell.Borders.Color = getInvertColor(cell.Borders.Color)
    Next cell
GoTo ExecuteProcedure
ErrorHandle:
    Call system.tackleErrors
ExecuteProcedure:
    Call system.speedOff
End Sub
'2. Bold First Line
Public Sub boldFirstLine()
On Error GoTo ErrorHandle
    If Not hasVariables Then GoTo ExecuteProcedure
    If Not isRange(Selection) Then GoTo ExecuteProcedure
    Call system.speedOn
    Dim cell As Range
    'Loop through ranges
    For Each cell In targetRange
        'Trigger Bold First Line
        Call triggerBoldFirstLine(cell)
    Next cell
    GoTo ExecuteProcedure
ErrorHandle:
    Call system.tackleErrors
ExecuteProcedure:
    Call system.speedOff
End Sub
'3.1 Store Format
Public Sub storeHighlightFormat(Optional ByVal onSheet As Worksheet)
On Error GoTo ErrorHandle
    If Not hasVariables Then GoTo ExecuteProcedure
    If onSheet Is Nothing Then Set onSheet = system.ws
    Call system.speedOn
    Call system.storeSheetFormat( _
        fromSheet:=onSheet, _
        toSheetName:="formatStored")
GoTo ExecuteProcedure
ErrorHandle:
    Call system.tackleErrors
ExecuteProcedure:
    Call system.speedOff
End Sub
'3.2 Paste Format
Public Sub pasteHighlightFormat(Optional ByVal onSheet As Worksheet)
On Error GoTo ErrorHandle
    If Not hasVariables Then GoTo ExecuteProcedure
    If onSheet Is Nothing Then Set onSheet = system.ws
    Call system.speedOn
    '  Change sheets address because targetRange now is in revious sheet *
    If targetRange Is Nothing Then
        If isRange(Selection, False) Then
            'Do Nothing
        End If
    End If
    Call system.pasteSheetFormat( _
        toSheetName:=onSheet, _
        fromSheet:="formatStored")
     If targetRange Is Nothing Then GoTo ExecuteProcedure
    Call system.ws.Range(targetRange.Address).Select
ErrorHandle:
    Call system.tackleErrors
ExecuteProcedure:
    Call system.speedOff
End Sub
'3.3 Highlight
Public Sub highlight(ByVal Target As Object)
On Error GoTo ErrorHandle
    If Not hasVariables Then GoTo ExecuteProcedure
    If Not isRange(Target, False) Then GoTo ExecuteProcedure
    Call system.speedOn
    Call HighlightRange(targetRange)
GoTo ExecuteProcedure
ErrorHandle:
    Call system.tackleErrors
ExecuteProcedure:
    Call system.speedOff
    Set system = Nothing 'Clear cache
End Sub
'4 Highlight
Public Sub displayPageBreak(ByRef isDisplay As Boolean, Optional isApplyAll As Boolean = False)
On Error GoTo ErrorHandle
    If Not hasVariables Then GoTo ExecuteProcedure
    Call system.speedOn
    Call togglePageBreaks(isDisplay, isApplyAll)
GoTo ExecuteProcedure
ErrorHandle:
    Call system.tackleErrors
ExecuteProcedure:
    Call system.speedOff
End Sub
'5 Replace
Public Function multipleReplace( _
    ByRef findArea As Range _
    , ByRef replaceArea As Range _
    , ByVal withinIndex As Byte _
    , Optional ByVal selectedArea As Range _
    , Optional ByVal isMatchCase As Boolean _
    , Optional ByVal isMatchByte As Boolean _
    , Optional ByVal isMatchContent As Boolean _
    , Optional ByVal searchOrderCd As Byte _
    , Optional ByVal isOrderByLength As Boolean _
) As Boolean
On Error GoTo ErrorHandle
    If Not hasVariables Then GoTo ExecuteProcedure
    Call system.speedOn
    Dim i As Long
    Dim lookAtCd As Byte: Let lookAtCd = IIf(isMatchContent, xlWhole, xlPart)
    Let searchOrderCd = IIf(searchOrderCd = 0, xlByRows, xlByColumns)
    ' TODO optional below
    Dim isSearchFormat As Boolean: Let isSearchFormat = False
    Dim isReplaceFormat As Boolean: Let isReplaceFormat = False
    Dim within As Object ' 0: Selection, 1: ActiveSheet, 2: ActiveWorkbook,
    
    Select Case withinIndex
        ' Selection
        Case 0: Set within = selectedArea
        ' ActiveSheet
        Case 1: Set within = system.ws.UsedRange
        ' ActiveWorkbook
        Case 2: Set within = Cells
    End Select
    
    If isOrderByLength Then
        Dim tmpWs As Worksheet
        Dim size As Long: Let size = findArea.Rows.Count
        Dim keyArea As Range
        Dim sortArea As Range
        Set tmpWs = ThisWorkbook.Sheets(MULTI_REPLACE_TMP_SHEET_NAME)
        
        With tmpWs
            Set keyArea = .Range("C1:C" & size)
            Set sortArea = .Range("A1:C" & size)
            Let .Range("A1").Resize(size, 1).value = findArea.value
            Let .Range("B1").Resize(size, 1).value = replaceArea.value
            Let .Range("C1").Resize(size, 1).Formula = "=LEN(A1)"
            With .Sort
                .SortFields.Clear
                .SortFields.add _
                    key:=keyArea, _
                    SortOn:=xlSortOnValues, _
                    Order:=xlDescending, _
                    DataOption:=xlSortNormal
                .SetRange sortArea
                .Header = xlNo
                .Apply
            End With
            Set findArea = .Range("A1:A" & size)
            Set replaceArea = .Range("B1:B" & size)
         End With
    End If
    
    For i = 1 To findArea.Count
        If Len(findArea.Cells(i)) > 0 Then
            within.Replace _
                What:=findArea.Cells(i).value _
                , Replacement:=replaceArea.Cells(i).value _
                , LookAt:=lookAtCd _
                , SearchOrder:=searchOrderCd _
                , MatchCase:=isMatchCase _
                , MatchByte:=isMatchByte _
                , SearchFormat:=isSearchFormat _
                , ReplaceFormat:=isReplaceFormat
        End If
    Next i
    If isOrderByLength Then
        Set tmpWs = ThisWorkbook.Sheets(MULTI_REPLACE_TMP_SHEET_NAME)
        tmpWs.UsedRange.Clear
    End If
        Let userResponse = MsgBox( _
            Prompt:=info.getPrompt & "Multiple replacment successfull!", _
            Buttons:=vbOKOnly + vbInformation, _
            Title:=info.getAuthor)
GoTo ExecuteProcedure
ErrorHandle:
    Call system.tackleErrors
ExecuteProcedure:
    Call system.speedOff
    Call MultipleReplaceForm.closeForm
End Function
'EVENTS

Private Sub appEvent_SheetDeactivate(ByVal sheetObject As Object)
    If typeName(sheetObject) = SHEET_TYPE Then
        Call pasteHighlightFormat(sheetObject)
    End If
End Sub

Private Sub appEvent_SheetActivate(ByVal sheetObject As Object)
   If typeName(sheetObject) = SHEET_TYPE Then
        Call storeHighlightFormat(sheetObject)
        Call highlight(Selection)
   End If
End Sub

Private Sub appEvent_SheetSelectionChange(ByVal sheetObject As Object, ByVal Target As Range)
    If typeName(sheetObject) = SHEET_TYPE Then
        Call pasteHighlightFormat(sheetObject)
        Call highlight(Target)
    End If
End Sub

