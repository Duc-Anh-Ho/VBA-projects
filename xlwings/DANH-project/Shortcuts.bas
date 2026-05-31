Attribute VB_Name = "Shortcuts"
' Check README.md for more information
Option Explicit
' NOTE: Must put Onkey procedure methods in a module for calling as full global scope
'METHODS
Private Sub copyName()
    Dim fileController As FilesController
    Set fileController = New FilesController
    Call fileController.copyFileName("name")
    Set fileController = Nothing
End Sub

Private Sub copyFullName()
    Dim fileController As FilesController
    Set fileController = New FilesController
    Call fileController.copyFileName("fullName")
    Set fileController = Nothing
End Sub

Private Sub copyShortName()
    Dim fileController As FilesController
    Set fileController = New FilesController
    Call fileController.copyFileName("shortName")
    Set fileController = Nothing
End Sub

Private Sub copyPath()
    Dim fileController As FilesController
    Set fileController = New FilesController
    Call fileController.copyFileName("path")
    Set fileController = Nothing
End Sub

Private Sub copyExtensionName()
    Dim fileController As FilesController
    Set fileController = New FilesController
    Call fileController.copyFileName("extension")
    Set fileController = Nothing
End Sub

Private Sub copyF()
    Dim formatC As FormatController
    Set formatC = New FormatController
    Call formatC.copyFormat
    Set formatC = Nothing
End Sub

Private Sub pasteF()
    Dim formatC As FormatController
    Set formatC = New FormatController
    Call formatC.pasteFormat
    Set formatC = Nothing
End Sub

Private Sub pasteV()
    Dim formatC As FormatController
    Set formatC = New FormatController
    Call formatC.pasteValue
    Set formatC = Nothing
End Sub

Private Sub sheetSelectN()
    Dim sheetC As SheetsController
    Set sheetC = New SheetsController
    Call sheetC.selectNext
    Set sheetC = Nothing
End Sub

Private Sub sheetSelectP()
    Dim sheetC As SheetsController
    Set sheetC = New SheetsController
    Call sheetC.selectPrevious
    Set sheetC = Nothing
End Sub

Private Sub sheetFocusRename()
    Dim sheetC As SheetsController
    Set sheetC = New SheetsController
    Call sheetC.focusRename
    Set sheetC = Nothing
End Sub

Private Sub shapeMoveAndSize()
    Dim formatC As FormatController
    Set formatC = New FormatController
    Call formatC.setPlacement(xlMoveAndSize)
    Set formatC = Nothing
End Sub

Private Sub shapeMove()
    Dim formatC As FormatController
    Set formatC = New FormatController
    Call formatC.setPlacement(xlMove)
    Set formatC = Nothing
End Sub

Private Sub shapeFree()
    Dim formatC As FormatController
    Set formatC = New FormatController
    Call formatC.setPlacement(xlFreeFloating)
    Set formatC = Nothing
End Sub

Private Sub clearContent()
    Dim formatC As FormatController
    Set formatC = New FormatController
    Call formatC.clearContent
    Set formatC = Nothing
End Sub

Private Sub clearFormat()
    Dim formatC As FormatController
    Set formatC = New FormatController
    Call formatC.clearFormat
    Set formatC = Nothing
End Sub

Private Sub clearAll()
    Dim formatC As FormatController
    Set formatC = New FormatController
    Call formatC.clearAll
    Set formatC = Nothing
End Sub

Private Sub toggleZenMode()
    Dim modeC As ModeController
    Set modeC = New ModeController
    Call modeC.toggleZenMode
    Set modeC = Nothing
End Sub

Private Sub toggleZoomMode()
    Dim modeC As ModeController
    Set modeC = New ModeController
    Call modeC.toggleZoomMode(False)
    Set modeC = Nothing
End Sub

Private Sub toggleZoomModeMax()
    Dim modeC As ModeController
    Set modeC = New ModeController
    Call modeC.toggleZoomMode(True)
    Set modeC = Nothing
End Sub

Private Sub zoomMode100()
    Dim modeC As ModeController
    Set modeC = New ModeController
    Call modeC.zoom100
    Set modeC = Nothing
End Sub


Private Sub openMultipleReplaceForm()
    Dim form As MultipleReplaceForm
    Set form = New MultipleReplaceForm
    Call form.Show(vbModal)  ' vbModeless or vbModal
    Set form = Nothing
End Sub

Private Sub openShortcutForm()
    ' NOTE: Reject Below cannot use Hover event (Maybe overload issue) need to research more)
    ' Dim form As KeyboardShortcutForm
    ' Set form = New KeyboardShortcutForm
    ' Call form.Show(vbModal)  ' vbModeless or vbModal
    ' Set form = Nothing
    Call KeyboardShortcutForm.Show(vbModal)
End Sub

' NEW

Private Sub setRedFont()
    If Not TypeOf Selection Is Excel.Range Then Exit Sub
    Let Selection.Font.COLOR = IIf(Selection.Font.COLOR = 255, vbBlack, 255) ' xlColorIndexAutomatic
End Sub

Private Sub setYellowBackground()
    If Not TypeOf Selection Is Excel.Range Then Exit Sub
    Let Selection.Interior.COLOR = IIf(Selection.Interior.COLOR = 65535, xlNone, 65535)
End Sub

Private Sub nameToContent(Optional ByRef target As Object = Nothing)
    Dim SKIP As String: Let SKIP = "{S}"
    Dim obj As Object
    Dim i As Integer
    Set target = IIf(target Is Nothing, Selection, target)
    Select Case True
        Case TypeOf target Is Excel.Shape: GoTo pickedIsShape
        Case TypeOf target Is Excel.Rectangle: GoTo pickedSingle
        Case TypeOf target Is Excel.Oval: GoTo pickedSingle
        Case TypeOf target Is Excel.GroupObject: GoTo pickedGroup
        Case TypeOf target Is Excel.DrawingObjects: GoTo pickedMultiple
        ' NO USE
        Case TypeOf target Is Excel.Picture: Exit Sub
        Case TypeOf target Is Excel.Range: Exit Sub
        Case TypeOf target Is Excel.line: Exit Sub
        Case TypeOf target Is Excel.ChartArea: Exit Sub
        Case TypeOf target Is Excel.PlotArea: Exit Sub
        ' TODO: Handle
        Case Else
            MsgBox "Unhandled Type Name: " & typeName(target): Exit Sub
    End Select
    
' TODO: Clean code create method
pickedIsShape:
    If left(target.name, 3) = SKIP Then Exit Sub
    If Not target.TextFrame2.HasText Then Exit Sub ' Skip empty init also
    Select Case target.Type
        Case msoAutoShape: Let target.TextFrame.Characters.text = target.name ' All ver can use
        ' Case msoAutoShape: Let target.TextFrame2.TextRange.text = target.name ' Excel 2007 Up only
        Case msoPicture: Exit Sub
        Case Else: MsgBox "Unhandled Type Code: " & target.Type: Exit Sub
    End Select
    Exit Sub
pickedSingle:
    Call nameToContent(target.ShapeRange(1))
    Exit Sub
pickedGroup:
    For Each obj In target.ShapeRange(1).GroupItems
        Call nameToContent(obj)
    Next obj
    Exit Sub
pickedMultiple:
    For Each obj In target
        Call nameToContent(obj)
    Next obj
    Exit Sub
End Sub

Private Sub listAllShapes( _
    Optional ByRef target As Object = Nothing _
    , Optional ByRef index As Long = 0 _
)
    If Not TypeOf Selection Is Excel.Range Then Exit Sub
    Dim textFormat As String: Let textFormat = "@"
    Dim defaultHeader As String: Let defaultHeader = "Selection Pane"
    Dim shp As Shape
    ' target can be Shapes or GroupShapes
    Set target = IIf(target Is Nothing, ActiveSheet.Shapes, target)
    If index = 0 Then
        Let Selection.Cells(1, 1).NumberFormat = "@"
        Let Selection.Cells(1, 1).value = defaultHeader
        Let index = index + 1
    End If
    For Each shp In target
        Let Selection.Cells(1, 1).offset(index, 0).NumberFormat = "@"
        Let Selection.Cells(1, 1).offset(index, 0).value = shp.name
        Let index = index + 1
        If shp.Type = msoGroup Then Call listAllShapes(shp.GroupItems, index)
    Next shp
End Sub

Private Sub renameAllShapes( _
    Optional ByRef target As Object = Nothing _
    , Optional ByRef index As Long = 0 _
)
    If Not TypeOf Selection Is Excel.Range Then Exit Sub
    If Selection.Cells(1, 1).value <> "Selection Pane" Then Exit Sub
    Dim shp As Shape
    Dim name As String
    ' target can be Shapes or GroupShapes
    Set target = IIf(target Is Nothing, ActiveSheet.Shapes, target)
    If index = 0 Then Let index = index + 1
    For Each shp In target
        Let name = Selection.Cells(1, 1).offset(index, 0).value
        If RTrim(name) = vbNullString Then
            Let shp.name = "index_" & index
        Else
            Let shp.name = Selection.Cells(1, 1).offset(index, 0).value
        End If
        Let index = index + 1
        If shp.Type = msoGroup Then Call renameAllShapes(shp.GroupItems, index)
    Next shp
End Sub

Private Sub showListSheets()
    With Application.CommandBars("Workbook tabs")
        If .controls(.controls.count).caption = "More Sheets..." Then
            .controls(.controls.count).Execute
        Else
            .ShowPopup
        End If
    End With
End Sub

Private Sub showListWorkbooks()
    Application.Dialogs(xlDialogActivate).Show
End Sub

Private Sub autoFill(Optional ByRef fillType As Byte = xlFillDefault)
    If Not TypeOf Selection Is Excel.Range Then Exit Sub
    If Application.CountA(Selection) = 0 Then Exit Sub
    Dim lastData As Range: Set lastData = Selection.Find( _
        What:="*", _
        After:=Selection.Cells(1), _
        LookIn:=xlFormulas, _
        LookAt:=xlPart, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlPrevious _
    )
    Dim area As Range: Set area = Selection.CurrentRegion
    Dim lastRow As Long: Let lastRow = IIf( _
        Selection(Selection.count).Row = lastData.Row, _
        area.Row + area.Rows.count - 1, _
        Selection(Selection.count).Row _
    )
    Dim target As Range: Set target = IIf( _
        Selection(Selection.count).Row = lastData.Row, _
        Selection, _
        lastData _
    )
    Call target.autoFill( _
        Destination:=Range(target, Cells(lastRow, target.Column)) _
        , Type:=fillType _
    )
End Sub

Private Sub autoFillSeries()
    Call autoFill(xlFillDefault)
End Sub

' UNGROUP/GROUP ALL SHAPES

Private Function isIntersect(ByVal sh1 As Shape, ByVal sh2 As Shape) As Boolean
    Let isIntersect = Not ( _
        sh1.left + sh1.width <= sh2.left Or _
        sh1.left >= sh2.left + sh2.width Or _
        sh1.top + sh1.height <= sh2.top Or _
        sh1.top >= sh2.top + sh2.height)
End Function

' Disjoint Set Union (Union-Find) Algorithm
Private Function findRoot(ByRef groupRoot() As Long, ByVal index As Long) As Long
    ' Find until meet root
    If groupRoot(index) <> index Then
        Let groupRoot(index) = findRoot(groupRoot, groupRoot(index))
    End If
    Let findRoot = groupRoot(index)
End Function

Private Function mergeRoot(ByRef groupRoot() As Long, ByVal x As Long, ByVal y As Long) As Long
    Dim rootX As Long, rootY As Long
    Let rootX = findRoot(groupRoot, x)
    Let rootY = findRoot(groupRoot, y)
    If rootX <> rootY Then Let groupRoot(rootY) = rootX
    Let mergeRoot = groupRoot(rootY)
End Function

Private Function getIntersectedArr(ByVal ws As Worksheet) As Variant
    Dim maxShape As Long: Let maxShape = ws.Shapes.count
    Dim i As Long, j As Long, groupRootArr() As Long
    ' Init roots group (self root)
    ReDim groupRootArr(1 To maxShape)
    For i = 1 To ws.Shapes.count
        Let groupRootArr(i) = i
    Next
    ' Merge Intersect into group (merge root)
    For i = 1 To maxShape - 1
        For j = i + 1 To maxShape
            If isIntersect(ws.Shapes(i), ws.Shapes(j)) Then
                Call mergeRoot(groupRootArr, i, j)
            End If
        Next j
    Next i
    ' Flatten all group pointers to final (normalize root)
    For i = 1 To maxShape
        Let groupRootArr(i) = findRoot(groupRootArr, i)
    Next i
    ' Assign result
    Let getIntersectedArr = groupRootArr
End Function

'mode: ID = 0, Index = 1
Private Function getGroupShapeColl(ByVal ws As Worksheet, Optional ByVal mode As Byte = 0) As Collection
    Dim maxShape As Long: Let maxShape = ws.Shapes.count
    Dim shapeGroupIdArr() As Long
    Dim groupListColl As Collection, rootListColl As Collection, currentListColl As Collection
    Dim i As Long, j As Long, groupId As Long
    ' Init
    Let shapeGroupIdArr = getIntersectedArr(ws)
    Set groupListColl = New Collection
    Set rootListColl = New Collection
    ' Group by root
    For i = 1 To maxShape
        Let groupId = shapeGroupIdArr(i)
        For j = 1 To rootListColl.count
            If rootListColl(j) = groupId Then
                Set currentListColl = groupListColl(j)
                Call currentListColl.Add(IIf(mode = 0, ws.Shapes(i).id, i))
                GoTo NextShape
            End If
        Next j
        Set currentListColl = New Collection
        Call currentListColl.Add(IIf(mode = 0, ws.Shapes(i).id, i))
        Call rootListColl.Add(groupId)
        Call groupListColl.Add(currentListColl)
NextShape:
    Next i
    Set getGroupShapeColl = groupListColl
End Function

' TODO: Delete test
Private Sub TESTgetGroupShape()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim g, id, shp, line$: Set g = getGroupShapeColl(ws)
    Dim i&: For i = 1 To g.count
        line = "group" & i & ": { "
        For Each id In g(i)
            For Each shp In ws.Shapes
                If shp.id = id Then line = line & "[Idx: " & shp.ZOrderPosition & ", Id: " & id & ", """ & shp.name & """], ": Exit For
            Next
        Next
        Debug.Print left(line, Len(line) - 2) & " }"
    Next
End Sub

Private Sub renameAllShapesByPrefix( _
    ByVal ws As Worksheet _
    , ByVal groupColl As Collection _
    , Optional ByVal prefix As String = "00000" _
)
    Dim idColl As Collection
    Dim sh As Shape
    Dim i As Long, j As Long, shapeId As Long
    For i = 1 To groupColl.count
        Set idColl = groupColl(i)
        For j = 1 To idColl.count
            Let shapeId = idColl(j)
            For Each sh In ws.Shapes
                If sh.id = shapeId Then
                    Let sh.name = Format(sh.id, prefix) & "_" & sh.name ' Add ID prefix
                    Exit For
                End If
            Next sh
        Next j
    Next i
End Sub

Private Sub groupIntersectedShapes(ByVal ws As Worksheet)
    Dim maxShape As Long: Let maxShape = ws.Shapes.count
    ' No Shape handle
    If maxShape = 0 Then Exit Sub
    ' Get total number of shapes
    Dim groupColl As Collection, idColl As Collection
    Dim sh As Shape
    Dim i As Long, j As Long, shapeId As Long
    Dim nameArr() As String
    ' Get all intersected shape by ID
    Set groupColl = getGroupShapeColl(ws, 0) 'mode: ID = 0, Index = 1
    Call renameAllShapesByPrefix(ws, groupColl, "00000")
    ' Loop each group and group shapes
    For i = 1 To groupColl.count
        Set idColl = groupColl(i)
        If idColl.count < 2 Then GoTo SkipGroup
        ReDim nameArr(1 To idColl.count)
        For j = 1 To idColl.count
            Let shapeId = idColl(j)
            For Each sh In ws.Shapes
                If sh.id = shapeId Then
                    Let nameArr(j) = sh.name
                    Exit For
                End If
            Next sh
        Next j
        Call ws.Shapes.Range(nameArr).group
SkipGroup:
    Next i
    ' TODO: Undo Rename
End Sub

Private Sub groupAllShapes()
    Dim wb As Workbook: Set wb = ActiveWorkbook
    Dim ws As Worksheet
    Dim i As Integer
    If wb Is Nothing Then Exit Sub
    If wb.Worksheets.count = 0 Then Exit Sub
    Let Application.ScreenUpdating = False
    For i = wb.Sheets.count To 1 Step -1
        Set ws = wb.Sheets(i)
        If ws.Type = xlWorksheet Then
            Call ws.Activate
            Call groupIntersectedShapes(ws)
        End If
    Next i
    Let Application.ScreenUpdating = True
End Sub
    
Private Sub formatFileEnd()
    Dim wb As Workbook: Set wb = ActiveWorkbook
    Dim ws As Worksheet
    Dim i As Integer
    If wb Is Nothing Then Exit Sub
    If wb.Worksheets.count = 0 Then Exit Sub
    If Not isBackupSaved Then Exit Sub
    Let Application.ScreenUpdating = False
    For i = wb.Sheets.count To 1 Step -1
        Set ws = wb.Sheets(i)
        If ws.Type = xlWorksheet Then
            Call ws.Activate
            Call ws.Range("A1").Select
            Call zoomMode100
            Call convertGroupToImage(ws)
        End If
    Next i
    Let Application.ScreenUpdating = True
End Sub

'TODO: Refactor
Private Sub saveBackup()
    If Not isBackupSaved Then Exit Sub
End Sub

Private Sub unGroup(group As Variant)
    Dim sh As Shape
    If TypeOf group Is Excel.Shapes Then
        For Each sh In group
            Call unGroup(sh)
        Next
    ElseIf TypeOf group Is Excel.Shape And group.Type = msoGroup Then
        For Each sh In group.unGroup
            Call unGroup(sh)
        Next
    End If
End Sub

Private Sub ungroupAllShapes()
    ' Closed all sheets or called from the other app handle
    If ActiveSheet Is Nothing Then Exit Sub
    Call unGroup(ActiveSheet.Shapes)
End Sub

Private Function isBackupSaved() As Boolean
    If ActiveWorkbook Is Nothing Then Let isBackupSaved = False: Exit Function
    Dim wb As Workbook: Set wb = ActiveWorkbook
    With wb
    If .path = "" Then
        If MsgBox( _
            "This workbook has not been saved." & vbCrLf & "Do you want to Save it?" _
            , vbYesNo + vbQuestion _
            , "Save Required" _
        ) = vbNo Then
            Let isBackupSaved = False
            Exit Function
        End If
        Call Application.Dialogs(xlDialogSaveAs).Show
        If wb.path = "" Then Let isBackupSaved = False: Exit Function ' User cancelled
    End If
    Dim fName As String, ext As String, timestamp As String
    Let fName = .name
    Let ext = Mid(fName, InStrRev(fName, ".") + 1)
    Let fName = left(fName, InStrRev(fName, ".") - 1)
    Let timestamp = Format(Now, "yyyymmdd_hhnnss")
    Call .SaveCopyAs(.path & "\" & fName & "_bk_" & timestamp & "." & ext)
    End With
    Let isBackupSaved = True
End Function

Private Sub convertGroupToImage(ByVal ws As Worksheet)
    Dim sh As Shape, newSh As Shape
    Dim leftTmp As Single, topTmp As Single
    ' Dim placement As Byte: Let placement = xlFreeFloating ' Don't Move And Size With Cell
    Dim placement As Byte: Let placement = xlMoveAndSize ' Move And Size With Cell
    Dim i As Long
    For i = ws.Shapes.count To 1 Step -1 ' Reverse to avoid index shift
        Set sh = ws.Shapes(i)
        Let sh.placement = placement
        If sh.Type = msoGroup Then
            ' Save tmp position
            Let leftTmp = sh.left
            Let topTmp = sh.top
            ' Convert to image by copy paste as pic
            Call sh.CopyPicture(Appearance:=xlScreen, Format:=xlPicture)
            Call ws.Paste
            ' Get newest sh and move to saved position
            Set newSh = ws.Shapes(ws.Shapes.count)
            Let newSh.left = leftTmp
            Let newSh.top = topTmp
            Let newSh.placement = placement
            ' Delete Origin (Need backup)
            Call sh.Delete
        End If
    Next i
End Sub

Private Sub formatFileStart()
    Dim wb As Workbook: Set wb = ActiveWorkbook
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If ws.Type = xlWorksheet Then
            With ws
                Let .Cells.RowHeight = 18  ' Standard height
                Let .Cells.ColumnWidth = 8   ' Standard width
            End With
        End If
    Next ws
End Sub

Private Sub captureShareX()
    Dim picC As PicturesController
    Set picC = New PicturesController
    Call picC.snipShareX("capture-last-region-workflow")
'    Call picC.snipShareX("D:\screenshots")
End Sub

Private Sub captureShareX_1()
    Dim picC As PicturesController
    Set picC = New PicturesController
    Call picC.snipShareX("capture-pre-config-region-workflow-1")
End Sub

Private Sub captureShareX_2()
    Dim picC As PicturesController
    Set picC = New PicturesController
    Call picC.snipShareX("capture-pre-config-region-workflow-2")
End Sub


Public Sub decreaseWidth(Optional ByRef offset As Byte = 2)
    Dim obj As Object: Set obj = Selection
    Select Case True
        Case TypeOf obj Is Excel.Shape: GoTo changeSize
        Case TypeOf obj Is Excel.Rectangle: GoTo changeSize
        Case TypeOf obj Is Excel.Oval: GoTo changeSize
        Case TypeOf obj Is Excel.GroupObject: GoTo changeSize
        Case TypeOf obj Is Excel.DrawingObjects: GoTo changeSize
        Case TypeOf obj Is Excel.Picture: GoTo changeSize
        Case TypeOf obj Is Excel.line: GoTo changeSize
        Case TypeOf obj Is Excel.ChartArea: GoTo changeSize
        ' NO USE
        Case TypeOf obj Is Excel.Range: Exit Sub
        Case TypeOf obj Is Excel.PlotArea: Exit Sub
    End Select
changeSize:
    obj.width = obj.width - offset
End Sub

Public Sub increaseWidth(Optional ByRef offset As Byte = 2)
    Dim obj As Object: Set obj = Selection
    Select Case True
        Case TypeOf obj Is Excel.Shape: GoTo changeSize
        Case TypeOf obj Is Excel.Rectangle: GoTo changeSize
        Case TypeOf obj Is Excel.Oval: GoTo changeSize
        Case TypeOf obj Is Excel.GroupObject: GoTo changeSize
        Case TypeOf obj Is Excel.DrawingObjects: GoTo changeSize
        Case TypeOf obj Is Excel.Picture: GoTo changeSize
        Case TypeOf obj Is Excel.line: GoTo changeSize
        Case TypeOf obj Is Excel.ChartArea: GoTo changeSize
        ' NO USE
        Case TypeOf obj Is Excel.Range: Exit Sub
        Case TypeOf obj Is Excel.PlotArea: Exit Sub
    End Select
changeSize:
    obj.width = obj.width + offset
End Sub

Public Sub decreaseHeight(Optional ByRef offset As Byte = 2)
    Dim obj As Object: Set obj = Selection
    Select Case True
        Case TypeOf obj Is Excel.Shape: GoTo changeSize
        Case TypeOf obj Is Excel.Rectangle: GoTo changeSize
        Case TypeOf obj Is Excel.Oval: GoTo changeSize
        Case TypeOf obj Is Excel.GroupObject: GoTo changeSize
        Case TypeOf obj Is Excel.DrawingObjects: GoTo changeSize
        Case TypeOf obj Is Excel.Picture: GoTo changeSize
        Case TypeOf obj Is Excel.line: GoTo changeSize
        Case TypeOf obj Is Excel.ChartArea: GoTo changeSize
        ' NO USE
        Case TypeOf obj Is Excel.Range: Exit Sub
        Case TypeOf obj Is Excel.PlotArea: Exit Sub
    End Select
changeSize:
    obj.height = obj.height - offset
End Sub

Public Sub increaseHeight(Optional ByRef offset As Byte = 2)
    Dim obj As Object: Set obj = Selection
    Select Case True
        Case TypeOf obj Is Excel.Shape: GoTo changeSize
        Case TypeOf obj Is Excel.Rectangle: GoTo changeSize
        Case TypeOf obj Is Excel.Oval: GoTo changeSize
        Case TypeOf obj Is Excel.GroupObject: GoTo changeSize
        Case TypeOf obj Is Excel.DrawingObjects: GoTo changeSize
        Case TypeOf obj Is Excel.Picture: GoTo changeSize
        Case TypeOf obj Is Excel.line: GoTo changeSize
        Case TypeOf obj Is Excel.ChartArea: GoTo changeSize
        ' NO USE
        Case TypeOf obj Is Excel.Range: Exit Sub
        Case TypeOf obj Is Excel.PlotArea: Exit Sub
    End Select
changeSize:
    obj.height = obj.height + offset
End Sub

Public Sub Test()
' Dim Test As PJ1_Logic
' Dim Test2 As Utils_Scripts
' Set Test = New PJ1_Logic
' Set Test2 = New Utils_Scripts
' Dim xx As Object: Set xx = Test2.WinSCP
' Call Test.Init(Excel.Application.ActiveSheet)
' Call Test.CheckFileName
' Call uatTest2()
' Call findShapeIntersectRange
Call sheetFocusRename
End Sub

' xxxxx

' Simple check for Windows
Private Function IsWindows() As Boolean
    IsWindows = (InStr(1, Application.OperatingSystem, "Windows", vbTextCompare) > 0)
End Function

' TODO xxx
Private Sub storeFormatA1()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("formatStored")
End Sub

' ----

' ' TODO: This can move to common
Private Function checkPathExist(ByVal path As String) As Boolean
    If Dir(path) <> "" Then
        Let checkPathExist = True
    Else
        Let checkPathExist = False
    End If
End Function

Private Sub CheckFileName()
    Dim ws As Worksheet: Set ws = ActiveSheet
    With ws
    Dim PASS As String: Let PASS = .Range("H1").value
    Dim FAIL As String: Let FAIL = .Range("I1").value
    Dim testCasePath As String: Let testCasePath = .Range("E6").value
    Dim checkListPath As String: Let checkListPath = .Range("E7").value
    Let .Range("F6").value = IIf(checkPathExist(testCasePath), PASS, FAIL)
    Let .Range("F7").value = IIf(checkPathExist(checkListPath), PASS, FAIL)
    End With
End Sub

Private Sub findDbClose()
    Dim wsh As Object: Set wsh = CreateObject("WScript.Shell")
    Dim NODE_PATH As String: Let NODE_PATH = "D:\environments\node\node.exe"
    Dim EXEC_PATH As String: Let EXEC_PATH = "D:\share\ASS\Tasks\template\find_db_close.js"
    Dim exec As Object: Set exec = wsh.exec("""" & NODE_PATH & """ """ & EXEC_PATH & """")
    Dim csLog As String: Let csLog = Replace(exec.StdOut.ReadAll, vbCrLf, vbLf)
    Dim lines() As String: Let lines = Split(csLog, vbLf)
    Dim startCell As Range: Set startCell = ActiveSheet.Range("P2")
    Dim startRow As Long: startRow = startCell.Row
    Dim startCol As Long: startCol = startCell.Column
    ActiveSheet.Range(startCell, ActiveSheet.Cells(startRow + 50, startCol)).ClearContents
    Dim i As Long
    For i = LBound(lines) To UBound(lines)
        If Trim(lines(i)) <> "" Then ActiveSheet.Cells(startRow + i, startCol).value = lines(i)
    Next i
End Sub

Private Sub findPhase1()
    Dim wsh As Object: Set wsh = CreateObject("WScript.Shell")
    Dim NODE_PATH As String: Let NODE_PATH = "D:\environments\node\node.exe"
    Dim EXEC_PATH As String: Let EXEC_PATH = "D:\share\ASS\Tasks\template\find_src_P1.js"
    Dim exec As Object: Set exec = wsh.exec("""" & NODE_PATH & """ """ & EXEC_PATH & """")
    Dim csLog As String: Let csLog = Replace(exec.StdOut.ReadAll, vbCrLf, vbLf)
    Dim lines() As String: Let lines = Split(csLog, vbLf)
    Dim startCell As Range: Set startCell = ActiveSheet.Range("R2")
    Dim startRow As Long: startRow = startCell.Row
    Dim startCol As Long: startCol = startCell.Column
    ActiveSheet.Range(startCell, ActiveSheet.Cells(startRow + 300, startCol)).ClearContents
    Dim i As Long
    For i = LBound(lines) To UBound(lines)
        If Trim(lines(i)) <> "" Then ActiveSheet.Cells(startRow + i, startCol).value = lines(i)
    Next i
End Sub

'Private Function getRelativePath(ByRef filename As String) As String
'    Dim shortName As String: Let shortName = left(filename, Len(filename) - 4)
'    Dim subPath As String: Let subPath = left(shortName, 4) & "\" & Mid(shortName, 5)
'    Let getRelativePath = "AP2_CPGM\" & subPath & "\" & filename
'End Function
Private Function getRelativePath(ByRef filename As String) As String
    Dim shortName As String: Let shortName = left(filename, InStrRev(filename, ".") - 1)
    Dim subPath As String: Let subPath = left(shortName, 4) & "\" & Mid(shortName, 5)
    Let getRelativePath = "AP2_CPGM\" & subPath & "\" & filename
End Function

Private Function getFullName(ByRef relativePath As String) As String
    Dim repo As String: Let repo = ActiveSheet.Range("U1").value
    Let getFullName = repo & relativePath
End Function

Private Sub commitDate(ByRef filename As String, ByRef target As String)
    Dim ws As Worksheet: Set ws = ActiveSheet
    With ws
    Dim path As String: Let path = getRelativePath(filename)
    Dim fullName As String: Let fullName = getFullName(path)
    If Not checkPathExist(fullName) Then
        Let ws.Range(target).value = "???"
        Exit Sub
    End If
    Dim wsh As Object: Set wsh = CreateObject("WScript.Shell")
    Dim repo As String: Let repo = .Range("U1").value
    Dim GIT As String: Let GIT = .Range("T1").value
    Dim Script As String: Let Script = "cmd /c cd /d """ & repo & """ && """ & GIT & """ log -1 --no-merges --format=""%ci"" -- """ & path & """"
    Dim exec As Object: Set exec = wsh.exec(Script)
    Dim fullDate As String: Let fullDate = exec.StdOut.ReadLine
    If fullDate <> "" Then
        Dim dateOnly As String: Let dateOnly = left(fullDate, 10)
        Let .Range(target).offset(0, -1).value = filename
        Let .Range(target).value = CDate(dateOnly)
        Let .Range(target).NumberFormat = "yyyy/mm/dd"
    End If
    End With
End Sub

Private Sub dateFromTestCase()
    Dim ws As Worksheet: Set ws = ActiveSheet
    With ws
    Dim fullName As String: Let fullName = .Range("E6").value
    Dim sheetName As String: Let sheetName = .Range("V1").value
    Dim targetCol As Long: Let targetCol = 1 ' B Col
    Dim sSQL As String
    Dim i As Byte
    If Not checkPathExist(fullName) Then
        For i = 16 To 18
            Let .Range("E" & i).value = "???"
        Next i
        Exit Sub
    End If
    If InStr(sheetName, " ") > 0 Then
        Let sSQL = "SELECT * FROM ['" & sheetName & "$']"
    Else
        Let sSQL = "SELECT * FROM [" & sheetName & "$]"
    End If
    Dim Connect As Object: Set Connect = CreateObject("ADODB.Connection")
    Dim Recordset As Object: Set Recordset = CreateObject("ADODB.Recordset")
    Call Connect.Open("Provider=Microsoft.ACE.OLEDB.12.0;" & _
             "Data Source=" & fullName & ";" & _
             "Extended Properties=""Excel 12.0 Xml;HDR=No;IMEX=1"";")
    Call Recordset.Open(sSQL, Connect, 3, 1)
    If Not Recordset.eof Then Call Recordset.MoveLast
    If Recordset.Fields.count <= targetCol Then Exit Sub
    Do While Not Recordset.BOF
        If InStr(Recordset.Fields(targetCol).value, "LINK END") > 0 Then Exit Do
        For i = 16 To 18
            If InStr(Recordset.Fields(targetCol).value, .Range("C" & i).value) > 0 Then
                Dim linePart() As String: linePart = Split(Application.Trim(Recordset.Fields(targetCol).value), Space(1))
' Dim j As Long: For j = LBound(linePart) To UBound(linePart): Debug.Print j & " '" & linePart(j) & "'": Next j 'xxx debug
                Dim yyyy As String: yyyy = IIf(Len(linePart(7)) = 4, linePart(7), Year(Date))
                Dim mm As String: mm = left$(linePart(5), Len(linePart(5)) - 1)
                Dim dd As String: dd = linePart(6)
                Let .Range("E" & i).value = Format(DateSerial(yyyy, mm, dd), "yyyy/mm/dd")
                
            End If
        Next i
        Call Recordset.MovePrevious
    Loop
    Call Recordset.Close
    Call Connect.Close
    End With
End Sub

Public Sub checkEncodeAndEOF(ByRef task As String)
    Dim ws As Worksheet: Set ws = ActiveSheet
    With ws
    Dim NKF As String: Let NKF = "D:\share\ASS\Tasks\template\autotest.exe"
    Dim guessParam As String: Let guessParam = "--guess"
    Dim fileNameSrc As String: Let fileNameSrc = task & ".src"
    Dim fileNameMak As String: Let fileNameMak = task & ".mak"
    Dim fileNameIni As String: Let fileNameIni = task & ".ini"
    Dim fileNameC As String: Let fileNameC = task & ".c"
    Dim fullNameSrc As String: Let fullNameSrc = getFullName(getRelativePath(fileNameSrc))
    Dim fullNameMak As String: Let fullNameMak = getFullName(getRelativePath(fileNameMak))
    Dim fullNameIni As String: Let fullNameIni = getFullName(getRelativePath(fileNameIni))
    Dim fullNameC As String: Let fullNameC = getFullName(getRelativePath(fileNameC))
    Let .Range("C27").value = fileNameSrc
    Dim wsh As Object: Set wsh = CreateObject("WScript.Shell")
    Dim Script As String: Let Script = """" & NKF & """ " & guessParam & " """ & _
        fullNameSrc & """ """ & _
        fullNameMak & """ """ & _
        fullNameIni & """ """ & _
        fullNameC & """"
    Dim exec As Object: Set exec = wsh.exec(Script)
    Do While Not exec.StdOut.AtEndOfStream
        Dim line As String: Let line = exec.StdOut.ReadLine
        Dim linePart() As String
        If InStr(line, fileNameSrc) > 0 Then
            Let linePart = Split(line, ":")
            Let .Range("D27").value = linePart(2)
            Let .Range("C27").value = fileNameSrc
        ElseIf InStr(line, fileNameMak) > 0 Then
            Let linePart = Split(line, ":")
            Let .Range("D28").value = linePart(2)
            Let .Range("C28").value = fileNameMak
        ElseIf InStr(line, fileNameIni) > 0 Then
            Let linePart = Split(line, ":")
            Let .Range("D29").value = linePart(2)
            Let .Range("C29").value = fileNameIni
        ElseIf InStr(line, fileNameC) > 0 Then
            Let linePart = Split(line, ":")
            Let .Range("D30").value = linePart(2)
            Let .Range("C30").value = fileNameC
        End If
    Loop
    End With
End Sub

Private Sub checkMakFile(ByRef task As String)
    Dim fullNameMak As String: Let fullNameMak = getFullName(getRelativePath(task & ".mak"))
    Dim fStream As Object: Set fStream = CreateObject("ADODB.Stream")
    Let fStream.Type = 2
    Let fStream.charset = "x-euc-jp"
    Call fStream.Open
    Call fStream.LoadFromFile(fullNameMak)
    Dim fileContent As String: Let fileContent = fStream.ReadText(-1)
    Call fStream.Close
    Dim lines() As String: Let lines = Split(Replace(fileContent, vbCrLf, vbLf), vbLf)
    Dim checkOKTime As Byte: Let checkOKTime = 0
    Dim i As Long: For i = LBound(lines) To UBound(lines)
        Dim line As String: Let line = lines(i)
        If InStr(line, ": " & task & ".m") > 0 Then
            If InStr(line, ": " & task & ".mak") > 0 Then
                Let checkOKTime = checkOKTime + 1
            Else
                Let checkOKTime = 0
                Exit For
            End If
        ElseIf InStr(line, "rm -f $(EXE) *.lis *.o") > 0 Then
            If InStr(line, "rm -f $(EXE) *.lis *.o *.c") > 0 Then
                Let checkOKTime = checkOKTime + 1
            Else
                Let checkOKTime = 0
                Exit For
            End If
        ElseIf InStr(line, "make CC=$(CC) -f $(EXE).m") > 0 Then
            If InStr(line, "make CC=$(CC) -f $(EXE).mak") > 0 Then
                Let checkOKTime = checkOKTime + 1
            Else
                Let checkOKTime = 0
                Exit For
            End If
        End If
    Next i
    Dim ws As Worksheet: Set ws = ActiveSheet
    With ws
    Dim PASS As String: Let PASS = .Range("H1").value
    Dim FAIL As String: Let FAIL = .Range("I1").value
    If checkOKTime = 3 Then
        Let .Range("F9").value = PASS
    Else
        Let .Range("F9").value = FAIL
    End If
    End With
End Sub

' TODO: This can move to common
Private Function diffCheck(fileName1 As String, fileName2 As String) As Boolean
    Const BLOCKSIZE As Long = 65536 ' 64KB
    Dim file1 As Object: Set file1 = CreateObject("ADODB.Stream")
    Let file1.Type = 1
    Call file1.Open
    Call file1.LoadFromFile(fileName1)
    Dim file2 As Object: Set file2 = CreateObject("ADODB.Stream")
    Let file2.Type = 1
    Call file2.Open
    Call file2.LoadFromFile(fileName2)
    Do While Not file1.EOS Or Not file2.EOS
        Dim block1() As Byte: Let block1 = file1.Read(BLOCKSIZE)
        Dim block2() As Byte: Let block2 = file2.Read(BLOCKSIZE)
        Dim bytesRead1 As Long: Let bytesRead1 = UBound(block1) - LBound(block1) + 1
        Dim bytesRead2 As Long: Let bytesRead2 = UBound(block2) - LBound(block2) + 1
        If bytesRead1 <> bytesRead2 Then
            Call file1.Close: Call file2.Close
            Let diffCheck = False
            Exit Function
        End If
        Dim i As Long
        For i = LBound(block1) To UBound(block1)
            If block1(i) <> block2(i) Then
                Call file1.Close: Call file2.Close
                Let diffCheck = False
                Exit Function
            End If
        Next i
    Loop
    Call file1.Close: Call file2.Close
    Let diffCheck = True
End Function

Private Sub checkDiffPhase1(ByRef task As String)
    Dim ws As Worksheet: Set ws = ActiveSheet
    With ws
    Dim AUTO As String: Let AUTO = .Range("J2").value
    If .Range("E32").value <> AUTO Then Exit Sub
    Dim PASS As String: Let PASS = .Range("H1").value
    Dim FAIL As String: Let FAIL = .Range("I1").value
    Dim fileNameSrc As String: Let fileNameSrc = task & ".src"
    Dim fileNameMak As String: Let fileNameMak = task & ".mak"
    Dim fileNameIni As String: Let fileNameIni = task & ".ini"
    Dim fullNameSrc As String: Let fullNameSrc = getFullName(getRelativePath(fileNameSrc))
    Dim fullNameMak As String: Let fullNameMak = getFullName(getRelativePath(fileNameMak))
    Dim fullNameIni As String: Let fullNameIni = getFullName(getRelativePath(fileNameIni))
    Dim fullNameSrcCP As String: Let fullNameSrcCP = Replace(fullNameSrc, "ass_bos2\AP2_CPGM", "BYCP_ASS_DEVELOP_CPGM")
    Dim fullNameMakCP As String: Let fullNameMakCP = Replace(fullNameMak, "ass_bos2\AP2_CPGM", "BYCP_ASS_DEVELOP_CPGM")
    Dim fullNameIniCP As String: Let fullNameIniCP = Replace(fullNameIni, "ass_bos2\AP2_CPGM", "BYCP_ASS_DEVELOP_CPGM")
    Let .Range("C33").value = fileNameSrc
    Let .Range("E33").value = IIf(diffCheck(fullNameSrc, fullNameSrcCP), PASS, FAIL)
    Let .Range("C34").value = fileNameMak
    Let .Range("E34").value = IIf(diffCheck(fullNameMak, fullNameMakCP), PASS, FAIL)
    Let .Range("C35").value = fileNameIni
    Let .Range("E35").value = IIf(diffCheck(fullNameIni, fullNameIniCP), PASS, FAIL)
    End With
End Sub

Private Function gitCheckout(ByRef GIT As String, ByRef repoPath As String, ByRef key As String) As String
    Dim wsh As Object: Set wsh = CreateObject("WScript.Shell")
    Dim Script As String: Let Script = "cmd /c cd /d """ & repoPath & """ && """ & GIT & """ fetch"
    Call wsh.Run(Script, 0, True)
    Let Script = "cmd /c cd /d """ & repoPath & """ && """ & GIT & """ branch -r | findstr /i " & key
    Dim exec As Object: Set exec = wsh.exec(Script)
    Dim gitLog As String: Let gitLog = Trim(Replace(exec.StdOut.ReadAll, vbCrLf, vbLf))
    If gitLog = vbNullString Then
        Let gitCheckout = "0-#-NOT FIND"
        Exit Function
    End If
    Dim lines() As String: Let lines = Split(Trim(gitLog), vbLf)
    Do While UBound(lines) >= 0 And Trim(lines(UBound(lines))) = ""
        ReDim Preserve lines(LBound(lines) To UBound(lines) - 1)
    Loop
'Dim j As Long: For j = LBound(lines) To UBound(lines): Debug.Print j & " '" & lines(j) & "'": Next j 'xxx debug
    If UBound(lines) > 0 Then
        Let gitCheckout = "0-#- > 1"
        Exit Function
    End If
    Dim line As String: Let line = Trim(lines(0))
    Let Script = "cmd /c cd /d """ & repoPath & """ && """ & GIT & """ checkout " & line

    Set exec = wsh.exec(Script)
    Do While exec.status = 0: DoEvents: Loop
' Let gitLog = exec.StdOut.ReadAll & vbCrLf & exec.StdErr.ReadAll: Debug.Print "gitLog"; gitLog
    If exec.exitCode <> 0 Then
        Let gitCheckout = "0-#-CHECKOUT fail"
    Else
        Let gitCheckout = "1-#-" & line
    End If
End Function

Private Sub gitCheckoutByKey(ByRef task As String)
    Dim ws As Worksheet: Set ws = ActiveSheet
    With ws
    Dim AUTO As String: Let AUTO = .Range("J2").value
    Dim MANUAL As String: Let MANUAL = .Range("J1").value
    Dim repo As String: Let repo = .Range("U1").value
    Dim GIT As String: Let GIT = .Range("T1").value
    Dim pullResult() As String: Let pullResult = Split(gitCheckout(GIT, repo, task), "-#-")
    Let .Range("D3").value = IIf(pullResult(0) = "0", MANUAL, AUTO)
    Let .Range("E3").value = pullResult(1)
    End With
End Sub

Public Sub uatTest()
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
    
Cleanup:
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
End Sub

Public Sub uatTest2()
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim Tst As PJ1_Logic: Set Tst = New PJ1_Logic
    Call Tst.Init(ws)
    
    ' Call Tst.ChecklistExits
    ' Call Tst.UTExits
    ' Call Tst.CountInUT
    ' Call Tst.CountInOutput
    ' Call Tst.SfcmaplgExist
    ' Call Tst.SfcmerlgExist
    ' Call Tst.LogExist

    ' Call Tst.LogEOF ' lac
    ' Call Tst.SfcmaplgEOF ' lac
    ' Call Tst.SfcmerlgEOF ' lac

    ' Call Tst.CheckLogContent
    ' Call Tst.LastRuntime
    ' Call Tst.GetTestcsh
    ' Call Tst.GetTestcshExec
    ' Call Tst.GetALELog
    ' Call Tst.CheckSfcmaplgOutput
    ' Call Tst.CheckSfcmerlgOutput

    ' Call Tst.CheckSfcmaplgLine
    ' Call Tst.CheckSfcmerlgLine
    ' Call Tst.CheckSfcmerlgError

    ' Call Tst.CheckElsePattern ' rat lac

    ' HELPER
    ' (deprecated) Note: Use Node Faster
    ' Call Tst.ExportImages 

    Call Tst.XXXTEST

    With ws
    End With
Cleanup:
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
End Sub

Private Sub findShapeIntersectRange()
    Dim picC As PicturesController
    Set picC = New PicturesController
    Call picC.selectShapeInRange(picC.OVERLAP_MODE)
End Sub

Private Sub findShapeInRange()
    Dim picC As PicturesController
    Set picC = New PicturesController
    Call picC.selectShapeInRange(picC.INSIDE_MODE)
End Sub