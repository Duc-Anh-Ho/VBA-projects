Attribute VB_Name = "Shortcuts"
' Must put Onkey procedure methods in a module for calling as full global scope
Option Explicit

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
            MsgBox "Unhandled Type Name: " & TypeName(target): Exit Sub
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
        Let Selection.Cells(1, 1).Offset(index, 0).NumberFormat = "@"
        Let Selection.Cells(1, 1).Offset(index, 0).value = shp.name
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
        Let name = Selection.Cells(1, 1).Offset(index, 0).value
        If RTrim(name) = vbNullString Then
            Let shp.name = "index_" & index
        Else
            Let shp.name = Selection.Cells(1, 1).Offset(index, 0).value
        End If
        Let index = index + 1
        If shp.Type = msoGroup Then Call renameAllShapes(shp.GroupItems, index)
    Next shp
End Sub

Private Sub showListSheets()
    With Application.CommandBars("Workbook tabs")
        If .controls(.controls.Count).caption = "More Sheets..." Then
            .controls(.controls.Count).Execute
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
        Selection(Selection.Count).Row = lastData.Row, _
        area.Row + area.Rows.Count - 1, _
        Selection(Selection.Count).Row _
    )
    Dim target As Range: Set target = IIf( _
        Selection(Selection.Count).Row = lastData.Row, _
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
    Dim maxShape As Long: Let maxShape = ws.Shapes.Count
    Dim i As Long, j As Long, groupRootArr() As Long
    ' Init roots group (self root)
    ReDim groupRootArr(1 To maxShape)
    For i = 1 To ws.Shapes.Count
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
    Dim maxShape As Long: Let maxShape = ws.Shapes.Count
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
        For j = 1 To rootListColl.Count
            If rootListColl(j) = groupId Then
                Set currentListColl = groupListColl(j)
                Call currentListColl.add(IIf(mode = 0, ws.Shapes(i).id, i))
                GoTo NextShape
            End If
        Next j
        Set currentListColl = New Collection
        Call currentListColl.add(IIf(mode = 0, ws.Shapes(i).id, i))
        Call rootListColl.add(groupId)
        Call groupListColl.add(currentListColl)
NextShape:
    Next i
    Set getGroupShapeColl = groupListColl
End Function

' TODO: Delete test
Private Sub TESTgetGroupShape()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim g, id, shp, line$: Set g = getGroupShapeColl(ws)
    Dim i&: For i = 1 To g.Count
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
    For i = 1 To groupColl.Count
        Set idColl = groupColl(i)
        For j = 1 To idColl.Count
            Let shapeId = idColl(j)
            For Each sh In ws.Shapes
                If sh.id = shapeId Then
                    Let sh.name = format(sh.id, prefix) & "_" & sh.name ' Add ID prefix
                    Exit For
                End If
            Next sh
        Next j
    Next i
End Sub

Private Sub groupIntersectedShapes(ByVal ws As Worksheet)
    Dim maxShape As Long: Let maxShape = ws.Shapes.Count
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
    For i = 1 To groupColl.Count
        Set idColl = groupColl(i)
        If idColl.Count < 2 Then GoTo SkipGroup
        ReDim nameArr(1 To idColl.Count)
        For j = 1 To idColl.Count
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
    Let timestamp = format(Now, "yyyymmdd_hhnnss")
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
    For i = ws.Shapes.Count To 1 Step -1 ' Reverse to avoid index shift
        Set sh = ws.Shapes(i)
        Let sh.placement = placement
        If sh.Type = msoGroup Then
            ' Save tmp position
            Let leftTmp = sh.left
            Let topTmp = sh.top
            ' Convert to image by copy paste as pic
            Call sh.CopyPicture(Appearance:=xlScreen, format:=xlPicture)
            Call ws.Paste
            ' Get newest sh and move to saved position
            Set newSh = ws.Shapes(ws.Shapes.Count)
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
    
Private Sub formatFileEnd()
    Dim wb As Workbook: Set wb = ActiveWorkbook
    Dim ws As Worksheet
    Dim i As Integer
    If wb Is Nothing Then Exit Sub
    If wb.Worksheets.Count = 0 Then Exit Sub
    If Not isBackupSaved Then Exit Sub
    Application.ScreenUpdating = False
    For i = wb.Sheets.Count To 1 Step -1
        Set ws = wb.Sheets(i)
        If ws.Type = xlWorksheet Then
            Call ws.Activate
            Call ws.Range("A1").Select
            Call zoomMode100
            Call groupIntersectedShapes(ws)
            Call convertGroupToImage(ws)
        End If
    Next i
    Application.ScreenUpdating = True
End Sub

Private Sub testShareX()
    Dim picC As PicturesController
    Set picC = New PicturesController
    Call picC.snipShareX
End Sub

Public Sub test()
    ' Application.OnKey "^P", "showListWorkbooks"
    Call testShareX
End Sub

' TODO xxx
Private Sub storeFormatA1()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("formatStored")
End Sub