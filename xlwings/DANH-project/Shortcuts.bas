Attribute VB_Name = "Shortcuts"
'Must put Onkey procedure methods in a module for calling as full global scope
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
        Case TypeOf target Is Excel.Line: Exit Sub
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
    Dim shp As Shape
    ' target can be Shapes or GroupShapes
    Set target = IIf(target Is Nothing, ActiveSheet.Shapes, target)
    If index = 0 Then
        Let Selection.Cells(1, 1).value = "Selection Pane"
        Let index = index + 1
    End If
    For Each shp In target
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

Private Function isIntersect(sh1 As Shape, sh2 As Shape) As Boolean
    Dim top1 As Double, right1 As Double, bottom1 As Double, left1 As Double
    Dim top2 As Double, right2 As Double, bottom2 As Double, left2 As Double
    With sh1
    Let top1 = .top
    Let left1 = .left
    Let right1 = left1 + .width
    Let bottom1 = top1 + .height
    End With
    With sh2
    Let top2 = .top
    Let left2 = .left
    Let right2 = left2 + .width
    Let bottom2 = top2 + .height
    End With
    Let isIntersect = Not ( _
        right1 < left2 _
        Or right2 < left1 _
        Or bottom1 < top2 _
        Or bottom2 < top1 _
    )
End Function

Private Function getIntersectedColl(ws) As Collection
    Dim maxShape As Long: Let maxShape = ws.Shapes.Count
    ' No Shape handle
    If maxShape = 0 Then Exit Function
    Dim groupColl As Collection ' Parent Coll
    Dim duplicatedColl As Collection ' Check Coll
    Dim shapeIndexColl As Collection ' Child Coll
    Dim i As Long, j As Long, k As Long
    Set groupColl = New Collection
    Set duplicatedColl = New Collection
    For i = maxShape To 1 Step -1
        For j = 1 To duplicatedColl.Count
            If duplicatedColl(j) = i Then GoTo Next_i
        Next j
        Set shapeIndexColl = New Collection
        Call shapeIndexColl.Add(i)
        For j = i - 1 To 1 Step -1
            If isIntersect(ws.Shapes(i), ws.Shapes(j)) Then
                Call shapeIndexColl.add(j)
                Call duplicatedColl.add(j)
                Exit For
            End If
        Next j
        If shapeIndexColl.Count > 1 Then Call groupColl.add(shapeIndexColl)
Next_i:
    Next i
    Set getIntersectedColl = groupColl
End Function

Private Sub unGroupAllShapes
End Sub

Private Sub unGroup(group As Variant)
    Dim sh As Shape
    If TypeOf group Is Excel.Shapes Then
        For Each sh In group
            Call unGroup(sh)
        Next
    ElseIf TypeOf group Is Excel.Shape And group.Type = msoGroup Then
        For Each sh In group.Ungroup
            Call unGroup(sh)
        Next
    End If
End Sub

Private Sub groupAllShapes()
    ' Closed all sheets or called from the other app handle
    If ActiveSheet Is Nothing Then Exit Sub
    Call unGroup(ActiveSheet.Shapes)
    Dim groupColl As Collection
    Dim shapeIndexColl As Collection
    Set groupColl = New Collection
    Set groupColl = getIntersectedColl(ActiveSheet)
    ' No intersect handle
    Dim maxGroup As Long: Let maxGroup = groupColl.Count
    Dim maxItem As Long
    If maxGroup = 0 Then Exit Sub
    Dim shapeArray() As Long
    Dim i As Long, j As Long
    For i = 1 To maxGroup
        Set shapeIndexColl = New Collection
        Set shapeIndexColl = groupColl(i)
        Let maxItem = shapeIndexColl.Count
        ReDim shapeArray(1 To maxItem)
        For j = 1 To maxItem
            Let shapeArray(j) = shapeIndexColl(j)
        Next j
    Next i
    ActiveSheet.Shapes.Range(shapeArray).Group
End Sub


Public Sub test()
    ' Application.OnKey "^P", "showListWorkbooks"
    Call groupAllShapes
End Sub

' TODO
Private Sub storeFormatA1()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("formatStored")
End Sub
