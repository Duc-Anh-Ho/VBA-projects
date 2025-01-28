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

Private Sub multipleReplace()
    Dim form As MultipleReplaceForm
    Set form = New MultipleReplaceForm
    form.Show vbModal ' vbModeless or vbModal
    Set form = Nothing
End Sub
