Attribute VB_Name = "Shortcuts"
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

Private Sub multipleReplace()
    Dim form As MultipleReplaceForm
    Set form = New MultipleReplaceForm
    form.Show vbModal ' vbModeless or vbModal
    Set form = Nothing
End Sub

'MAIN
' TODO: Menu ribbon create shortcut
Public Sub add()
    'TODO:
End Sub

Public Sub remove()
    'TODO
End Sub

Public Sub install()
    ' Ctrl + Shift +C
    Application.OnKey _
        key:="^+{C}", _
        procedure:="Shortcuts.copyF"
    ' Ctrl + Shift + V
    Application.OnKey _
        key:="^+{V}", _
        procedure:="Shortcuts.pasteF"
    ' Ctrl + Shift + Alt + V
    Application.OnKey _
        key:="^+%{V}", _
        procedure:="Shortcuts.pasteV"
    ' Ctrl + Tab
    Application.OnKey _
        key:="^{TAB}", _
        procedure:="Shortcuts.sheetSelectN"
    ' Ctrl + Shift + Tab
    Application.OnKey _
        key:="^+{TAB}", _
        procedure:="Shortcuts.sheetSelectP"
    ' Ctrl + M -> Duplicate system shortkey
    Application.OnKey _
        key:="^{M}", _
        procedure:="Shortcuts.shapeMoveAndSize"
    ' Ctrl + Shift + M
    Application.OnKey _
        key:="^+{M}", _
        procedure:="Shortcuts.shapeMove"
    ' Ctrl + Shift + Alt + M
    Application.OnKey _
        key:="^+%{M}", _
        procedure:="Shortcuts.shapeFree"
    ' Ctrl + Delete
    Application.OnKey _
        key:="^{DEL}", _
        procedure:="Shortcuts.clearFormat"
    ' Ctrl + Shift + Delete
    Application.OnKey _
        key:="^+{DEL}", _
        procedure:="Shortcuts.clearAll"
    ' Shift + F12
    Application.OnKey _
        key:="+{F12}", _
        procedure:="Shortcuts.copyFullName"
    ' Ctrl + Shift + Alt  + C
    Application.OnKey _
        key:="^+%{C}", _
        procedure:="Shortcuts.copyFullName"
    ' Ctrl + Shift + S
    Application.OnKey _
        key:="^+{S}", _
        procedure:="Shortcuts.copyPath"
    ' F11
    Application.OnKey _
        key:="{F11}", _
        procedure:="Shortcuts.toggleZenMode"
    ' F1 - Override Help
    Application.OnKey _
        key:="{F1}", _
        procedure:=""
    ' Ctrl + Shift + H
    Application.OnKey _
        key:="^+{H}", _
        procedure:="Shortcuts.multipleReplace"
End Sub

Public Sub unInstall()
    Application.OnKey key:="^+{C}"
    Application.OnKey key:="^+{V}"
    Application.OnKey key:="^+%{V}"
    Application.OnKey key:="^{TAB}"
    Application.OnKey key:="^+{TAB}"
    Application.OnKey key:="^{M}"
    Application.OnKey key:="^+{M}"
    Application.OnKey key:="^+%{M}"
    Application.OnKey key:="^{DEL}"
    Application.OnKey key:="^+{DEL}"
    Application.OnKey key:="+{F12}"
    Application.OnKey key:="^+{S}"
    Application.OnKey key:="{F11}"
End Sub
