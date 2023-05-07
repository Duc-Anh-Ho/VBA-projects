Attribute VB_Name = "Shortcuts"
Option Explicit

'METHODS
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

Private Sub shapeMoveWIthSize()
    'TODO
End Sub

'MAIN

Public Sub add()
    'TODO
End Sub

Public Sub remove()
    'TODO
End Sub

Public Sub install()
    Application.OnKey _
        key:="^+{C}", _
        procedure:="Shortcuts.copyF"
    Application.OnKey _
        key:="^+{V}", _
        procedure:="Shortcuts.pasteF"
    Application.OnKey _
        key:="^{TAB}", _
        procedure:="Shortcuts.sheetSelectN"
    Application.OnKey _
        key:="^+{TAB}", _
        procedure:="Shortcuts.sheetSelectP"
End Sub

Public Sub uninstall()
    Application.OnKey key:="^+{C}"
    Application.OnKey key:="^+{V}"
    Application.OnKey key:="^{TAB}"
    Application.OnKey key:="^+{TAB}"
End Sub
