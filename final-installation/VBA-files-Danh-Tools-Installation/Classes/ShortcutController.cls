VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ShortcutController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Declare Variables
Private shortcutWs As Worksheet
Private shortcutTb As ListObject
Private row As ListRow
Private Const SHORTCUT_MODULE_NAME As String = "Shortcuts."
Private Const SHORTCUT_SHEET_NAME As String = "keyboard-shortcut"
Private Const SHORTCUT_TABLE_NAME As String = "Keybinding"
Private Const NO_COLUMN As Byte = 1
Private Const PROCEDURE_COLUMN As Byte = 2
Private Const DEFAULT_KEYBIND_COLUMN As Byte = 3
Private Const CUSTOM_KEYBIND_COLUMN As Byte = 4
Private Const COMMAND_COLUMN As Byte = 5
Private Const WHEN_COLUMN As Byte = 6
Private Const SOURCE_COULMN As Byte = 7
Private Const CTRL_CODE As String = "^"
Private Const SHIFT_CODE As String = "+"
Private Const ALT_CODE As String = "%"
Private Const ENTER_CODE As String = "~"
Private Const BACKSPACE_CODE As String = "{BS}"
Private Const BREAK_CODE As String = "{BREAK}"
Private Const CLEAR_CODE As String = "{CLEAR}"
Private Const DELETE_CODE As String = "{DEL}"
Private Const END_CODE As String = "{END}"
Private Const ENTER_NUMPAD_CODE As String = "{ENTER}"
Private Const ESCAPE_CODE As String = "{ESC}"
Private Const HELP_CODE As String = "{HELP}"
Private Const HOME_CODE As String = "{HOME}"
Private Const INS_CODE As String = "{INSERT}"
Private Const RETURN_CODE As String = "{RETURN}"
Private Const CAPSLOCK_CODE As String = "{CAPSLOCK}"
Private Const NUMLOCK_CODE As String = "{NUMLOCK}"
Private Const SCROLLLOCK_CODE As String = "{SCROLLLOCK}"
Private Const TAB_CODE As String = "{TAB}"
Private Const UP_ARROW_CODE As String = "{UP}"
Private Const DOWN_ARROW_CODE As String = "{DOWN}"
Private Const LEFT_ARROW_CODE As String = "{LEFT}"
Private Const RIGHT_ARROW_CODE As String = "{RIGHT}"
Private Const PAGE_DOWN_CODE As String = "{PGDN}"
Private Const PAGE_UP_CODE As String = "{PGUP}"
'Private Const COMMAND_CODE As String = "*" 'MAC below Excel 2011
Private Const MUILTIPLE_KEY As String = "Mult(Num)"
Private Const PLUS_KEY As String = "Plus(Num)"
Private Const ENTERNUM_KEY As String = "Enter(Num)"
Private Const SUBTRACT_KEY As String = "Sub(Num)"
Private Const DOT_KEY As String = "Dot(Num)"
Private Const DIVIDE_KEY As String = "Div(Num)"
Private Const BACKSPACE_KEY As String = "Backspace"
Private Const BREAK_KEY As String = "Break"
Private Const CLEAR_KEY As String = "Clear"
Private Const MENU_KEY As String = "Menu"
Private Const PAUSE_KEY As String = "Pause"
Private Const SPACE_KEY As String = "Space"
Private Const PAGEUP_KEY As String = "PageUp"
Private Const PAGEDOWN_KEY As String = "PageDown"
Private Const END_KEY As String = "End"
Private Const HOME_KEY As String = "Home"
Private Const PRINT_KEY As String = "Print"
Private Const SELECT_KEY As String = "Select"
Private Const EXECUTE_KEY As String = "Exc"
Private Const SMAPSHOT_KEY As String = "Snap"
Private Const INSERT_KEY As String = "Insert"
Private Const DELETE_KEY As String = "Delete"
Private Const HELP_KEY As String = "Help"
Private Const ENTER_NUMPAD_KEY As String = "Enter(Num)"
Private Const ENTER_KEY As String = "Enter"
Private Const ESCAPE_KEY As String = "Escape"
Private Const INS_KEY As String = "Insert"
Private Const RETURN_KEY As String = "Return"
Private Const CAPSLOCK_KEY As String = "Capslock"
Private Const NUMLOCK_KEY As String = "Numlock"
Private Const SCROLLLOCK_KEY As String = "Scrolllock"
Private Const TAB_KEY As String = "Tab"
Private Const UP_ARROW_KEY As String = "Up Arrow"
Private Const DOWN_ARROW_KEY As String = "Down Arrow"
Private Const LEFT_ARROW_KEY As String = "Left Arrow"
Private Const RIGHT_ARROW_KEY As String = "Right Arrow"
Private Const PAGE_DOWN_KEY As String = "Page Down"
Private Const PAGE_UP_KEY As String = "Page Up"
Private Const COMMAND_KEY As String = "Command"
Private Const PLUS As String = " + "
Private Const CTRL_KEY As String = "Ctrl" & PLUS
Private Const SHIFT_KEY As String = "Shift" & PLUS
Private Const ALT_KEY As String = "Alt" & PLUS
Private Const CTRL_SHIFT_KEY As String = CTRL_KEY & SHIFT_KEY
Private Const CTRL_ALT_KEY As String = CTRL_KEY & ALT_KEY
Private Const CTRL_SHIFT_ALT_KEY As String = CTRL_SHIFT_KEY & ALT_KEY
Private Const SHIFT_ALT_KEY As String = SHIFT_KEY & ALT_KEY
Private Const SHIFT_MASK As Byte = 1
Private Const CTRL_MASK As Byte = 2
Private Const ALT_MASK As Byte = 4

' Constructor
Private Sub Class_Initialize()
    Set shortcutWs = ThisWorkbook.Sheets(SHORTCUT_SHEET_NAME)
    Set shortcutTb = shortcutWs.ListObjects(SHORTCUT_TABLE_NAME)
End Sub

' Destructor
Private Sub Class_Terminate()
    Set shortcutWs = Nothing
    Set shortcutTb = Nothing
End Sub

' ACCESSORS
Public Property Get getShortcutTable() As ListObject
    Set shortcutWs = ThisWorkbook.Sheets(SHORTCUT_SHEET_NAME)
    Set getShortcutTable = shortcutWs.ListObjects(SHORTCUT_TABLE_NAME)
End Property

Public Property Get getNoColumn() As Byte
    Let getNoColumn = NO_COLUMN
End Property

Public Property Get getProcedureColumn() As Byte
    Let getProcedureColumn = PROCEDURE_COLUMN
End Property

Public Property Get getDefaultKeybindColumn() As Byte
    Let getDefaultKeybindColumn = DEFAULT_KEYBIND_COLUMN
End Property

Public Property Get getCustomKeybindColumn() As Byte
    Let getCustomKeybindColumn = CUSTOM_KEYBIND_COLUMN
End Property

Public Property Get getCommandColumn() As Byte
    Let getCommandColumn = COMMAND_COLUMN
End Property

Public Property Get getWhenColumn() As Byte
    Let getWhenColumn = WHEN_COLUMN
End Property

Public Property Get getSourceCoulmn() As Byte
    Let getSourceCoulmn = SOURCE_COULMN
End Property

' MAIN
Private Sub add(ByRef key As String, ByRef procedure As String)
    Application.OnKey key:=key, procedure:=SHORTCUT_MODULE_NAME & procedure
End Sub
Private Sub remove(ByRef key As String)
    Application.OnKey key:=key
End Sub

Public Sub install()
    'NOTE: Shortcuts is a module for calling as full global scope
    Dim procedure As String
    Dim key As String
    For Each row In shortcutTb.ListRows
        Let procedure = row.Range(1, getProcedureColumn()).Value
        Let key = row.Range(1, getCustomKeybindColumn()).Value
        Call add(key, procedure)
    Next row
End Sub

Public Sub unInstall()
    Dim key As String
    For Each row In shortcutTb.ListRows
        Let key = row.Range(1, getCustomKeybindColumn()).Value
        Call remove(key)
    Next row
End Sub


Public Function convertCodeToName(ByRef code As String) As String
    Let code = Replace(code, SHIFT_CODE, SHIFT_KEY) 'Shift key is "+" so much be replace first
    Let code = Replace(code, CTRL_CODE, CTRL_KEY)
    Let code = Replace(code, ALT_CODE, ALT_KEY)
    Let code = Replace(code, ENTER_CODE, ENTER_KEY)
    Let code = Replace(code, BACKSPACE_CODE, BACKSPACE_KEY)
    Let code = Replace(code, BREAK_CODE, BREAK_KEY)
    Let code = Replace(code, CLEAR_CODE, CLEAR_KEY)
    Let code = Replace(code, DELETE_CODE, DELETE_KEY)
    Let code = Replace(code, END_CODE, END_KEY)
    Let code = Replace(code, ENTER_NUMPAD_CODE, ENTER_NUMPAD_KEY)
    Let code = Replace(code, ESCAPE_CODE, ESCAPE_KEY)
    Let code = Replace(code, HELP_CODE, HELP_KEY)
    Let code = Replace(code, HOME_CODE, HOME_KEY)
    Let code = Replace(code, INS_CODE, INS_KEY)
    Let code = Replace(code, RETURN_CODE, RETURN_KEY)
    Let code = Replace(code, CAPSLOCK_CODE, CAPSLOCK_KEY)
    Let code = Replace(code, NUMLOCK_CODE, NUMLOCK_KEY)
    Let code = Replace(code, SCROLLLOCK_CODE, SCROLLLOCK_KEY)
    Let code = Replace(code, TAB_CODE, TAB_KEY)
    Let code = Replace(code, UP_ARROW_CODE, UP_ARROW_KEY)
    Let code = Replace(code, DOWN_ARROW_CODE, DOWN_ARROW_KEY)
    Let code = Replace(code, LEFT_ARROW_CODE, LEFT_ARROW_KEY)
    Let code = Replace(code, RIGHT_ARROW_CODE, RIGHT_ARROW_KEY)
    Let code = Replace(code, PAGE_DOWN_CODE, PAGE_DOWN_KEY)
    Let code = Replace(code, PAGE_UP_CODE, PAGE_UP_KEY)
    Let code = Replace(code, "{", vbNullString) 'TODO: use constant
    Let code = Replace(code, "}", vbNullString) 'TODO: use constant
    Let convertCodeToName = code
End Function

Public Function convertKeyToName(ByRef keycode As MSForms.ReturnInteger, ByRef Shift As Integer) As String
    Dim modifier As String
    Dim key As String
    'Listen modifier
    Select Case Shift
        Case (SHIFT_MASK): Let modifier = SHIFT_KEY
        Case (CTRL_MASK): Let modifier = CTRL_KEY
        Case (ALT_MASK): Let modifier = ALT_KEY
        Case (CTRL_MASK + SHIFT_MASK): Let modifier = CTRL_SHIFT_KEY
        Case (CTRL_MASK + ALT_MASK): Let modifier = CTRL_ALT_KEY
        Case (CTRL_MASK + SHIFT_MASK + ALT_MASK): Let modifier = CTRL_SHIFT_ALT_KEY
        Case (SHIFT_MASK + ALT_MASK): Let modifier = SHIFT_ALT_KEY
    End Select
    'Listen main key
    Select Case keycode
        Case vbKeyF1 To vbKeyF16: key = "F" & (keycode - vbKeyF1 + 1) 'TODO: use constant
        Case vbKeyA To vbKeyZ: key = Chr(keycode)
        Case vbKey0 To vbKey9: key = Chr(keycode)
        Case vbKeyNumpad0 To vbKeyNumpad9: key = Chr(keycode) & "(Num)"
        Case vbKeyBack: key = BACKSPACE_KEY
        Case vbKeyTab: key = TAB_KEY
        Case vbKeyReturn: key = ENTER_KEY 'vbKeyReturn is Enter not Backspace, Backspace key is vbKeyBack
        Case vbKeyControl: key = vbNullString
        Case vbKeyShift: key = vbNullString
        Case vbKeyMenu: key = vbNullString 'vbKeyMenu is alt not menu, menu key is 93(0x5D)
        Case Else: key = "Unknown(" & keycode & ")"
    End Select
    'Skip modifier only
    If key = vbNullString Then
        Let convertKeyToName = vbNullString
    Else
        Let convertKeyToName = modifier & key
    End If
End Function

'Public Function convertNameToCode(ByRef name As String) As String
'    TODO
'
'End Function
