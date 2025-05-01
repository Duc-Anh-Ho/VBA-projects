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
Private rows As ListRows
Private cols as ListColumns
Private Const SHORTCUT_MODULE_NAME As String = "Shortcuts."
Private Const SHORTCUT_SHEET_NAME As String = "keyboard-shortcut"
Private Const SHORTCUT_TABLE_NAME As String = "Keybinding"
Private Enum COL
    no = 1
    procedure = no + 1
    default_keybind = procedure + 1
    custom_keybind = default_keybind + 1
    command = custom_keybind + 1
    when = command + 1
    status = when + 1
End Enum
Private ENUM MASK
    none = 0
    shiftKey =  1
    ctrlKey = 2
    altKey = 4
End Enum

Private Const NO_SET As String = "--No Set--"
Private Const NUMPAD_MARK AS String = "(Num)"
Private Const OPEN_CURLY_BRACKET As String = "{"
Private Const CLOSE_CURLY_BRACKET As String = "}"
Private Const CTRL_CODE As String = "^"
Private Const SHIFT_CODE As String = "+"
Private Const ALT_CODE As String = "%"
Private Const LEN_MODIFIER_CODE AS Byte = 1
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
Private Const ENTER_KEY As String = "Enter"
Private Const ENTER_NUMPAD_KEY As String = "Enter" & NUMPAD_MARK
Private Const MULTIPLE_KEY As String = "Multi" & NUMPAD_MARK
Private Const PLUS_KEY As String = "Plus" & NUMPAD_MARK
Private Const SUBTRACT_KEY As String = "Sub" & NUMPAD_MARK
Private Const DOT_KEY As String = "Dot" & NUMPAD_MARK
Private Const DIVIDE_KEY As String = "Div" & NUMPAD_MARK
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
Private Const FUNCTION_KEY As String = "F"
Private Const PLUS As String = " + "
Private Const CTRL_KEY As String = "Ctrl" & PLUS
Private Const SHIFT_KEY As String = "Shift" & PLUS
Private Const ALT_KEY As String = "Alt" & PLUS
Private Const CTRL_SHIFT_KEY As String = CTRL_KEY & SHIFT_KEY
Private Const CTRL_ALT_KEY As String = CTRL_KEY & ALT_KEY
Private Const CTRL_SHIFT_ALT_KEY As String = CTRL_SHIFT_KEY & ALT_KEY
Private Const SHIFT_ALT_KEY As String = SHIFT_KEY & ALT_KEY
' Loop iterators
Private row As ListRow

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

Public Property Get getNoSet() As String: Let getNoSet = NO_SET: End Property
Public Property Get getNoCol() As Byte: Let getNoCol = COL.no: End Property
Public Property Get getProcedureCol() As Byte: Let getProcedureCol = COL.procedure: End Property
Public Property Get getDefaultKeybindingCol() As Byte: Let getDefaultKeybindingCol = COL.default_keybind: End Property
Public Property Get getCustomKeybindingCol() As Byte: Let getCustomKeybindingCol = COL.custom_keybind: End Property
Public Property Get getCommandCol() As Byte: Let getCommandCol = COL.command: End Property
Public Property Get getWhenCol() As Byte: Let getWhenCol = COL.when: End Property
Public Property Get getStatusCol() As Byte: Let getStatusCol = COL.status: End Property
Public Property Get getEnterKey() As String: Let getEnterKey = ENTER_KEY: End Property
Public Property Get getEscKey() As String: Let getEscKey = ESCAPE_KEY: End Property
Public Property Get getBackspaceKey() As String: Let getBackspaceKey = BACKSPACE_KEY: End Property
Public Property Get getShortcutTable() As ListObject
    Set shortcutWs = ThisWorkbook.Sheets(SHORTCUT_SHEET_NAME)
    Set getShortcutTable = shortcutWs.ListObjects(SHORTCUT_TABLE_NAME)
End Property
Public Property Get getRows() As ListRows: Set getRows = getShortcutTable().ListRows: End Property
Public Property Get getCols() As ListColumns: Set getCols = getShortcutTable().ListColumns: End Property

' MUTATORS

Public Sub setColData(ByRef data() As String, colNo As Byte)
    Let getCols()(colNo).DataBodyRange.value = data
End Sub

' MAIN

Private Sub add(ByRef key As String, ByRef procedure As String)
    Application.OnKey key:=key, procedure:=SHORTCUT_MODULE_NAME & procedure
End Sub

Private Sub restore(ByRef key As String)
    Application.OnKey key:=key
End Sub

Private Sub remove(ByRef key As String)
    Application.OnKey key:=key, procedure:=vbNullString
End Sub

Public Sub install()
    ' NOTE: Shortcuts is a module for calling as full global scope
    Dim procedure As String
    Dim key As String
    For Each row In shortcutTb.ListRows
        Let procedure = row.Range(1, getProcedureCol()).value
        Let key = row.Range(1, getCustomKeybindingCol()).value
        ' Skip Blank Keybinding
        If key <> vbNullString Then Call add(key, procedure)
    Next row
End Sub

Public Sub unInstall()
    Dim key As String
    For Each row In shortcutTb.ListRows
        Let key = row.Range(1, getCustomKeybindingCol()).value
        If key <> vbNullString Then Call restore(key)
    Next row
End Sub

Public Function convertCodeToName(ByRef code As String) As String
    ' Handle "--No Set--"
    If code = vbNullString Then
        Let convertCodeToName = getNoSet()
        Exit Function
    End If
    Dim modifier As String
    ' Handle Ctrl modifier
    If left$(code, LEN_MODIFIER_CODE) = CTRL_CODE Then
        Let modifier = CTRL_KEY
        Let code = mid$(code, LEN_MODIFIER_CODE + 1)
    End If
    ' Handle Shift modifier
    If left$(code, LEN_MODIFIER_CODE) = SHIFT_CODE Then
        Let modifier = modifier & SHIFT_KEY
        Let code = mid$(code, LEN_MODIFIER_CODE + 1)
    End If
    ' Handle Alt modifier
    If left$(code, LEN_MODIFIER_CODE) = ALT_CODE Then
        Let modifier = modifier & ALT_KEY
        Let code = mid$(code, LEN_MODIFIER_CODE + 1)
    End If
    ' Map the remaining code to key name
    Select Case code
        Case ENTER_CODE: Let code = ENTER_KEY
        Case BACKSPACE_CODE: Let code = BACKSPACE_KEY
        Case BREAK_CODE: Let code = BREAK_KEY
        Case CLEAR_CODE: Let code = CLEAR_KEY
        Case DELETE_CODE: Let code = DELETE_KEY
        Case END_CODE: Let code = END_KEY
        Case ENTER_NUMPAD_CODE: Let code = ENTER_NUMPAD_KEY
        Case ESCAPE_CODE: Let code = ESCAPE_KEY
        Case HELP_CODE: Let code = HELP_KEY
        Case HOME_CODE: Let code = HOME_KEY
        Case INS_CODE: Let code = INS_KEY
        Case RETURN_CODE: Let code = RETURN_KEY
        Case CAPSLOCK_CODE: Let code = CAPSLOCK_KEY
        Case NUMLOCK_CODE: Let code = NUMLOCK_KEY
        Case SCROLLLOCK_CODE: Let code = SCROLLLOCK_KEY
        Case TAB_CODE: Let code = TAB_KEY
        Case UP_ARROW_CODE: Let code = UP_ARROW_KEY
        Case DOWN_ARROW_CODE: Let code = DOWN_ARROW_KEY
        Case LEFT_ARROW_CODE: Let code = LEFT_ARROW_KEY
        Case RIGHT_ARROW_CODE: Let code = RIGHT_ARROW_KEY
        Case PAGE_DOWN_CODE: Let code = PAGE_DOWN_KEY
        Case PAGE_UP_CODE: Let code = PAGE_UP_KEY
    End Select
    Let code = Replace(code, OPEN_CURLY_BRACKET, vbNullString)
    Let code = Replace(code, CLOSE_CURLY_BRACKET, vbNullString)
    Let convertCodeToName = modifier & IIf(code = vbNullString, getNoSet(), code)
End Function

Public Function convertKeyToName(ByRef keyCode As MsForms.ReturnInteger, ByRef Shift As Integer) As String
    Dim modifier As String
    Dim key As String
    'Listen modifier
    Select Case Shift
        Case (MASK.shiftKey): Let modifier = SHIFT_KEY
        Case (MASK.ctrlKey): Let modifier = CTRL_KEY
        Case (MASK.altKey): Let modifier = ALT_KEY
        Case (MASK.ctrlKey + MASK.shiftKey): Let modifier = CTRL_SHIFT_KEY
        Case (MASK.ctrlKey + MASK.altKey): Let modifier = CTRL_ALT_KEY
        Case (MASK.ctrlKey + MASK.shiftKey + MASK.altKey): Let modifier = CTRL_SHIFT_ALT_KEY
        Case (MASK.shiftKey + MASK.altKey): Let modifier = SHIFT_ALT_KEY
    End Select
    'Listen main key
    Select Case keyCode
        Case vbKeyF1 To vbKeyF16: key = FUNCTION_KEY & (keyCode - vbKeyF1 + 1) 
        Case vbKeyA To vbKeyZ: key = Chr(keyCode)
        Case vbKey0 To vbKey9: key = Chr(keyCode)
        Case vbKeyNumpad0 To vbKeyNumpad9: key = Chr(keyCode) & NUMPAD_MARK
        Case vbKeyBack: key = BACKSPACE_KEY
        Case vbKeyDelete: key = DELETE_KEY
        Case vbKeyTab: key = TAB_KEY
        Case vbKeyReturn: key = ENTER_KEY 'vbKeyReturn is Enter not Backspace, Backspace key is vbKeyBack
        Case vbKeyEscape: key = ESCAPE_KEY
        Case vbKeyControl: key = vbNullString
        Case vbKeyShift: key = vbNullString
        Case vbKeyMenu: key = vbNullString 'vbKeyMenu is alt not menu, menu key is 93(0x5D)
        Case Else: key = "Unknown(" & keyCode & ")"
    End Select
    'Skip modifier only
    If key = vbNullString Then
        Let convertKeyToName = vbNullString
    Else
        Let convertKeyToName = modifier & key
    End If
End Function

' NOTE: Use ByVal if reuse parameter as variable
Public Function convertNameToCode(ByVal name As String) As String
    Dim modifier As String
    ' Handle "--No Set--"
    If name = getNoSet() Then
        Let convertNameToCode = vbNullString
        Exit Function
    End If
    ' Handle Ctrl modifier
    If left$(name, Len(CTRL_KEY)) = CTRL_KEY Then
        Let modifier = CTRL_CODE
        Let name = mid$(name, Len(CTRL_KEY) + 1)
    End If
    ' Handle Shift modifier
    If left$(name, Len(SHIFT_KEY)) = SHIFT_KEY Then
        Let modifier = modifier & SHIFT_CODE
        Let name = mid$(name, Len(SHIFT_KEY) + 1)
    End If
    ' Handle Alt modifier
    If left$(name, Len(ALT_KEY)) = ALT_KEY Then
        Let modifier = modifier & ALT_CODE
        Let name = mid$(name, Len(ALT_KEY) + 1)
    End If
    ' Map the remaining key name back to its code
    Select Case name
        Case ENTER_KEY: Let name = ENTER_CODE
        Case BACKSPACE_KEY: Let name = BACKSPACE_CODE
        Case BREAK_KEY: Let name = BREAK_CODE
        Case CLEAR_KEY: Let name = CLEAR_CODE
        Case DELETE_KEY: Let name = DELETE_CODE
        Case END_KEY: Let name = END_CODE
        Case ENTER_NUMPAD_KEY: Let name = ENTER_NUMPAD_CODE
        Case ESCAPE_KEY: Let name = ESCAPE_CODE
        Case HELP_KEY: Let name = HELP_CODE
        Case HOME_KEY: Let name = HOME_CODE
        Case INS_KEY: Let name = INS_CODE
        Case RETURN_KEY: Let name = RETURN_CODE
        Case CAPSLOCK_KEY: Let name = CAPSLOCK_CODE
        Case NUMLOCK_KEY: Let name = NUMLOCK_CODE
        Case SCROLLLOCK_KEY: Let name = SCROLLLOCK_CODE
        Case TAB_KEY: Let name = TAB_CODE
        Case UP_ARROW_KEY: Let name = UP_ARROW_CODE
        Case DOWN_ARROW_KEY: Let name = DOWN_ARROW_CODE
        Case LEFT_ARROW_KEY: Let name = LEFT_ARROW_CODE
        Case RIGHT_ARROW_KEY: Let name = RIGHT_ARROW_CODE
        Case PAGE_DOWN_KEY: Let name = PAGE_DOWN_CODE
        Case PAGE_UP_KEY: Let name = PAGE_UP_CODE
        Case Else: Let name = OPEN_CURLY_BRACKET & name & CLOSE_CURLY_BRACKET
    End Select
    Let convertNameToCode = modifier & name
End Function
