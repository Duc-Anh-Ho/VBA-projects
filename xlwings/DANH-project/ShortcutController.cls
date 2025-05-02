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
Private Const DOT as String = "."
Private Const SHORTCUT_MODULE_NAME As String = "Shortcuts"
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

Enum KEY_CODE
    None = 0
    ' Mouse with VBA constants (1~6)
    LeftMouse = vbKeyLButton       ' 1
    RightMouse = vbKeyRButton      ' 2
    MiddleMouse = vbKeyMButton     ' 4
    XButton1 = 5
    XButton2 = 6
    ' Numbers (48~57 - D0~D9)
    Zero  = vbKey0: One  = vbKey1: Two = vbKey2: Three = vbKey3
    Four  = vbKey4: Five = vbKey5: Six = vbKey6: Seven = vbKey7
    Eight = vbKey8: Nine = vbKey9
    ' Letters (65~90)
    A = vbKeyA: B = vbKeyB: C = vbKeyC: D = vbKeyD: E = vbKeyE
    F = vbKeyF: G = vbKeyG: H = vbKeyH: I = vbKeyI: J = vbKeyJ
    K = vbKeyK: L = vbKeyL: M = vbKeyM: N = vbKeyN: O = vbKeyO
    P = vbKeyP: Q = vbKeyQ: R = vbKeyR: S = vbKeyS: T = vbKeyT
    U = vbKeyU: V = vbKeyV: W = vbKeyW: X = vbKeyX: Y = vbKeyY
    Z = vbKeyZ
    ' Numpad (96~111)
    ZeroNum      = vbKeyNumpad0  : OneNum      = vbKeyNumpad1
    TwoNum       = vbKeyNumpad2  : ThreeNum    = vbKeyNumpad3
    FourNum      = vbKeyNumpad4  : FiveNum     = vbKeyNumpad5
    SixNum       = vbKeyNumpad6  : SevenNum    = vbKeyNumpad7
    EightNum     = vbKeyNumpad8  : NineNum     = vbKeyNumpad9
    MultiplyNum  = vbKeyMultiply : PlusNum     = vbKeyAdd
    SeparatorNum = vbKeySeparator: SubtractNum = vbKeySubtract
    DecimalNum   = vbKeyDecimal  : DivideNum   = vbKeyDivide
    LockNum      = vbKeyNumlock
    ' Function keys (112~127)
    FOne       = vbKeyF1 : FTwo      = vbKeyF2 : FThree   = vbKeyF3
    FFour      = vbKeyF4 : FFive     = vbKeyF5 : FSix     = vbKeyF6
    FSeven     = vbKeyF7 : FEight    = vbKeyF8 : FNine    = vbKeyF9
    FTen       = vbKeyF10: FEleven   = vbKeyF11: FTwelve  = vbKeyF12
    FThirteen  = vbKeyF13: FFourteen = vbKeyF14: FFifteen = vbKeyF15
    FSixteen   = vbKeyF16
    FSeventeen   = 128: FEighteen   = 129: FNineteen  = 130
    FTwenty      = 131: FTwentyOne  = 132: FTwentyTwo = 133
    FTwentyThree = 134: FTwentyFour = 135
    ' Control keys with VBA constants
    CancelKey   = vbKeyCancel   ' 3  - Can conflict with build-in
    Backspace   = vbKeyBack     ' 8
    TabKey      = vbKeyTab      ' 9  - Can conflict with build-in
    ClearKey    = vbKeyClear    ' 12 - Can conflict with build-in
    EnterKey    = vbKeyReturn   ' 13 - Can conflict with build-in
    Shift       = vbKeyShift    ' 16
    Ctrl        = vbKeyControl  ' 17
    Alt         = vbKeyMenu     ' 18
    Pause       = vbKeyPause    ' 19
    CapsLock    = vbKeyCapital  ' 20
    Escape      = vbKeyEscape   ' 27
    SpaceKey    = vbKeySpace    ' 32 - Can conflict with build-in
    PageUp      = vbKeyPageUp   ' 33 ' Prior
    PageDown    = vbKeyPageDown ' 34 ' Next
    EndKey      = vbKeyEnd      ' 35 - Can conflict with build-in
    Home        = vbKeyHome     ' 36
    LeftArrow   = vbKeyLeft     ' 37
    UpArrow     = vbKeyUp       ' 38
    RightArrow  = vbKeyRight    ' 39
    DownArrow   = vbKeyDown     ' 40
    SelectKey   = vbKeySelect   ' 41 - Can conflict with build-in
    PrintKey    = vbKeyPrint    ' 42 - Can conflict with build-in
    Execute     = vbKeyExecute  ' 43
    PrintScreen = vbKeySnapshot ' 44
    Insert      = vbKeyInsert   ' 45
    Delete      = vbKeyDelete   ' 46
    Help        = vbKeyHelp     ' 47
    ' Else (System, Symbol, Browser, Media ...)
    LineFeed            = 10
    KanaMode            = 21 ' = HangulMode = HanguelMode
    JunjaMode           = 23
    FinalMode           = 24
    KanjiMode           = 25 ' = HanjaMode
    IMEConvert          = 28
    IMENonconvert       = 29
    IMEAccept           = 30 ' = IMEAceept
    IMEModeChange       = 31
    LeftWin             = 91
    RightWin            = 92
    MenuKey             = 93 ' Can conflict with build-in
    Sleep               = 95
    NumLock             = 144
    ScrollLock          = 145
    LeftShift           = 160
    RightShift          = 161
    LeftControl         = 162
    RightControl        = 163
    LeftMenu            = 164
    RightMenu           = 165
    BrowserBack         = 166
    BrowserForward      = 167
    BrowserRefresh      = 168
    BrowserStop         = 169
    BrowserSearch       = 170
    BrowserFavorites    = 171
    BrowserHome         = 172
    VolumeMute          = 173
    VolumeDown          = 174
    VolumeUp            = 175
    NextTrack           = 176
    PrevTrack           = 177
    StopMedia           = 178
    PlayPause           = 179
    StartMail           = 180
    SelectMedia         = 181
    StartApp1           = 182
    StartApp2           = 183
    Semicolon           = 186 ' Oem1
    Equal               = 187 ' Plus
    Comma               = 188
    Minus               = 189
    Period              = 190
    Slash               = 191 ' Oem2 - Question
    Tilde               = 192 ' Oem3
    OpenBracket         = 219 ' Oem4
    Pipe                = 220 ' Oem5
    CloseBracket        = 221 ' Oem6
    Quote               = 222 ' Oem7
    Power               = 223 ' Oem8
    Backslash           = 226 ' Oem102
    Wake                = 227
    ProcessKey          = 229
    Packet              = 231
    Reset               = 233
    Jump                = 234
    PlayStop            = 235
    Eject               = 236
    Attn                = 246
    Crsel               = 247
    Exsel               = 248
    EraseEof            = 249
    Play                = 250
    Zoom                = 251
    NoName              = 252
    Pa1                 = 253
    ClearEnd            = 254
End Enum

Private Const NO_SET As String = "--No Set--"
Private Const UNKNOWN As String = "Unknown"
Private Const OPEN_CURLY_BRACKET As String = "{"
Private Const CLOSE_CURLY_BRACKET As String = "}"
Private Const OPEN_ROUND_BRACKET As String = "("
Private Const CLOSE_ROUND_BRACKET As String = ")"
' CODE constant
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
Private Const ENTER_NUM_CODE As String = "{ENTER}"
Private Const ESCAPE_CODE As String = "{ESC}"
Private Const HELP_CODE As String = "{HELP}"
Private Const HOME_CODE As String = "{HOME}"
Private Const INSERT_CODE As String = "{INSERT}"
Private Const RETURN_CODE As String = "{RETURN}"
Private Const CAPS_LOCK_CODE As String = "{CAPSLOCK}"
Private Const NUM_LOCK_CODE As String = "{NUMLOCK}"
Private Const SCROLL_LOCK_CODE As String = "{SCROLLLOCK}"
Private Const TAB_CODE As String = "{TAB}"
Private Const UP_ARROW_CODE As String = "{UP}"
Private Const DOWN_ARROW_CODE As String = "{DOWN}"
Private Const LEFT_ARROW_CODE As String = "{LEFT}"
Private Const RIGHT_ARROW_CODE As String = "{RIGHT}"
Private Const PAGE_DOWN_CODE As String = "{PGDN}"
Private Const PAGE_UP_CODE As String = "{PGUP}"
Private Const COMMAND_CODE As String = "*" 'MAC below Excel 2011

' KEY Constant
Private Const ALT_KEY As String = "Alt"
Private Const ATTN_KEY As String = "Attn"
Private Const BACK_KEY As String = "Back"
Private Const BACKSPACE_KEY As String = "Backspace"
Private Const BREAK_KEY As String = "Break"
Private Const BROWSER_BACK_KEY As String = "Browser Back"
Private Const BROWSER_FAVORITES_KEY As String = "Browser Favorites"
Private Const BROWSER_FORWARD_KEY As String = "Browser Forward"
Private Const BROWSER_HOME_KEY As String = "Browser Home"
Private Const BROWSER_REFRESH_KEY As String = "Browser Refresh"
Private Const BROWSER_SEARCH_KEY As String = "Browser Search"
Private Const BROWSER_STOP_KEY As String = "Browser Stop"
Private Const CANCEL_KEY As String = "Cancel"
Private Const CAPITAL_KEY As String = "Capital"
Private Const CAPS_LOCK_KEY As String = "CapsLock"
Private Const CLEAR_KEY As String = "Clear"
Private Const COMMAND_KEY As String = "Command"
Private Const CRSEL_KEY As String = "Crsel"
Private Const CTRL_KEY As String = "Ctrl"
Private Const DELETE_KEY As String = "Delete"
Private Const DOT_KEY As String = "Dot(Num)"
Private Const DOWN_ARROW_KEY As String = "Down Arrow"
Private Const DOWN_KEY As String = "Down"
Private Const END_KEY As String = "End"
Private Const ENTER_KEY As String = "Enter"
Private Const ENTER_NUM_KEY As String = "Enter(Num)"
Private Const ERASE_EOF_KEY As String = "Erase Eof"
Private Const ESCAPE_KEY As String = "Escape"
Private Const EXECUTE_KEY As String = "Execute"
Private Const EXSEL_KEY As String = "Exsel"
Private Const F1_KEY As String = "F1"
Private Const F2_KEY As String = "F2"
Private Const F3_KEY As String = "F3"
Private Const F4_KEY As String = "F4"
Private Const F5_KEY As String = "F5"
Private Const F6_KEY As String = "F6"
Private Const F7_KEY As String = "F7"
Private Const F8_KEY As String = "F8"
Private Const F9_KEY As String = "F9"
Private Const F10_KEY As String = "F10"
Private Const F11_KEY As String = "F11"
Private Const F12_KEY As String = "F12"
Private Const F13_KEY As String = "F13"
Private Const F14_KEY As String = "F14"
Private Const F15_KEY As String = "F15"
Private Const F16_KEY As String = "F16"
Private Const F17_KEY As String = "F17"
Private Const F18_KEY As String = "F18"
Private Const F19_KEY As String = "F19"
Private Const F20_KEY As String = "F20"
Private Const F21_KEY As String = "F21"
Private Const F22_KEY As String = "F22"
Private Const F23_KEY As String = "F23"
Private Const F24_KEY As String = "F24"
Private Const FINAL_MODE_KEY As String = "Final Mode"
Private Const HELP_KEY As String = "Help"
Private Const HOME_KEY As String = "Home"
Private Const IME_ACCEPT_KEY As String = "IME Accept" ' "IME Aceept" MS typo
Private Const IME_CONVERT_KEY As String = "IME Convert"
Private Const IME_MODE_CHANGE_KEY As String = "IME ModeChange"
Private Const IME_NONCONVERT_KEY As String = "IME Nonconvert"
Private Const INSERT_KEY As String = "Insert"
Private Const JUNJA_MODE_KEY As String = "Junja Mode"
Private Const KANA_MODE_KEY As String = "Kana Mode"
Private Const KANJI_MODE_KEY As String = "Kanji Mode"
Private Const KEY_CODE_KEY As String = "Key Code"
Private Const START_APP_1_KEY As String = "Start App 1"
Private Const START_APP_2_KEY As String = "Start App 2"
Private Const START_MAIL_KEY As String = "Start Mail"
Private Const LEFT_BUTTON_KEY As String = "Left Button"
Private Const LEFT_CONTROL_KEY As String = "Left Ctrl"
Private Const LEFT_ARROW_KEY As String = "Left Arrow"
Private Const LINE_FEED_KEY As String = "Line Feed"
Private Const LEFT_MENU_KEY As String = "Left Menu"
Private Const LEFT_SHIFT_KEY As String = "Left Shift Key"
Private Const LEFT_WIN_KEY As String = "LeftWin"
Private Const MBUTTON_KEY As String = "MButton"
Private Const NEXT_TRACK_KEY As String = "Next Track"
Private Const PLAY_PAUSE_KEY As String = "Play/Pause"
Private Const PREVIOUS_TRACK_KEY As String = "Previous Track"
Private Const STOP_KEY As String = "Stop Media"
Private Const MENU_KEY As String = "Menu" ' Apps
Private Const NEXT_KEY As String = "Next"
Private Const NO_NAME_KEY As String = "No Name"
Private Const BACKSLASH_KEY As String = "Backslash" ' Oem102
Private Const CLEAR2_KEY As String = "Clear"
Private Const CLOSE_BRACKET_KEY As String = "Close Brackets" ' Oem6
Private Const MINUS_KEY As String = "Minus"
Private Const OPEN_BRACKET_KEY As String = "Open Brackets" ' Oem4
Private Const PERIOD_KEY As String = "Period"
Private Const PIPE_KEY As String = "Pipe"  ' Oem5
Private Const SLASH_KEY As String = "Slash" ' Oem2 - Question
Private Const QUOTE_KEY As String = "Quote" ' Oem7
Private Const POWER_KEY As String = "POWER" ' Oem8
Private Const SEMICOLON_KEY As String = "Semicolon" ' Oem1
Private Const COMMA_KEY As String = "Comma"
Private Const EQUAL_KEY As String = "Equal" ' Plus
Private Const TILDE_KEY As String = "Tilde" ' Oem3
Private Const PA1_KEY As String = "Pa1"
Private Const PACKET_KEY As String = "Packet"
Private Const PAGE_DOWN_KEY As String = "Page Down" ' Next
Private Const PAGE_UP_KEY As String = "Page Up" ' Prior
Private Const PAUSE_KEY As String = "Pause"
Private Const PLAY_KEY As String = "Play"
Private Const PRINT_KEY As String = "Print"
Private Const PRINT_SCREEN_KEY As String = "Print Screen"
Private Const PROCESS_KEY As String = "Process"
Private Const RIGHT_BUTTON_KEY As String = "Right Button"
Private Const RIGHT_CONTROL_KEY As String = "Right Ctrl"
Private Const RETURN_KEY As String = "Return"
Private Const RIGHT_ARROW_KEY As String = "Right Arrow"
Private Const RIGHT_KEY As String = "Right"
Private Const RIGHT_MENU_KEY As String = "RightMenu"
Private Const RIGHT_SHIFT_KEY As String = "Right Shift"
Private Const RIGHT_WIN_KEY As String = "Right Win"
Private Const SCROLL_LOCK_KEY As String = "Scrolllock"
Private Const SELECT_KEY As String = "Select"
Private Const SELECT_MEDIA_KEY As String = "Select Media"
Private Const SHIFT_KEY As String = "Shift"
Private Const SLEEP_KEY As String = "Sleep"
Private Const SCREENSHOT_KEY As String = "Screenshot" 'Snapshot
Private Const SPACE_KEY As String = "Space"
Private Const TAB_KEY As String = "Tab"
Private Const UP_ARROW_KEY As String = "Up Arrow"
Private Const VOLUME_DOWN_KEY As String = "Volume Down"
Private Const VOLUME_MUTE_KEY As String = "Volume Mute"
Private Const VOLUME_UP_KEY As String = "Volume Up"
Private Const XBUTTON1_KEY As String = "XButton1"
Private Const XBUTTON2_KEY As String = "XButton2"
Private Const ZOOM_KEY As String = "Zoom"

Private Const MULTIPLE_KEY As String = "Multiply(Num)"
Private Const PLUS_KEY As String = "Plus(Num)"
Private Const SEPARATOR_KEY As String = "Separator"
Private Const SUBTRACT_KEY As String = "Subtract(Num)"
Private Const DECIMAL_KEY As String = "Decimal"
Private Const DIVIDE_KEY As String = "Divide(Num)"
Private Const NUM_LOCK_KEY As String = "NumLock"

' MARK Constants
Private Const FUNCTION_MARK As String = "F"
Private Const NUM_MARK AS String = "(Num)"
Private Const PLUS As String = " + "
Private Const CTRL_PLUS As String = CTRL_KEY + PLUS
Private Const SHIFT_PLUS As String = SHIFT_KEY + PLUS
Private Const ALT_PLUS As String = ALT_KEY + PLUS
Private Const CTRL_SHIFT_PLUS As String = CTRL_PLUS & SHIFT_PLUS
Private Const CTRL_ALT_PLUS As String = CTRL_PLUS & ALT_PLUS
Private Const CTRL_SHIFT_ALT_PLUS As String = CTRL_SHIFT_PLUS & ALT_PLUS
Private Const SHIFT_ALT_PLUS As String = SHIFT_PLUS & ALT_PLUS

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
Public Property Get getUnknown() As String: Let getUnknown = UNKNOWN: End Property
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
    Application.OnKey key:=key, procedure:=SHORTCUT_MODULE_NAME & DOT & procedure
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
        Let modifier = CTRL_PLUS
        Let code = mid$(code, LEN_MODIFIER_CODE + 1)
    End If
    ' Handle Shift modifier
    If left$(code, LEN_MODIFIER_CODE) = SHIFT_CODE Then
        Let modifier = modifier & SHIFT_PLUS
        Let code = mid$(code, LEN_MODIFIER_CODE + 1)
    End If
    ' Handle Alt modifier
    If left$(code, LEN_MODIFIER_CODE) = ALT_CODE Then
        Let modifier = modifier & ALT_PLUS
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
        Case ENTER_NUM_CODE: Let code = ENTER_NUM_KEY
        Case ESCAPE_CODE: Let code = ESCAPE_KEY
        Case HELP_CODE: Let code = HELP_KEY
        Case HOME_CODE: Let code = HOME_KEY
        Case INSERT_CODE: Let code = INSERT_KEY
        Case RETURN_CODE: Let code = RETURN_KEY
        Case CAPS_LOCK_CODE: Let code = CAPS_LOCK_KEY
        Case NUM_LOCK_CODE: Let code = NUM_LOCK_KEY
        Case SCROLL_LOCK_CODE: Let code = SCROLL_LOCK_KEY
        Case TAB_CODE: Let code = TAB_KEY
        Case UP_ARROW_CODE: Let code = UP_ARROW_KEY
        Case DOWN_ARROW_CODE: Let code = DOWN_ARROW_KEY
        Case LEFT_ARROW_CODE: Let code = LEFT_ARROW_KEY
        Case RIGHT_ARROW_CODE: Let code = RIGHT_ARROW_KEY
        Case PAGE_DOWN_CODE: Let code = PAGE_DOWN_KEY
        Case PAGE_UP_CODE: Let code = PAGE_UP_KEY
    End Select
    ' Remove Brackets
    Let code = Replace(code, OPEN_CURLY_BRACKET, vbNullString)
    Let code = Replace(code, CLOSE_CURLY_BRACKET, vbNullString)
    Let convertCodeToName = modifier & IIf(code = vbNullString, getNoSet(), code)
End Function

Public Function convertKeyToName(ByRef keyCode As MsForms.ReturnInteger, ByRef Shift As Integer) As String
    Dim modifier As String
    Dim key As String
    'Listen modifier
    Select Case Shift
        Case (MASK.shiftKey): Let modifier = SHIFT_PLUS
        Case (MASK.ctrlKey): Let modifier = CTRL_PLUS
        Case (MASK.altKey): Let modifier = ALT_PLUS
        Case (MASK.ctrlKey + MASK.shiftKey): Let modifier = CTRL_SHIFT_PLUS
        Case (MASK.ctrlKey + MASK.altKey): Let modifier = CTRL_ALT_PLUS
        Case (MASK.ctrlKey + MASK.shiftKey + MASK.altKey): Let modifier = CTRL_SHIFT_ALT_PLUS
        Case (MASK.shiftKey + MASK.altKey): Let modifier = SHIFT_ALT_PLUS
    End Select
    'Listen main key
    Select Case keyCode
        ' A~Z, F1~F24, 1~9, 1~9(num)
        Case KEY_CODE.A To KEY_CODE.Z: Let key = Chr(keyCode)
        Case KEY_CODE.FOne To KEY_CODE.FTwentyFour: Let key = FUNCTION_MARK & (keyCode - KEY_CODE.FOne + 1)
        Case KEY_CODE.Zero To KEY_CODE.Nine: Let key = Chr(keyCode)
        Case KEY_CODE.ZeroNum To KEY_CODE.NineNum: Let key = Chr(keyCode) & NUM_MARK

        ' TO BE CONTINUE
        Case vbKeyBack: Let key = BACKSPACE_KEY
        Case vbKeyDelete: Let key = DELETE_KEY
        Case vbKeyTab: Let key = TAB_KEY
        Case vbKeyReturn: Let key = ENTER_KEY 'vbKeyReturn is Enter not Backspace, Backspace key is vbKeyBack
        Case vbKeyEscape: Let key = ESCAPE_KEY
        Case vbKeyControl: Let key = vbNullString
        Case vbKeyShift: Let key = vbNullString
        Case vbKeyMenu: Let key = vbNullString 'vbKeyMenu is alt not menu, menu key is 93(0x5D)


        Case KEY_CODE.MultiplyNum: Let key = MULTIPLE_KEY
        Case KEY_CODE.PlusNum: Let key = PLUS_KEY
        Case KEY_CODE.SeparatorNum: Let key = SEPARATOR_KEY
        Case KEY_CODE.SubtractNum: Let key = SUBTRACT_KEY
        Case KEY_CODE.DecimalNum: Let key = DECIMAL_KEY
        Case KEY_CODE.DivideNum: Let key = DIVIDE_KEY
        Case KEY_CODE.NumLock: Let key = NUM_LOCK_KEY

        
        Case Else: Let key = getUnknown() & OPEN_ROUND_BRACKET & keyCode & CLOSE_ROUND_BRACKET ' Eg. Unknown(20)
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
    If left$(name, Len(CTRL_PLUS)) = CTRL_PLUS Then
        Let modifier = CTRL_CODE
        Let name = mid$(name, Len(CTRL_PLUS) + 1)
    End If
    ' Handle Shift modifier
    If left$(name, Len(SHIFT_PLUS)) = SHIFT_PLUS Then
        Let modifier = modifier & SHIFT_CODE
        Let name = mid$(name, Len(SHIFT_PLUS) + 1)
    End If
    ' Handle Alt modifier
    If left$(name, Len(ALT_PLUS)) = ALT_PLUS Then
        Let modifier = modifier & ALT_CODE
        Let name = mid$(name, Len(ALT_PLUS) + 1)
    End If
    ' Map the remaining key name back to its code
    Select Case name
        Case ENTER_KEY: Let name = ENTER_CODE
        Case BACKSPACE_KEY: Let name = BACKSPACE_CODE
        Case BREAK_KEY: Let name = BREAK_CODE
        Case CLEAR_KEY: Let name = CLEAR_CODE
        Case DELETE_KEY: Let name = DELETE_CODE
        Case END_KEY: Let name = END_CODE
        Case ENTER_NUM_KEY: Let name = ENTER_NUM_CODE
        Case ESCAPE_KEY: Let name = ESCAPE_CODE
        Case HELP_KEY: Let name = HELP_CODE
        Case HOME_KEY: Let name = HOME_CODE
        Case INSERT_KEY: Let name = INSERT_CODE
        Case RETURN_KEY: Let name = RETURN_CODE
        Case CAPS_LOCK_KEY: Let name = CAPS_LOCK_CODE
        Case NUM_LOCK_KEY: Let name = NUM_LOCK_CODE
        Case SCROLL_LOCK_KEY: Let name = SCROLL_LOCK_CODE
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
