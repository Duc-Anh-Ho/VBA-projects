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
Private Rows As ListRows
Private cols As ListColumns
Private Const DOT As String = "."
Private Const SHORTCUT_MODULE_NAME As String = "Shortcuts"
Private Const SHORTCUT_SHEET_NAME As String = "keyboard-shortcut"
Private Const SHORTCUT_TABLE_NAME As String = "Keybinding"

Private Const NO_SET As String = "--No Set--"
Private Const UNKNOWN As String = "Unknown"
Private Const ERR As String = "(Err)"
Private Const OPEN_CURLY_BRACKET As String = "{"
Private Const CLOSE_CURLY_BRACKET As String = "}"
Private Const OPEN_ROUND_BRACKET As String = "("
Private Const CLOSE_ROUND_BRACKET As String = ")"

Private Enum COL
    no = 1
    procedure = no + 1
    default_NAMEbind = procedure + 1
    custom_NAMEbind = default_NAMEbind + 1
    command = custom_NAMEbind + 1
    when = command + 1
    status = when + 1
End Enum

Private Enum MASK
    None = 0
    shiftKey = 1
    ctrlKey = 2
    altKey = 4
End Enum

Enum KEY_CODE
    Shift = vbKeyShift   ' 16
    ctrl = vbKeyControl  ' 17
    Alt = vbKeyMenu      ' 18
    ' Numbers (48~57 - D0~D9)
    ZERO = vbKey0: One = vbKey1: Two = vbKey2: Three = vbKey3
    Four = vbKey4: Five = vbKey5: Six = vbKey6: Seven = vbKey7
    Eight = vbKey8: Nine = vbKey9
    ' Letters (65~90)
    A = vbKeyA: B = vbKeyB: C = vbKeyC: D = vbKeyD: E = vbKeyE
    F = vbKeyF: G = vbKeyG: H = vbKeyH: i = vbKeyI: j = vbKeyJ
    k = vbKeyK: L = vbKeyL: m = vbKeyM: N = vbKeyN: o = vbKeyO
    p = vbKeyP: Q = vbKeyQ: r = vbKeyR: S = vbKeyS: T = vbKeyT
    U = vbKeyU: v = vbKeyV: W = vbKeyW: x = vbKeyX: y = vbKeyY
    Z = vbKeyZ
    ' Numpad (96~105)
    Numpad0 = vbKeyNumpad0: Numpad1 = vbKeyNumpad1
    Numpad2 = vbKeyNumpad2: Numpad3 = vbKeyNumpad3
    Numpad4 = vbKeyNumpad4: Numpad5 = vbKeyNumpad5
    Numpad6 = vbKeyNumpad6: Numpad7 = vbKeyNumpad7
    Numpad8 = vbKeyNumpad8: Numpad9 = vbKeyNumpad9
    ' Function keys (112~127)
    FOne = vbKeyF1: FTwo = vbKeyF2: FThree = vbKeyF3
    FFour = vbKeyF4: FFive = vbKeyF5: FSix = vbKeyF6
    FSeven = vbKeyF7: FEight = vbKeyF8: FNine = vbKeyF9
    FTen = vbKeyF10: FEleven = vbKeyF11: FTwelve = vbKeyF12
    FThirteen = vbKeyF13: FFourteen = vbKeyF14: FFifteen = vbKeyF15
    FSixteen = vbKeyF16
    ' Available keys
    Backspace = vbKeyBack             ' 8
    TabKey = vbKeyTab                 ' 9
    ClearKey = vbKeyClear             ' 12
    EnterKey = vbKeyReturn            ' 13
    Break = vbKeyPause                ' 19
    CapsLock = vbKeyCapital           ' 20
    Escape = vbKeyEscape              ' 27
    SpaceKey = vbKeySpace             ' 32
    PageUp = vbKeyPageUp              ' 33
    PageDown = vbKeyPageDown          ' 34
    EndKey = vbKeyEnd                 ' 35
    Home = vbKeyHome                  ' 36
    LeftArrow = vbKeyLeft             ' 37
    UpArrow = vbKeyUp                 ' 38
    RightArrow = vbKeyRight           ' 39
    DownArrow = vbKeyDown             ' 40
    Insert = vbKeyInsert              ' 45
    Delete = vbKeyDelete              ' 46
    Help = vbKeyHelp                  ' 47
    Multiply = vbKeyMultiply          ' 106
    AddKey = vbKeyAdd                 ' 107
    EnterNum = vbKeySeparator         ' 108
    Subtract = vbKeySubtract          ' 109
    DecimalKey = vbKeyDecimal         ' 110
    Divide = vbKeyDivide              ' 111
    NumLock = vbKeyNumlock            ' 144
    ScrollLock = 145

    Semicolon = 186
    Equals = 187
    Comma = 188
    Hyphen = 189
    Period = 190
    ForwardSlash = 191
    GraveAccent = 192        ' Tilde
    OpenBracket = 219
    Backslash = 220          ' Pipe
    CloseBracket = 221
    Apostrophe = 222         ' Quote
    Slash = 226

    ' Non-test

    FSeventeen = 128: FEighteen = 129: FNineteen = 130
    FTwenty = 131: FTwentyOne = 132: FTwentyTwo = 133
    FTwentyThree = 134: FTwentyFour = 135

    None = 0

    ' Mouse with VBA constants (1~6)
    LeftMouse = vbKeyLButton         ' 1
    RightMouse = vbKeyRButton        ' 2
    CancelKey = vbKeyCancel          ' 3
    MiddleMouse = vbKeyMButton       ' 4
    XButton1 = 5
    XButton2 = 6

    ' Other keys
    LineFeed = 10
    KanaMode = 21     ' = HangulMode
    JunjaMode = 23
    FinalMode = 24
    KanjiMode = 25    ' = HanjaMode
    IMEConvert = 28
    IMENonconvert = 29
    IMEAccept = 30    ' = IMEAceept
    IMEModeChange = 31
    SelectKey = vbKeySelect          ' 41
    PrintKey = vbKeyPrint            ' 42
    Execute = vbKeyExecute           ' 43
    PrintScreen = vbKeySnapshot      ' 44

    ' Windows keys
    LeftWin = 91
    RightWin = 92
    MenuKey = 93      ' Can conflict with build-in

    ' Other special keys
    Sleep = 95
    LeftShift = 160
    RightShift = 161
    LeftControl = 162
    RightControl = 163
    LeftMenu = 164
    RightMenu = 165

    ' Browser keys
    BrowserBack = 166
    BrowserForward = 167
    BrowserRefresh = 168
    BrowserStop = 169
    BrowserSearch = 170
    BrowserFavorites = 171
    BrowserHome = 172

    ' Media keys
    VolumeMute = 173
    VolumeDown = 174
    VolumeUp = 175
    NextTrack = 176
    PrevTrack = 177
    StopMedia = 178
    PlayPause = 179

    ' Application keys
    StartMail = 180
    SelectMedia = 181
    StartApp1 = 182
    StartApp2 = 183

    ' Other characters and special keys
    Power = 223
    Wake = 227
    ProcessKey = 229
    Packet = 231
    Reset = 233
    Jump = 234
    PlayStop = 235
    Eject = 236
    Attn = 246
    Crsel = 247
    Exsel = 248
    EraseEof = 249
    Play = 250
    Zoom = 251
    NoName = 252
    Pa1 = 253
    ClearEnd = 254
End Enum

' CODE constant
Private Const LEN_MODIFIER_CODE As Byte = 1
Private Const CTRL_CODE As String = "^"
Private Const SHIFT_CODE As String = "+"
Private Const ALT_CODE As String = "%"

Private Const BACKSPACE_CODE        As String = "{BS}" ' {BACKSPACE}
Private Const TAB_CODE              As String = "{TAB}"
Private Const CLEAR_CODE            As String = "{CLEAR}"
Private Const ENTER_CODE            As String = "~"
Private Const RETURN_CODE           As String = "{RETURN}"  ' Enter in linux/mac
Private Const BREAK_CODE            As String = "{BREAK}"
Private Const CAPS_LOCK_CODE        As String = "{CAPSLOCK}"
Private Const ESCAPE_CODE           As String = "{ESC}"
Private Const SPACE_CODE            As String = " "
Private Const PAGE_UP_CODE          As String = "{PGUP}"
Private Const PAGE_DOWN_CODE        As String = "{PGDN}"
Private Const END_CODE              As String = "{END}"
Private Const HOME_CODE             As String = "{HOME}"
Private Const LEFT_ARROW_CODE       As String = "{LEFT}"
Private Const UP_ARROW_CODE         As String = "{UP}"
Private Const RIGHT_ARROW_CODE      As String = "{RIGHT}"
Private Const DOWN_ARROW_CODE       As String = "{DOWN}"
Private Const INSERT_CODE           As String = "{INSERT}"
Private Const DELETE_CODE           As String = "{DEL}"
Private Const HELP_CODE             As String = "{HELP}"
Private Const MULTIPLY_CODE         As String = "{*}"       ' Multiply
Private Const ADD_CODE              As String = "{+}"       ' Plus
Private Const ENTER_NUM_CODE        As String = "{ENTER}"
Private Const SUBTRACT_CODE         As String = "{-}"       ' Minus
Private Const DECIMAL_CODE          As String = "{.}"       ' Decimal
Private Const DIVIDE_CODE           As String = "{/}"       ' Divide
Private Const NUM_LOCK_CODE         As String = "{NUMLOCK}"
Private Const SCROLL_LOCK_CODE      As String = "{SCROLLLOCK}"
' Private Const COMMAND_CODE          As String = "*" 'MAC below Excel 2011

Private Const NUMPAD0_CODE          As String = "{96}"
Private Const NUMPAD1_CODE          As String = "{97}"
Private Const NUMPAD2_CODE          As String = "{98}"
Private Const NUMPAD3_CODE          As String = "{99}"
Private Const NUMPAD4_CODE          As String = "{100}"
Private Const NUMPAD5_CODE          As String = "{101}"
Private Const NUMPAD6_CODE          As String = "{102}"
Private Const NUMPAD7_CODE          As String = "{103}"
Private Const NUMPAD8_CODE          As String = "{104}"
Private Const NUMPAD9_CODE          As String = "{105}"

Private Const HYPHEN_CODE           As String = "-"
Private Const EQUALS_CODE           As String = "="
Private Const OPEN_BRACKET_CODE     As String = "["
Private Const CLOSE_BRACKET_CODE    As String = "]"
Private Const BACKSLASH_CODE        As String = "\"
Private Const SLASH_CODE            As String = "\"
Private Const SEMICOLON_CODE        As String = ";"
Private Const APOSTROPHE_CODE       As String = "'"
Private Const COMMA_CODE            As String = ","
Private Const PERIOD_CODE           As String = "."
Private Const FORWARD_SLASH_CODE    As String = "/"
Private Const GRAVE_ACCENT_CODE     As String = "`"

' KEY Constant
Private Const ALT_NAME              As String = "Alt"
Private Const CTRL_NAME             As String = "Ctrl"
Private Const SHIFT_NAME            As String = "Shift"

Private Const BACKSPACE_NAME        As String = "Backspace"
Private Const TAB_NAME              As String = "Tab"
Private Const CLEAR_NAME            As String = "Clear"
Private Const ENTER_NAME            As String = "Enter"
Private Const BREAK_NAME            As String = "Break"
Private Const CAPS_LOCK_NAME        As String = "CapsLock"
Private Const ESCAPE_NAME           As String = "Escape"
Private Const SPACE_NAME            As String = "Space"
Private Const PAGE_UP_NAME          As String = "Page Up"
Private Const PAGE_DOWN_NAME        As String = "Page Down"
Private Const END_NAME              As String = "End"
Private Const HOME_NAME             As String = "Home"
Private Const LEFT_ARROW_NAME       As String = "Left"
Private Const UP_ARROW_NAME         As String = "Up"
Private Const RIGHT_ARROW_NAME      As String = "Right"
Private Const DOWN_ARROW_NAME       As String = "Down"
Private Const INSERT_NAME           As String = "Insert"
Private Const DELETE_NAME           As String = "Delete"
Private Const HELP_NAME             As String = "Help"
Private Const NUMPAD0_NAME          As String = "0(num)"
Private Const NUMPAD1_NAME          As String = "1(num)"
Private Const NUMPAD2_NAME          As String = "2(num)"
Private Const NUMPAD3_NAME          As String = "3(num)"
Private Const NUMPAD4_NAME          As String = "4(num)"
Private Const NUMPAD5_NAME          As String = "5(num)"
Private Const NUMPAD7_NAME          As String = "7(num)"
Private Const NUMPAD6_NAME          As String = "6(num)"
Private Const NUMPAD8_NAME          As String = "8(num)"
Private Const NUMPAD9_NAME          As String = "9(num)"
Private Const MULTIPLY_NAME         As String = "Multiply(num)"
Private Const ADD_NAME              As String = "Plus(num)"
Private Const ENTER_NUM_NAME        As String = "Enter(num)"
Private Const SUBTRACT_NAME         As String = "Subtract(num)"
Private Const DECIMAL_NAME          As String = "Decimal(num)"
Private Const DIVIDE_NAME           As String = "Divide(num)"
Private Const NUM_LOCK_NAME         As String = "NumLock"
Private Const SCROLL_LOCK_NAME      As String = "ScrollLock"

Private Const HYPHEN_NAME           As String = "-"
Private Const EQUALS_NAME           As String = "="
Private Const OPEN_BRACKET_NAME     As String = "["
Private Const CLOSE_BRACKET_NAME    As String = "]"
Private Const BACKSLASH_NAME        As String = "\"
Private Const SEMICOLON_NAME        As String = ";"
Private Const APOSTROPHE_NAME       As String = "'"
Private Const COMMA_NAME            As String = ","
Private Const PERIOD_NAME           As String = "."
Private Const FORWARD_SLASH_NAME    As String = "/"
Private Const SLASH_NAME            As String = "/"
Private Const GRAVE_ACCENT_NAME     As String = "`"

Private Const ATTN_NAME              As String = "Attn"
Private Const BACK_NAME              As String = "Back"
Private Const BROWSER_BACK_NAME      As String = "Browser Back"
Private Const BROWSER_FAVORITES_NAME As String = "Browser Favorites"
Private Const BROWSER_FORWARD_NAME   As String = "Browser Forward"
Private Const BROWSER_HOME_NAME      As String = "Browser Home"
Private Const BROWSER_REFRESH_NAME   As String = "Browser Refresh"
Private Const BROWSER_SEARCH_NAME    As String = "Browser Search"
Private Const BROWSER_STOP_NAME      As String = "Browser Stop"
Private Const CANCEL_NAME            As String = "Cancel"
Private Const CAPITAL_NAME           As String = "Capital"
Private Const COMMAND_NAME           As String = "Command"
Private Const CRSEL_NAME             As String = "Crsel"
Private Const DOWN_NAME              As String = "Down"
Private Const ERASE_EOF_NAME         As String = "Erase Eof"
Private Const EXECUTE_NAME           As String = "Execute"
Private Const EXSEL_NAME             As String = "Exsel"
Private Const EJECT_NAME             As String = "Eject"
Private Const FINAL_MODE_NAME        As String = "Final Mode"
Private Const IME_ACCEPT_NAME        As String = "IME Accept" ' "IME Aceept" MS typo
Private Const IME_CONVERT_NAME       As String = "IME Convert"
Private Const IME_MODE_CHANGE_NAME   As String = "IME ModeChange"
Private Const IME_NONCONVERT_NAME    As String = "IME Nonconvert"
Private Const JUMP_NAME              As String = "Jump"
Private Const JUNJA_MODE_NAME        As String = "Junja Mode"
Private Const KANA_MODE_NAME         As String = "Kana Mode"
Private Const KANJI_MODE_NAME        As String = "Kanji Mode"
Private Const KEY_CODE_NAME          As String = "Key Code"
Private Const START_APP_1_NAME       As String = "Start App 1"
Private Const START_APP_2_NAME       As String = "Start App 2"
Private Const START_MAIL_NAME        As String = "Start Mail"
Private Const LEFT_BUTTON_NAME       As String = "Left Button"
Private Const LEFT_CONTROL_NAME      As String = "Left Ctrl"
Private Const LINE_FEED_NAME         As String = "Line Feed"
Private Const LEFT_MENU_NAME         As String = "Left Menu"
Private Const LEFT_SHIFT_NAME        As String = "Left Shift Key"
Private Const LEFT_WIN_NAME          As String = "LeftWin"
Private Const MBUTTON_NAME           As String = "MButton"
Private Const NEXT_TRACK_NAME        As String = "Next Track"
Private Const PLAY_PAUSE_NAME        As String = "Play/Pause"
Private Const PLAY_STOP_NAME         As String = "Play/Stop"
Private Const PREVIOUS_TRACK_NAME    As String = "Previous Track"
Private Const STOP_NAME              As String = "Stop Media"
Private Const MENU_NAME              As String = "Menu" ' Apps
Private Const NEXT_NAME              As String = "Next"
Private Const NO_NAME_NAME           As String = "No Name"
Private Const RESET_NAME             As String = "Reset"
Private Const POWER_NAME             As String = "Power"
Private Const EQUAL_NAME             As String = "Equal"
Private Const PA1_NAME               As String = "Pa1"
Private Const PACKET_NAME            As String = "Packet"
Private Const PAUSE_NAME             As String = "Pause"
Private Const PLAY_NAME              As String = "Play"
Private Const PRINT_NAME             As String = "Print"
Private Const PRINT_SCREEN_NAME      As String = "Print Screen"
Private Const PROCESS_NAME           As String = "Process"
Private Const RIGHT_BUTTON_NAME      As String = "Right Button"
Private Const RIGHT_CONTROL_NAME     As String = "Right Ctrl"
Private Const RETURN_NAME            As String = "Return"
Private Const RIGHT_NAME             As String = "Right"
Private Const RIGHT_MENU_NAME        As String = "RightMenu"
Private Const RIGHT_SHIFT_NAME       As String = "Right Shift"
Private Const RIGHT_WIN_NAME         As String = "Right Win"
Private Const SELECT_NAME            As String = "Select"
Private Const SELECT_MEDIA_NAME      As String = "Select Media"
Private Const SLEEP_NAME             As String = "Sleep"
Private Const SCREENSHOT_NAME        As String = "Screenshot" 'Snapshot
Private Const VOLUME_DOWN_NAME       As String = "Volume Down"
Private Const VOLUME_MUTE_NAME       As String = "Volume Mute"
Private Const VOLUME_UP_NAME         As String = "Volume Up"
Private Const XBUTTON1_NAME          As String = "XButton1"
Private Const XBUTTON2_NAME          As String = "XButton2"
Private Const WAKE_NAME              As String = "Wake"
Private Const ZOOM_NAME              As String = "Zoom"

' MARK Constants
Private Const FUNCTION_MARK         As String = "F"
Private Const PLUS                  As String = " + "
Private Const CTRL_PLUS             As String = CTRL_NAME + PLUS
Private Const SHIFT_PLUS            As String = SHIFT_NAME + PLUS
Private Const ALT_PLUS              As String = ALT_NAME + PLUS
Private Const CTRL_SHIFT_PLUS       As String = CTRL_PLUS & SHIFT_PLUS
Private Const CTRL_ALT_PLUS         As String = CTRL_PLUS & ALT_PLUS
Private Const CTRL_SHIFT_ALT_PLUS   As String = CTRL_SHIFT_PLUS & ALT_PLUS
Private Const SHIFT_ALT_PLUS        As String = SHIFT_PLUS & ALT_PLUS

' Loop iterators
Private Row As ListRow

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
Public Property Get getErr() As String: Let getErr = ERR: End Property
Public Property Get getNoCol() As Byte: Let getNoCol = COL.no: End Property
Public Property Get getProcedureCol() As Byte: Let getProcedureCol = COL.procedure: End Property
Public Property Get getDefaultKeybindingCol() As Byte: Let getDefaultKeybindingCol = COL.default_NAMEbind: End Property
Public Property Get getCustomKeybindingCol() As Byte: Let getCustomKeybindingCol = COL.custom_NAMEbind: End Property
Public Property Get getCommandCol() As Byte: Let getCommandCol = COL.command: End Property
Public Property Get getWhenCol() As Byte: Let getWhenCol = COL.when: End Property
Public Property Get getStatusCol() As Byte: Let getStatusCol = COL.status: End Property
Public Property Get getEnterKey() As String: Let getEnterKey = ENTER_NAME: End Property
Public Property Get getEscKey() As String: Let getEscKey = ESCAPE_NAME: End Property
Public Property Get getBackspaceKey() As String: Let getBackspaceKey = BACKSPACE_NAME: End Property
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
    For Each Row In shortcutTb.ListRows
        Let procedure = Row.Range(1, getProcedureCol()).value
        Let key = Row.Range(1, getCustomKeybindingCol()).value
        ' Skip Blank Keybinding
        If key <> vbNullString Then
            Call restore(key)
            Call add(key, procedure)
        End If
    Next Row
End Sub

Public Sub unInstall()
    Dim key As String
    For Each Row In shortcutTb.ListRows
        Let key = Row.Range(1, getCustomKeybindingCol()).value
        If key <> vbNullString Then Call restore(key)
    Next Row
End Sub

Public Function convertKeyToName(ByRef keyCode As MsForms.ReturnInteger, ByRef Shift As Integer) As String
    Dim modifier As String
    Dim key As String
    'Listen modifier
    Select Case Shift
        Case (MASK.shiftKey): Let modifier = SHIFT_PLUS
        Case (MASK.ctrlKey):  Let modifier = CTRL_PLUS
        Case (MASK.altKey):   Let modifier = ALT_PLUS
        Case (MASK.ctrlKey + MASK.shiftKey):  Let modifier = CTRL_SHIFT_PLUS
        Case (MASK.ctrlKey + MASK.altKey):    Let modifier = CTRL_ALT_PLUS
        Case (MASK.ctrlKey + MASK.shiftKey + MASK.altKey): Let modifier = CTRL_SHIFT_ALT_PLUS
        Case (MASK.shiftKey + MASK.altKey):   Let modifier = SHIFT_ALT_PLUS
    End Select
    'Listen main key
    Select Case keyCode
        'Skip modifier
        Case KEY_CODE.ctrl:   Let key = vbNullString
        Case KEY_CODE.Shift:  Let key = vbNullString
        Case KEY_CODE.Alt:    Let key = vbNullString
        ' A~Z, F1~F24, 1~9, 1~9(num)
        Case KEY_CODE.A To KEY_CODE.Z:              Let key = Chr(keyCode)
        Case KEY_CODE.FOne To KEY_CODE.FFifteen:    Let key = FUNCTION_MARK & (keyCode - KEY_CODE.FOne + 1)
        Case KEY_CODE.ZERO To KEY_CODE.Nine:        Let key = Chr(keyCode)
        Case KEY_CODE.Numpad0:      Let key = NUMPAD0_NAME
        Case KEY_CODE.Numpad1:      Let key = NUMPAD1_NAME
        Case KEY_CODE.Numpad2:      Let key = NUMPAD2_NAME
        Case KEY_CODE.Numpad3:      Let key = NUMPAD3_NAME
        Case KEY_CODE.Numpad4:      Let key = NUMPAD4_NAME
        Case KEY_CODE.Numpad5:      Let key = NUMPAD5_NAME
        Case KEY_CODE.Numpad6:      Let key = NUMPAD7_NAME
        Case KEY_CODE.Numpad7:      Let key = NUMPAD6_NAME
        Case KEY_CODE.Numpad8:      Let key = NUMPAD8_NAME
        Case KEY_CODE.Numpad9:      Let key = NUMPAD9_NAME
        ' Follow documentation
        Case KEY_CODE.Backspace:    Let key = BACKSPACE_NAME
        Case KEY_CODE.TabKey:       Let key = TAB_NAME
        Case KEY_CODE.ClearKey:     Let key = CLEAR_NAME
        Case KEY_CODE.EnterKey:     Let key = ENTER_NAME
        Case KEY_CODE.Break:        Let key = BREAK_NAME
        Case KEY_CODE.CapsLock:     Let key = CAPS_LOCK_NAME
        Case KEY_CODE.Escape:       Let key = ESCAPE_NAME
        Case KEY_CODE.SpaceKey:     Let key = SPACE_NAME
        Case KEY_CODE.PageUp:       Let key = PAGE_UP_NAME
        Case KEY_CODE.PageDown:     Let key = PAGE_DOWN_NAME
        Case KEY_CODE.EndKey:       Let key = END_NAME
        Case KEY_CODE.Home:         Let key = HOME_NAME
        Case KEY_CODE.LeftArrow:    Let key = LEFT_ARROW_NAME
        Case KEY_CODE.UpArrow:      Let key = UP_ARROW_NAME
        Case KEY_CODE.RightArrow:   Let key = RIGHT_ARROW_NAME
        Case KEY_CODE.DownArrow:    Let key = DOWN_ARROW_NAME
        Case KEY_CODE.Insert:       Let key = INSERT_NAME
        Case KEY_CODE.Delete:       Let key = DELETE_NAME
        Case KEY_CODE.Help:         Let key = HELP_NAME
        Case KEY_CODE.Multiply:     Let key = MULTIPLY_NAME
        Case KEY_CODE.AddKey:       Let key = ADD_NAME
        Case KEY_CODE.EnterNum:     Let key = ENTER_NUM_NAME
        Case KEY_CODE.Subtract:     Let key = SUBTRACT_NAME
        Case KEY_CODE.DecimalKey:   Let key = DECIMAL_NAME
        Case KEY_CODE.Divide:       Let key = DIVIDE_NAME
        Case KEY_CODE.NumLock:      Let key = NUM_LOCK_NAME
        Case KEY_CODE.ScrollLock:   Let key = SCROLL_LOCK_NAME

        Case KEY_CODE.Hyphen:       Let key = HYPHEN_NAME
        Case KEY_CODE.Equals:       Let key = EQUALS_NAME
        Case KEY_CODE.Comma:        Let key = COMMA_NAME
        Case KEY_CODE.OpenBracket:  Let key = OPEN_BRACKET_NAME
        Case KEY_CODE.CloseBracket: Let key = CLOSE_BRACKET_NAME
        Case KEY_CODE.Backslash:    Let key = BACKSLASH_NAME
        Case KEY_CODE.Semicolon:    Let key = SEMICOLON_NAME
        Case KEY_CODE.Apostrophe:   Let key = APOSTROPHE_NAME
        Case KEY_CODE.Period:       Let key = PERIOD_NAME
        Case KEY_CODE.ForwardSlash: Let key = FORWARD_SLASH_NAME
        Case KEY_CODE.GraveAccent:  Let key = GRAVE_ACCENT_NAME

        ' Non-test
        Case KEY_CODE.FSixteen To KEY_CODE.FTwentyFour: Let key = FUNCTION_MARK & (keyCode - KEY_CODE.FOne + 1)
        Case KEY_CODE.Attn:             Let key = ATTN_NAME
        Case KEY_CODE.Backslash:        Let key = BACKSLASH_NAME
        Case KEY_CODE.BrowserBack:      Let key = BROWSER_BACK_NAME
        Case KEY_CODE.BrowserFavorites: Let key = BROWSER_FAVORITES_NAME
        Case KEY_CODE.BrowserForward:   Let key = BROWSER_FORWARD_NAME
        Case KEY_CODE.BrowserHome:      Let key = BROWSER_HOME_NAME
        Case KEY_CODE.BrowserRefresh:   Let key = BROWSER_REFRESH_NAME
        Case KEY_CODE.BrowserSearch:    Let key = BROWSER_SEARCH_NAME
        Case KEY_CODE.BrowserStop:      Let key = BROWSER_STOP_NAME
        Case KEY_CODE.CancelKey:        Let key = CANCEL_NAME
        Case KEY_CODE.ClearEnd:         Let key = CLEAR_NAME
        Case KEY_CODE.Crsel:            Let key = CRSEL_NAME
        Case KEY_CODE.EraseEof:         Let key = ERASE_EOF_NAME
        Case KEY_CODE.Execute:          Let key = EXECUTE_NAME
        Case KEY_CODE.Exsel:            Let key = EXSEL_NAME
        Case KEY_CODE.Eject:            Let key = EJECT_NAME
        Case KEY_CODE.FinalMode:        Let key = FINAL_MODE_NAME
        Case KEY_CODE.IMEAccept:        Let key = IME_ACCEPT_NAME
        Case KEY_CODE.IMEConvert:       Let key = IME_CONVERT_NAME
        Case KEY_CODE.IMEModeChange:    Let key = IME_MODE_CHANGE_NAME
        Case KEY_CODE.IMENonconvert:    Let key = IME_NONCONVERT_NAME
        Case KEY_CODE.Jump:             Let key = JUMP_NAME
        Case KEY_CODE.JunjaMode:        Let key = JUNJA_MODE_NAME
        Case KEY_CODE.KanaMode:         Let key = KANA_MODE_NAME
        Case KEY_CODE.KanjiMode:        Let key = KANJI_MODE_NAME
        Case KEY_CODE.LeftControl:      Let key = LEFT_CONTROL_NAME
        Case KEY_CODE.LeftMenu:         Let key = LEFT_MENU_NAME
        Case KEY_CODE.LeftShift:        Let key = LEFT_SHIFT_NAME
        Case KEY_CODE.LeftWin:          Let key = LEFT_WIN_NAME
        Case KEY_CODE.LineFeed:         Let key = LINE_FEED_NAME
        Case KEY_CODE.MenuKey:          Let key = MENU_NAME
        Case KEY_CODE.NextTrack:        Let key = NEXT_NAME
        Case KEY_CODE.NoName:           Let key = NO_NAME_NAME
        Case KEY_CODE.Pa1:              Let key = PA1_NAME
        Case KEY_CODE.Packet:           Let key = PACKET_NAME
        Case KEY_CODE.Play:             Let key = PLAY_NAME
        Case KEY_CODE.PlayPause:        Let key = PLAY_PAUSE_NAME
        Case KEY_CODE.PlayStop:         Let key = PLAY_STOP_NAME
        Case KEY_CODE.Power:            Let key = POWER_NAME
        Case KEY_CODE.PrevTrack:        Let key = PREVIOUS_TRACK_NAME
        Case KEY_CODE.PrintKey:         Let key = PRINT_NAME
        Case KEY_CODE.PrintScreen:      Let key = PRINT_SCREEN_NAME
        Case KEY_CODE.ProcessKey:       Let key = PROCESS_NAME
        Case KEY_CODE.Reset:            Let key = RESET_NAME
        Case KEY_CODE.RightControl:     Let key = RIGHT_BUTTON_NAME
        Case KEY_CODE.RightMenu:        Let key = RIGHT_CONTROL_NAME
        Case KEY_CODE.RightShift:       Let key = RIGHT_SHIFT_NAME
        Case KEY_CODE.RightWin:         Let key = RIGHT_MENU_NAME
        Case KEY_CODE.PrintKey:         Let key = SCREENSHOT_NAME
        Case KEY_CODE.SelectKey:        Let key = SELECT_NAME
        Case KEY_CODE.SelectMedia:      Let key = SELECT_MEDIA_NAME
        Case KEY_CODE.Sleep:            Let key = SLEEP_NAME
        Case KEY_CODE.StartApp1:        Let key = START_APP_1_NAME
        Case KEY_CODE.StartApp2:        Let key = START_APP_2_NAME
        Case KEY_CODE.StartMail:        Let key = START_MAIL_NAME
        Case KEY_CODE.StopMedia:        Let key = STOP_NAME
        Case KEY_CODE.VolumeDown:       Let key = VOLUME_DOWN_NAME
        Case KEY_CODE.VolumeMute:       Let key = VOLUME_MUTE_NAME
        Case KEY_CODE.VolumeUp:         Let key = VOLUME_UP_NAME
        Case KEY_CODE.Wake:             Let key = WAKE_NAME
        Case KEY_CODE.Zoom:             Let key = ZOOM_NAME
        ' Unknown
        Case Else: Let key = getUnknown() & OPEN_ROUND_BRACKET & keyCode & CLOSE_ROUND_BRACKET ' Eg. Unknown(20)
    End Select
    'Skip modifier
    If key = vbNullString Then
        Let convertKeyToName = vbNullString
    Else
        Let convertKeyToName = modifier & key
    End If
End Function

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
        Let code = Mid$(code, LEN_MODIFIER_CODE + 1)
    End If
    ' Handle Shift modifier
    If left$(code, LEN_MODIFIER_CODE) = SHIFT_CODE Then
        Let modifier = modifier & SHIFT_PLUS
        Let code = Mid$(code, LEN_MODIFIER_CODE + 1)
    End If
    ' Handle Alt modifier
    If left$(code, LEN_MODIFIER_CODE) = ALT_CODE Then
        Let modifier = modifier & ALT_PLUS
        Let code = Mid$(code, LEN_MODIFIER_CODE + 1)
    End If
    ' Map the remaining code to key name
    Select Case code
        Case NUMPAD0_CODE:        Let code = NUMPAD0_NAME
        Case NUMPAD1_CODE:        Let code = NUMPAD1_NAME
        Case NUMPAD2_CODE:        Let code = NUMPAD2_NAME
        Case NUMPAD3_CODE:        Let code = NUMPAD3_NAME
        Case NUMPAD4_CODE:        Let code = NUMPAD4_NAME
        Case NUMPAD5_CODE:        Let code = NUMPAD5_NAME
        Case NUMPAD6_CODE:        Let code = NUMPAD7_NAME
        Case NUMPAD7_CODE:        Let code = NUMPAD6_NAME
        Case NUMPAD8_CODE:        Let code = NUMPAD8_NAME
        Case NUMPAD9_CODE:        Let code = NUMPAD9_NAME

        Case BACKSPACE_CODE:      Let code = BACKSPACE_NAME
        Case TAB_CODE:            Let code = TAB_NAME
        Case CLEAR_CODE:          Let code = CLEAR_NAME
        Case ENTER_CODE:          Let code = ENTER_NAME
        Case BREAK_CODE:          Let code = BREAK_NAME
        Case CAPS_LOCK_CODE:      Let code = CAPS_LOCK_NAME
        Case ESCAPE_CODE:         Let code = ESCAPE_NAME
        Case SPACE_CODE:          Let code = SPACE_NAME
        Case PAGE_UP_CODE:        Let code = PAGE_UP_NAME
        Case PAGE_DOWN_CODE:      Let code = PAGE_DOWN_NAME
        Case END_CODE:            Let code = END_NAME
        Case HOME_CODE:           Let code = HOME_NAME
        Case LEFT_ARROW_CODE:     Let code = LEFT_ARROW_NAME
        Case UP_ARROW_CODE:       Let code = UP_ARROW_NAME
        Case RIGHT_ARROW_CODE:    Let code = RIGHT_ARROW_NAME
        Case DOWN_ARROW_CODE:     Let code = DOWN_ARROW_NAME
        Case INSERT_CODE:         Let code = INSERT_NAME
        Case DELETE_CODE:         Let code = DELETE_NAME
        Case HELP_CODE:           Let code = HELP_NAME
        Case MULTIPLY_CODE:       Let code = MULTIPLY_NAME
        Case ADD_CODE:            Let code = ADD_NAME
        Case ENTER_NUM_CODE:      Let code = ENTER_NUM_NAME
        Case SUBTRACT_CODE:       Let code = SUBTRACT_NAME
        Case DECIMAL_CODE:        Let code = DECIMAL_NAME
        Case DIVIDE_CODE:         Let code = DIVIDE_NAME
        Case NUM_LOCK_CODE:       Let code = NUM_LOCK_NAME
        Case SCROLL_LOCK_CODE:    Let code = SCROLL_LOCK_NAME

        Case HYPHEN_CODE:         Let code = HYPHEN_NAME
        Case EQUALS_CODE:         Let code = EQUALS_NAME
        Case COMMA_CODE:          Let code = COMMA_NAME
        Case OPEN_BRACKET_CODE:   Let code = OPEN_BRACKET_NAME
        Case CLOSE_BRACKET_CODE:  Let code = CLOSE_BRACKET_NAME
        Case BACKSLASH_CODE:      Let code = BACKSLASH_NAME
        Case SLASH_CODE:          Let code = SLASH_NAME
        Case SEMICOLON_CODE:      Let code = SEMICOLON_NAME
        Case APOSTROPHE_CODE:     Let code = APOSTROPHE_NAME
        Case PERIOD_CODE:         Let code = PERIOD_NAME
        Case FORWARD_SLASH_CODE:  Let code = FORWARD_SLASH_NAME
        Case GRAVE_ACCENT_CODE:   Let code = GRAVE_ACCENT_NAME

    End Select
    ' Remove Brackets
    Let code = Replace(code, OPEN_CURLY_BRACKET, vbNullString)
    Let code = Replace(code, CLOSE_CURLY_BRACKET, vbNullString)
    Let convertCodeToName = modifier & IIf(code = vbNullString, getNoSet(), code)
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
        Let name = Mid$(name, Len(CTRL_PLUS) + 1)
    End If
    ' Handle Shift modifier
    If left$(name, Len(SHIFT_PLUS)) = SHIFT_PLUS Then
        Let modifier = modifier & SHIFT_CODE
        Let name = Mid$(name, Len(SHIFT_PLUS) + 1)
    End If
    ' Handle Alt modifier
    If left$(name, Len(ALT_PLUS)) = ALT_PLUS Then
        Let modifier = modifier & ALT_CODE
        Let name = Mid$(name, Len(ALT_PLUS) + 1)
    End If
    ' Map the remaining key name back to its code
    Select Case name
        Case NUMPAD0_NAME:        Let name = NUMPAD0_CODE
        Case NUMPAD1_NAME:        Let name = NUMPAD1_CODE
        Case NUMPAD2_NAME:        Let name = NUMPAD2_CODE
        Case NUMPAD3_NAME:        Let name = NUMPAD3_CODE
        Case NUMPAD4_NAME:        Let name = NUMPAD4_CODE
        Case NUMPAD5_NAME:        Let name = NUMPAD5_CODE
        Case NUMPAD6_NAME:        Let name = NUMPAD6_CODE
        Case NUMPAD7_NAME:        Let name = NUMPAD7_CODE
        Case NUMPAD8_NAME:        Let name = NUMPAD8_CODE
        Case NUMPAD9_NAME:        Let name = NUMPAD9_CODE
        
        Case BACKSPACE_NAME:      Let name = BACKSPACE_CODE
        Case TAB_NAME:            Let name = TAB_CODE
        Case CLEAR_NAME:          Let name = CLEAR_CODE
        Case ENTER_NAME:          Let name = ENTER_CODE
        Case BREAK_NAME:          Let name = BREAK_CODE
        Case CAPS_LOCK_NAME:      Let name = CAPS_LOCK_CODE
        Case ESCAPE_NAME:         Let name = ESCAPE_CODE
        Case SPACE_NAME:          Let name = SPACE_CODE
        Case PAGE_UP_NAME:        Let name = PAGE_UP_CODE
        Case PAGE_DOWN_NAME:      Let name = PAGE_DOWN_CODE
        Case END_NAME:            Let name = END_CODE
        Case HOME_NAME:           Let name = HOME_CODE
        Case LEFT_ARROW_NAME:     Let name = LEFT_ARROW_CODE
        Case UP_ARROW_NAME:       Let name = UP_ARROW_CODE
        Case RIGHT_ARROW_NAME:    Let name = RIGHT_ARROW_CODE
        Case DOWN_ARROW_NAME:     Let name = DOWN_ARROW_CODE
        Case INSERT_NAME:         Let name = INSERT_CODE
        Case DELETE_NAME:         Let name = DELETE_CODE
        Case HELP_NAME:           Let name = HELP_CODE
        Case MULTIPLY_NAME:       Let name = MULTIPLY_CODE
        Case ADD_NAME:            Let name = ADD_CODE
        Case ENTER_NUM_NAME:      Let name = ENTER_NUM_CODE
        Case SUBTRACT_NAME:       Let name = SUBTRACT_CODE
        Case DECIMAL_NAME:        Let name = DECIMAL_CODE
        Case DIVIDE_NAME:         Let name = DIVIDE_CODE
        Case NUM_LOCK_NAME:       Let name = NUM_LOCK_CODE
        Case SCROLL_LOCK_NAME:    Let name = SCROLL_LOCK_CODE

        Case HYPHEN_CODE:         Let name = HYPHEN_NAME
        Case EQUALS_CODE:         Let name = EQUALS_NAME
        Case OPEN_BRACKET_CODE:   Let name = OPEN_BRACKET_NAME
        Case CLOSE_BRACKET_CODE:  Let name = CLOSE_BRACKET_NAME
        Case BACKSLASH_CODE:      Let name = BACKSLASH_NAME
        Case SLASH_CODE:          Let name = SLASH_NAME
        Case SEMICOLON_CODE:      Let name = SEMICOLON_NAME
        Case APOSTROPHE_CODE:     Let name = APOSTROPHE_NAME
        Case COMMA_CODE:          Let name = COMMA_NAME
        Case PERIOD_CODE:         Let name = PERIOD_NAME
        Case FORWARD_SLASH_CODE:  Let name = FORWARD_SLASH_NAME
        Case GRAVE_ACCENT_CODE:   Let name = GRAVE_ACCENT_NAME
        ' No mapping
        Case Else: Let name = OPEN_CURLY_BRACKET & name & CLOSE_CURLY_BRACKET
    End Select
    Let convertNameToCode = modifier & name
End Function

