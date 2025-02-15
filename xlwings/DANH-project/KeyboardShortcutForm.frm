VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} KeyboardShortcutForm 
   Caption         =   "Settings"
   ClientHeight    =   7665
   ClientLeft      =   -21
   ClientTop       =   -224
   ClientWidth     =   10864
   OleObjectBlob   =   "KeyboardShortcutForm.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "KeyboardShortcutForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Check README.md for more information

Option Explicit
' Declare Variables
Private userResponse As VbMsgBoxResult
Private info As InfoConstants
Private shortcutTb As ListObject
Private eventColl As Collection
Private hoverLabel As MsForms.label
Private editingLabel As MsForms.label
Private editingTextBox As MsForms.textBox
Private editingIndex As String 'Find index by replace so use String is better
Private editingShortcut As String
Private editedArr() As String ' TODO: Create custom array CRUD
Private pickingLabel As MsForms.label
Private pickingIndex As String 'Find index by replace so use String is better
Private shortcutC As ShortcutController
Private Enum COLOR
    title_hover = 16378841 'RGB(229, 243, 255)
    line_hover = 16774117 'RGB(217, 235, 249)
    line_selected = 16772040 'RGB(200, 235, 255)
    line_border = 14935011 'RGB(227, 227, 227)
    line_edited = 61690 'RGB(250, 240, 0)
    line_edited_hover = 51450 'RGB(250, 200, 0)
    line_edited_picking = 41210 'RGB(250, 160, 0)
    ' Default Variables
    highlight = vbHighlight
    window_text = vbWindowText
    button_shadow = vbButtonShadow
    window_background = vbWindowBackground
    menu_text = vbMenuText
    menu_bar = vbMenuBar
End Enum
Private Const DEFAULT = "<Default>"
Private Const MODIFIED = "<Modified>"
Private Const FILTER_PLACEHOLDER As String = "<Type to filter text>"
Private Const TITLE_TAG As String = "title"
Private Const LINE_TAG As String = "line_"
Private Const KEYBINDING_LABEL As String = "KeyBindingLabel_"
Private Const COMMAND_LABEL As String = "CommandLabel_"
Private Const KEYBINDING_TEXTBOX As String = "KeyBindingTextBox_"
Private Const WHEN_LABEL As String = "WhenLabel_"
Private Const STATUS_LABEL As String = "StatusLabel_"
Private Const LINE_1 As String = "Line1_"
Private Const LINE_2 As String = "Line2_"
Private Const LINE_3 As String = "Line3_"
Private Const LINE_4 As String = "Line4_"
Private Const LINE_HEIGHT As Byte = 13.5
Private Const LINE_FONT_SIZE As Byte = 8.5
Private Const MAX_LINE As Byte = 15
Private Const PROG_ID_LABEL As String = "Forms.Label.1"
Private Const PROG_ID_TEXTBOX As String = "Forms.Textbox.1"
Private Const ASTERISK As String = "*"
' Loop iterators
Private ctrl As MsForms.control
Private row As ListRow

' NOTE: Used to use Mutators/Accessors as Public for fix bug when cls form as instance but not working

' MUTATORS

Private Sub letUserResponse(ByRef value As VbMsgBoxResult): Let userResponse = value: End Sub
Private Sub setInfo(ByRef value As InfoConstants): Set info = value: End Sub
Private Sub setShortcutTb(ByRef value As ListObject): Set shortcutTb = value: End Sub
Private Sub setEventColl(ByRef value As Collection): Set eventColl = value: End Sub
Private Sub setHoverLabel(ByRef value As MsForms.label): Set hoverLabel = value: End Sub
Private Sub setEditingLabel(ByRef value As MsForms.label): Set editingLabel = value: End Sub
Private Sub setEditingTextBox(ByRef value As MsForms.textBox): Set editingTextBox = value: End Sub
Private Sub letEditingIndex(ByRef value As String): Let editingIndex = value: End Sub
Private Sub letEditingShortcut(ByRef value As String): Let editingShortcut = value: End Sub
Private Sub setPickingLabel(ByRef value As MsForms.label): Set pickingLabel = value: End Sub
Private Sub letPickingIndex(ByRef value As String): Let pickingIndex = value: End Sub
Private Sub setShortcutC(ByRef value As ShortcutController): Set shortcutC = value: End Sub
Private Sub setCtrl(ByRef value As ShortcutController): Set ctrl = value: End Sub
Private Sub setRow(ByRef value As ShortcutController): Set row = value: End Sub

' ACCESSORS

Private Function getUserResponse() As VbMsgBoxResult: Let getUserResponse = userResponse: End Function
Private Function getInfo() As InfoConstants: Let getInfo = info: End Function
Private Function getShortcutTb() As ListObject: Set getShortcutTb = shortcutTb: End Function
Private Function getEventColl() As Collection: Set getEventColl = eventColl: End Function
Private Function getHoverLabel() As MsForms.label: Set getHoverLabel = hoverLabel: End Function
Private Function getEditingLabel() As MsForms.label: Set getEditingLabel = editingLabel: End Function
Private Function getEditingTextBox() As MsForms.textBox: Set getEditingTextBox = editingTextBox: End Function
Private Function getEditingIndex() As String: Let getEditingIndex = editingIndex: End Function
Private Function getEditingShortcut() As String: Let getEditingShortcut = editingShortcut: End Function
Private Function getPickingLabel() As MsForms.label: Set getPickingLabel = pickingLabel: End Function
Private Function getPickingIndex() As String: Let getPickingIndex = pickingIndex: End Function
Private Function getShortcutC() As ShortcutController: Set getShortcutC = shortcutC: End Function
Private Function getCtrl() As ShortcutController: Set getCtrl = ctrl: End Function
Private Function getRow() As ShortcutController: Set getShortcutC = row: End Function

' CONSTRUCTOR

Private Sub UserForm_Initialize()
    Call setShortcutC(New ShortcutController)
    Call setInfo(New InfoConstants)
    Call invisiblePattern
    Call initComboBox
    Call initRow
    Call initLabel
    Call storeCustomEvent
'    Call InitTabIndexes( _
'        , FindWhatLabel _
'        , FindAreaInput _
'        , ReplaceWithLabel _
'        , ReplaceAreaInput _
'        , WithinLabel _
'        , WithInComboBox _
'        , SelectedAreaInput _
'        , SearchLabel _
'        , SearchComboBox _
'        , MatchCaseCheckfBox _
'        , MatchByteCheckBox _
'        , MatchContentCheckBox _
'        , LengthOrderCheckBox _
'        , ReplaceAllButton _
'        , CloseButton _
'    )
    Call MouseScroll.EnableMouseScroll(Me) ' Apply the mousewheel scrolling
End Sub

' DESTRUCTOR

Private Sub UserForm_Terminate()
    Call MouseScroll.DisableMouseScroll(Me) ' Remove the mousewheel scrolling
    Call cleanUp
End Sub

' EVENTS

Private Sub KeyboardFrameContainer_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Call initLabel
End Sub

Private Sub KeyboardFrame_KeyDown(ByVal KeyCode As MsForms.ReturnInteger, ByVal Shift As Integer)
    Select Case KeyCode
        ' ENTER key only
        ' TODO: vbKeyReturn to GetShortcutC().getEnterKey()
        Case vbKeyReturn: If Shift <> 1 Then Call showEditing
    End Select
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Call initLabel
End Sub

Private Sub UserForm_Deactivate()
    Call initLabel
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Call MouseScroll.DisableMouseScroll(Me) ' Remove the mousewheel scrolling
    Call cleanUp
End Sub

Private Sub ApplyAndCloseButton_Click()
    Call KeyboardShortcutForm.closeForm
End Sub

Private Sub CancelButton_Click()
    Call KeyboardShortcutForm.closeForm
End Sub

Private Sub FilterPlaceTextBox_Enter()
    'Place Holder Handel
    If FilterPlaceTextBox.value = FILTER_PLACEHOLDER Then
        FilterPlaceTextBox.value = vbNullString
        FilterPlaceTextBox.foreColor = COLOR.window_text
    End If
End Sub

Private Sub FilterPlaceTextBox_Exit(ByVal Cancel As MsForms.ReturnBoolean)
    'Place Holder Handel
    If FilterPlaceTextBox.value = vbNullString Then
        FilterPlaceTextBox.value = FILTER_PLACEHOLDER
        FilterPlaceTextBox.foreColor = COLOR.button_shadow
    End If
End Sub

' METHODS

Private Sub initComboBox()
    With ProfilesComboBox
        .Style = fmStyleDropDownList
        .AddItem DEFAULT
        .ListIndex = "0" '0 Or False
    End With
End Sub
    
Private Sub initRow()
    Dim keybind As String
    Dim keybindDefault As String
    Dim lineIndex As String
    Dim defaultMark As String
    Dim shortcut As String
    Call setShortcutTb(getShortcutC().getShortcutTable())
    ReDim editedArr(getShortcutTb().ListRows.Count - 1)
    For Each row In getShortcutTb().ListRows
        Let keybind = row.Range(1, getShortcutC().getCustomKeybindColumn())
        Let keybindDefault = row.Range(1, getShortcutC().getDefaultKeybindColumn())
        Let lineIndex = row.Range(1, getShortcutC().getNoColumn())
        Let shortcut = getShortcutC().convertCodeToName(keybind)
        Let defaultMark = IIf(keybind = keybindDefault, DEFAULT, vbNullString)
        Call createRow( _
            index:=row.index _
            , command:=row.Range(1, getShortcutC().getCommandColumn()) _
            , shortcut:=shortcut _
            , when:=row.Range(1, getShortcutC().getWhenColumn()) _
            , status:=row.Range(1, getShortcutC().getStatusColumn()) & defaultMark _
        )
    Next row
    Call createScrollBar(getShortcutTb().DataBodyRange.Rows.Count) ' Max Row
End Sub

Private Sub createScrollBar(ByRef totalLines As Long)
    If totalLines > MAX_LINE Then
        Me.KeyboardFrame.ScrollBars = fmScrollBarsVertical
        Me.KeyboardFrame.ScrollHeight = totalLines * LINE_HEIGHT
    Else
        Me.KeyboardFrame.ScrollBars = fmScrollBarsNone
    End If
End Sub

Private Sub createRow( _
    ByRef index As Long _
    , ByRef command As String _
    , ByRef shortcut As String _
    , ByRef when As String _
    , ByRef status As String _
)
    Call createRowLabel( _
        name:=COMMAND_LABEL & index _
        , index:=index _
        , caption:=command _
        , width:=CommandLabel_0.width _
        , height:=CommandLabel_0.height _
        , top:=CommandLabel_0.top _
        , left:=CommandLabel_0.left _
    )
    ' TextBox need create first for it can be back of label
    Call createRowTextBox( _
        name:=KEYBINDING_TEXTBOX & index _
        , index:=index _
        , width:=KeybindingTextBox_0.width _
        , height:=KeybindingTextBox_0.height _
        , top:=KeybindingTextBox_0.top _
        , left:=KeybindingTextBox_0.left _
        , bold:=True _
        , italic:=True _
        , visible:=False _
    )
    Call createRowLabel( _
        name:=KEYBINDING_LABEL & index _
        , index:=index _
        , caption:=shortcut _
        , width:=KeyBindingLabel_0.width _
        , height:=KeyBindingLabel_0.height _
        , top:=KeyBindingLabel_0.top _
        , left:=KeyBindingLabel_0.left _
    )
    Call createRowLabel( _
        name:=WHEN_LABEL & index _
        , index:=index _
        , caption:=when _
        , width:=WhenLabel_0.width _
        , height:=WhenLabel_0.height _
        , top:=WhenLabel_0.top _
        , left:=WhenLabel_0.left _
    )
    Call createRowLabel( _
        name:=STATUS_LABEL & index _
        , index:=index _
        , caption:=Status _
        , width:=StatusLabel_0.width _
        , height:=StatusLabel_0.height _
        , top:=StatusLabel_0.top _
        , left:=StatusLabel_0.left _
    )
    Call createRowLabel( _
        name:=LINE_1 & index _
        , index:=index _
        , width:=Line1_0.width _
        , height:=Line1_0.height _
        , top:=Line1_0.top _
        , left:=Line1_0.left _
        , hasBorder:=fmBorderStyleSingle _
    )
    Call createRowLabel( _
        name:=LINE_2 & index _
        , index:=index _
        , width:=Line2_0.width _
        , height:=Line2_0.height _
        , top:=Line2_0.top _
        , left:=Line2_0.left _
        , hasBorder:=fmBorderStyleSingle _
    )
    Call createRowLabel( _
        name:=LINE_3 & index _
        , index:=index _
        , width:=Line3_0.width _
        , height:=Line3_0.height _
        , top:=Line3_0.top _
        , left:=Line3_0.left _
        , hasBorder:=fmBorderStyleSingle _
    )
    Call createRowLabel( _
        name:=LINE_4 & index _
        , index:=index _
        , width:=Line4_0.width _
        , height:=Line4_0.height _
        , top:=Line4_0.top _
        , left:=Line4_0.left _
        , hasBorder:=fmBorderStyleSingle _
    )
End Sub

Private Sub createRowLabel( _
    ByRef name As String _
    , ByRef index As Long _
    , ByRef width As Single _
    , ByRef height As Single _
    , ByRef top As Single _
    , ByRef left As Single _
    , Optional ByRef caption As String = vbNullString _
    , Optional ByRef backColor As String = COLOR.window_background _
    , Optional ByRef foreColor As String = COLOR.window_text _
    , Optional ByRef hasBorder As Byte = fmBorderStyleNone _
)
    Dim lineLabel As MsForms.label
    Set lineLabel = Me.KeyboardFrame.add( _
        bstrProgId:=PROG_ID_LABEL _
        , name:=name _
        , visible:=True)
    With lineLabel
    .caption = Space(1) & caption
    .Tag = LINE_TAG & index
    .height = height
    .width = width
    .top = top + (index - 1) * LINE_HEIGHT 'Top - 1 for the title
    .left = left
    .backColor = backColor
    .foreColor = foreColor
    .BorderStyle = hasBorder
    .bordercolor = COLOR.line_border
    .Font.size = LINE_FONT_SIZE
    End With
    Set lineLabel = Nothing
End Sub

Private Sub createRowTextBox( _
    ByRef name As String _
    , ByRef index As Long _
    , ByRef width As Single _
    , ByRef height As Single _
    , ByRef top As Single _
    , ByRef left As Single _
    , Optional ByRef backColor As String = COLOR.window_background _
    , Optional ByRef foreColor As String = COLOR.window_text _
    , Optional ByRef hasBorder As Byte = fmBorderStyleNone _
    , Optional ByRef effect As Byte = fmSpecialEffectFlat _
    , Optional ByRef visible As Boolean = True _
    , Optional ByRef italic As Boolean = False _
    , Optional ByRef bold As Boolean = False _
)
    ' Dim lineTextbox As MSForms.textBox
    Dim lineTextbox As MsForms.control
    Set lineTextbox = Me.KeyboardFrame.add( _
        bstrProgId:=PROG_ID_TEXTBOX _
        , name:=name _
        , visible:=True)
    With lineTextbox
    Let .Tag = LINE_TAG & index
    Let .height = height
    Let .width = width
    Let .top = top + (index - 1) * LINE_HEIGHT 'Top - 1 for the title
    Let .left = left
    Let .backColor = backColor
    Let .foreColor = foreColor
    Let .BorderStyle = hasBorder
    Let .bordercolor = COLOR.line_border
    Let .Font.size = LINE_FONT_SIZE
    Let .FontItalic = italic
    Let .FontBold = bold
    Let .SpecialEffect = effect
    Let .visible = visible
    End With
    Set lineTextbox = Nothing
End Sub

'INIT

Private Sub initLabel()
    Dim lineIndex As String
    Dim isLabel As Boolean
    Dim isHoverLabel As Boolean
    Dim isPickingLabel As Boolean
    Dim isTitles As Boolean
    Dim isLines As Boolean
    Dim isKeybind As Boolean
    ' Loop through all controls for find suitable
    For Each ctrl In Me.KeyboardFrameContainer.controls
        With ctrl
        Let lineIndex = Replace(.Tag, LINE_TAG, vbNullString)
        Let isLabel = (TypeOf ctrl Is MsForms.label)
        Let isHoverLabel = (ctrl Is getHoverLabel())
        Let isPickingLabel = (ctrl Is getPickingLabel()) Or (lineIndex = getPickingIndex())
        ' Skip reset pickingLabel, hoverLabel
        If isLabel _
            And Not isHoverLabel _
            And Not isPickingLabel _
        Then
            Let isTitles = (.Tag = TITLE_TAG)
            Let isLines = (.Tag Like (LINE_TAG & ASTERISK))
            Let isKeybind = (.name Like (KEYBINDING_LABEL & ASTERISK))
            'Titles reset
            If _
                isTitles _
                And .backColor <> COLOR.menu_bar _
            Then
                Let .backColor = COLOR.menu_bar
            'Lines reset
            ElseIf _
                isLines _
                And .backColor <> COLOR.window_background _
            Then
                Let .backColor = COLOR.window_background
            End If
            'Keybind label reset
            If _
                isKeybind _
                And .foreColor <> COLOR.menu_text _
            Then
                Let .foreColor = COLOR.menu_text
                Let .FontItalic = False
                Let .FontBold = False
            End If
        End If
        End With
    Next ctrl
End Sub

Private Sub addLabelEvent(ByRef ctrl As MsForms.control)
    With ctrl
    If .Tag = TITLE_TAG Then getEventColl().add createLabelEvent(ctrl)
    If .Tag Like (LINE_TAG & ASTERISK) Then getEventColl().add createLabelEvent(ctrl)
    End With
End Sub

Private Sub addTextBoxEvent(ByRef ctrl As MsForms.control)
    With ctrl
    If .Tag Like (LINE_TAG & ASTERISK) Then
        getEventColl().add createTextBoxEvent(ctrl)
    End If
    End With
End Sub

Private Sub storeCustomEvent()
    'Must store control with event inside a collection by adding functions return custom class in that collection.
    Call setEventColl(New Collection)
    ' Loop thourgh all controls for find suitable
    For Each ctrl In Me.KeyboardFrameContainer.controls
        If TypeOf ctrl Is MsForms.label Then Call addLabelEvent(ctrl)
        If TypeOf ctrl Is MsForms.textBox Then Call addTextBoxEvent(ctrl)
    Next ctrl
End Sub

' CUSTOM EVENTS

Private Function createLabelEvent(ByRef ctrl As MsForms.control) As CustomLabelEvent
    Dim labelE As CustomLabelEvent: Set labelE = New CustomLabelEvent
    Set labelE.setLabel = ctrl
    Set createLabelEvent = labelE
    Set labelE = Nothing
End Function

Private Function createTextBoxEvent(ByRef ctrl As MsForms.control) As CustomTextBoxEvent
    Dim textBoxE As CustomTextBoxEvent: Set textBoxE = New CustomTextBoxEvent
    Set textBoxE.setTextBox = ctrl
    Set createTextBoxEvent = textBoxE
    Set textBoxE = Nothing
End Function

Public Sub labelMoveOn(ByRef label As MsForms.label)
    Dim keybindingLabel As MsForms.label
    Dim lineIndex As String
    Dim isHoverLabel As Boolean
    Dim isPickingLabel As Boolean
    Dim isHover As Boolean
    ' Reset Label
    Call initLabel
    ' Titles Hover
    If label.Tag = TITLE_TAG Then
        Let label.backColor = COLOR.title_hover
    ' Lines Hover
    ElseIf label.Tag Like (LINE_TAG & ASTERISK) Then
        ' Highlight lines
        For Each ctrl In Me.KeyboardFrameContainer.controls
            With ctrl
            Let lineIndex = Replace(.Tag, LINE_TAG, vbNullString)
            Let isHoverLabel = (TypeOf ctrl Is MsForms.label) And (.Tag = label.Tag)
            Let isPickingLabel = (ctrl Is getPickingLabel()) Or (lineIndex = getPickingIndex())
            ' Better performance
            Let isHover = (.backColor = COLOR.line_hover)
            If _
                isHoverLabel _
                And Not isPickingLabel _
                And Not isHover _
            Then
                Let .backColor = COLOR.line_hover
                Call setHoverLabel(ctrl)
            End If
            End With
        Next ctrl
        ' Highlight label
        Set keybindingLabel = Me.KeyboardFrame.controls(KEYBINDING_LABEL & Replace(label.Tag, LINE_TAG, vbNullString))
        With keybindingLabel
        Let isPickingLabel = (label Is getPickingLabel())
        ' Better performance
        Let isHover = (.foreColor = COLOR.highlight)
        If _
            Not isPickingLabel _
            And Not isHover _
        Then
            Let .foreColor = COLOR.highlight
            Let .FontItalic = True
            Let .FontBold = True
        End If
        End With
    End If
    Set keybindingLabel = Nothing
End Sub

Public Sub labelClick(ByRef label As MsForms.label)
    With label
    ' Titles Click
    If .Tag = TITLE_TAG Then
        MsgBox ("TODO: Sort By" & .caption)
    ' Lines Click
    ElseIf .Tag Like (LINE_TAG & ASTERISK) Then
        Call resetPicking
        Call hideEditing
        ' Assign picked Label
        Call letPickingIndex(Replace(label.Tag, LINE_TAG, vbNullString))
        Call setPickingLabel(Me.KeyboardFrame.controls(KEYBINDING_LABEL & getPickingIndex()))
        Call highlightPicking
    End If
    End With
End Sub

Public Sub labelDbClick(ByRef label As MsForms.label)
    If label.Tag Like (LINE_TAG & ASTERISK) Then
        Call hideEditing
        Call showEditing
    End If
End Sub

Public Sub textBoxKeyDown( _
    ByRef textBox As MsForms.textBox _
    , ByRef KeyCode As MsForms.ReturnInteger _
    , ByRef Shift As Integer _
)
    letEditingShortcut (getShortcutC().convertKeyToName(KeyCode, Shift))
    If editingShortcut = getShortcutC().getEnterKey() Then
        Call hideEditing
    ElseIf getEditingShortcut() = getShortcutC().getEscKey() Then
        ' Load before edit shortcut
        Let textBox.text = LTrim(getEditingLabel().caption)
        Call hideEditing
    ElseIf getEditingShortcut() = getShortcutC().getBackspaceKey() Then
        Let textBox.text = vbNullString
    Else
        Let textBox.text = getEditingShortcut()
    End If
    'Prevent default keyDown
    Let KeyCode = 0
End Sub

Public Sub textBoxChange(ByRef textBox As MsForms.textBox)
'    If GetEditingLabel() Is Nothing Then Exit Sub
'    Let GetEditingLabel().caption = Space(1) & textBox.text
End Sub

Private Sub showEditing()
    Call letEditingIndex(Replace(getPickingLabel().Tag, LINE_TAG, vbNullString))
    'Assign editing shortcut object
    Call setEditingLabel(Me.KeyboardFrame.controls(KEYBINDING_LABEL & getEditingIndex()))
    Call setEditingTextBox(Me.KeyboardFrame.controls(KEYBINDING_TEXTBOX & getEditingIndex()))
    'If not keybinding set textbox to Blank
    Let getEditingTextBox.text = IIf( _
        LTrim(getEditingLabel().caption) = getShortcutC().getNoSet() _
        , vbNullString _
        , LTrim(getEditingLabel().caption) _
    )
    'Show display
    Let getEditingTextBox().visible = True
    Let getEditingLabel().visible = False
    Call getEditingTextBox().SetFocus
End Sub

Private Sub hideEditing()
    Dim lineIndex As String
    'Check for the 1st time
    If getEditingLabel() Is Nothing Then Exit Sub
    If getEditingTextBox() Is Nothing Then Exit Sub
    'If textBox blank set label to No Set
    Let getEditingLabel.caption = IIF( _
        getEditingTextBox().text = vbNullString _
        , getShortcutC().getNoSet() _
        , Space(1) & getEditingTextBox().text _
    )
    Let lineIndex = Replace(getEditingTextBox().Tag, LINE_TAG, vbNullString)
    'Update edited shortcut
    Call updateEdited( _
        key:= lineIndex _
        , value:=getEditingTextBox().text _
    )
    'Highlight edited
    Call highlightEdited
    'Hide display
    Let getEditingTextBox().visible = False
    Let getEditingLabel().visible = True
    'Clean editing
    Call setEditingTextBox(Nothing)
    Call setEditingLabel(Nothing)
End Sub

Private Sub highlightPicking()
    Dim lineIndex As String
    Dim isLabel As Boolean
    Dim isPickingLabel As Boolean
    'Highlight line
    For Each ctrl In Me.KeyboardFrame.controls
        With ctrl
        Let lineIndex = Replace(.Tag, LINE_TAG, vbNullString)
        Let isLabel = (TypeOf ctrl Is MsForms.label)
        Let isPickingLabel = (getPickingIndex() = lineIndex)
        If _
            isLabel _
            And isPickingLabel _
            And .backColor <> COLOR.line_selected _
        Then
            'Highlight line
            Let .backColor = COLOR.line_selected
            'Highlight label
            Let .foreColor = COLOR.highlight
            Let .FontItalic = False
            Let .FontBold = True
        End If
        End With
    Next ctrl
End Sub

Private Sub resetPicking()
    Dim lineIndex As String
    Dim isLabel As Boolean
    Dim isPickingLabel As Boolean
    For Each ctrl In Me.KeyboardFrame.controls
        With ctrl
        Let lineIndex = Replace(.Tag, LINE_TAG, vbNullString)
        Let isLabel = (TypeOf ctrl Is MsForms.label)
        Let isPickingLabel = (getPickingIndex() = lineIndex)
        If _
            isLabel _
            And isPickingLabel _
            And .backColor <> COLOR.window_background _
        Then
            'Highlight line
            Let .backColor = COLOR.window_background
            'Highlight label
            Let .foreColor = COLOR.menu_text
            Let .FontItalic = False
            Let .FontBold = False
        End If
        End With
    Next ctrl
End Sub

Private Sub highlightEdited()
    Dim i As Integer
    Dim lineIndex As String
    Dim isLabel As Boolean
    Dim isEditedLabel As Boolean
    Dim isEdited As Boolean
    For Each ctrl In Me.KeyboardFrame.controls
        With ctrl
        Let lineIndex = Replace(.Tag, LINE_TAG, vbNullString)
        Let isLabel = (TypeOf ctrl Is MsForms.label)
        'Loop through edited array
        For i = LBound(editedArr) To UBound(editedArr)
            Let isEditedLabel = (lineIndex = i + 1) And Not (editedArr(i) = DEFAULT)
            Let isEdited = (.backColor <> COLOR.line_edited_picking)
            'Highlight line
            If _
                isLabel _
                And isEditedLabel _
                And isEdited _
            Then
                Let .backColor = COLOR.line_edited_picking
                Let .FontItalic = True
            End If
        Next i
        End With
    Next ctrl
End Sub

Private Sub updateEdited(ByRef key As String, ByRef value As String)
    Dim keybind As String
    Dim lineIndex As String
    Dim shortcut As String
    For Each row In getShortcutTb().ListRows
        Let keybind = row.Range(1, getShortcutC().getCustomKeybindColumn())
        Let shortcut = getShortcutC().convertCodeToName(keybind)
        Let lineIndex = row.Range(1, getShortcutC().getNoColumn())
        If lineIndex = key Then
            'DEFAULT
            IF value = shortcut Then
                Let editedArr(lineIndex - 1) = DEFAULT
            'EDITED
            Else
                Let editedArr(lineIndex - 1) = value
            End If
        Exit For ' row loop
        End If
    Next row
End Sub

'Todo: create system update collection.exist

' CLEAN
Private Sub invisiblePattern()
    Let Me.KeyboardFrame.backColor = vbInactiveBorder
    ' _0 is pattern
    Let Me.Line1_0.visible = False
    Let Me.Line2_0.visible = False
    Let Me.Line3_0.visible = False
    Let Me.Line4_0.visible = False
    Let Me.CommandLabel_0.visible = False
    Let Me.KeyBindingLabel_0.visible = False
    Let Me.KeybindingTextBox_0.visible = False
    Let Me.WhenLabel_0.visible = False
    Let Me.StatusLabel_0.visible = False
End Sub

Public Sub closeForm()
    Call Unload(Me)
    End
End Sub

Private Sub cleanUp()
    ' Clear Objects
    Call setInfo(Nothing)
    Call setEventColl(Nothing)
    Call setEditingLabel(Nothing)
    Call setEditingTextBox(Nothing)
    Call setPickingLabel(Nothing)
    Call setShortcutTb(Nothing)
    Call setCtrl(Nothing)
    Call setRow(Nothing)
    Call setShortcutC(Nothing)
    ' Clear Arrays
    Erase editedArr
End Sub
