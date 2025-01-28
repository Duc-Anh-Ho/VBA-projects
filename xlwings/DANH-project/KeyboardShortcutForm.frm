VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} KeyboardShortcutForm 
   Caption         =   "Settings"
   ClientHeight    =   5610
   ClientLeft      =   105
   ClientTop       =   390
   ClientWidth     =   8115
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
Private editingLabel As MSForms.label
Private editingTextBox As MSForms.textBox
Private editingIndex As String 'Find index by replace so use String is better
Private editingShortcut As String
Private editedArr() As String
Private pickingLabel As MSForms.label
Private pickingIndex As String 'Find index by replace so use String is better
Private shortcutC As ShortcutController
Private Enum COLOR
    title_hover = 16378841 'RGB(229, 243, 255)
    line_hover = 16774117 'RGB(217, 235, 249)
    line_selected = 16772040 'RGB(200, 235, 255)
    line_border = 14935011 'RGB(227, 227, 227)
    ' Default Variables
    highlight = vbHighlight
    window_text = vbWindowText
    button_shadow = vbButtonShadow
    window_background = vbWindowBackground
    menu_text = vbMenuText
End Enum
Private Const FILTER_PLACEHOLDER As String = "<Type to filter text>"
Private Const TITLE_TAG As String = "title"
Private Const LINE_TAG As String = "line_"
Private Const KEYBINDING_LABEL As String = "KeyBindingLabel_"
Private Const COMMAND_LABEL As String = "CommandLabel_"
Private Const KEYBINDING_TEXTBOX As String = "KeyBindingTextBox_"
Private Const WHEN_LABEL As String = "WhenLabel_"
Private Const SOURCE_LABEL As String = "SourceLabel_"
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
Private ctrl As MSForms.control
Private row As ListRow

'EVENTS

Private Sub KeyboardFrameContainer_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Call initLabel
End Sub

Private Sub KeyboardFrame_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Select Case KeyCode
        ' ENTER key only
        ' TODO: vbKeyReturn to shortcutC.getEnterKey()
        Case vbKeyReturn: If Shift <> 1 Then Call showEditing
    End Select
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Call initLabel
End Sub

Private Sub UserForm_Deactivate()
    Call initLabel
End Sub

Private Sub UserForm_Initialize()
    Set shortcutC = New ShortcutController
    Set info = New InfoConstants
    Call invisiblePattern
    Call initComboBox
    Call initLine
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

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Call MouseScroll.DisableMouseScroll(Me) ' Remove the mousewheel scrolling
    Call cleanUp
End Sub

Private Sub UserForm_Terminate()
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

Private Sub FilterPlaceTextBox_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    'Place Holder Handel
    If FilterPlaceTextBox.value = vbNullString Then
        FilterPlaceTextBox.value = FILTER_PLACEHOLDER
        FilterPlaceTextBox.foreColor = COLOR.button_shadow
    End If
End Sub

' ACCESSORS

' MUTATORS

' FUNCTIONS

Private Sub initComboBox()
    With ProfilesComboBox
        .Style = fmStyleDropDownList
        .AddItem "Default"
        .ListIndex = "0" '0 Or False
    End With
End Sub
    
Private Sub initLine()
    Dim keybind As String
    Dim keybindDefault As String
    Dim lineIndex As String
    Dim defaultMark As String
    Dim shortcut As String
    Set shortcutTb = shortcutC.getShortcutTable()
    ReDim editedArr(shortcutTb.ListRows.Count - 1)
    For Each row In shortcutTb.ListRows
        Let keybind = row.Range(1, shortcutC.getCustomKeybindColumn())
        Let keybindDefault = row.Range(1, shortcutC.getDefaultKeybindColumn())
        Let lineIndex = row.Range(1, shortcutC.getNoColumn())
        Let shortcut = shortcutC.convertCodeToName(keybind)
        Let defaultMark = IIf(keybind = keybindDefault, " *", vbNullString)
        Call createLine( _
            index:=row.index _
            , command:=row.Range(1, shortcutC.getCommandColumn()) _
            , shortcut:=shortcut _
            , when:=row.Range(1, shortcutC.getWhenColumn()) _
            , source:=row.Range(1, shortcutC.getSourceCoulmn()) & defaultMark _
        )
    Next row
    Call createScrollBar(shortcutTb.DataBodyRange.Rows.Count) ' Max Row
End Sub

Private Sub createScrollBar(ByRef totalLines As Long)
    If totalLines > MAX_LINE Then
        Me.KeyboardFrame.ScrollBars = fmScrollBarsVertical
        Me.KeyboardFrame.ScrollHeight = totalLines * LINE_HEIGHT
    Else
        Me.KeyboardFrame.ScrollBars = fmScrollBarsNone
    End If
End Sub

Private Sub createLine( _
    ByRef index As Long _
    , ByRef command As String _
    , ByRef shortcut As String _
    , ByRef when As String _
    , ByRef source As String _
)
    Call createLineLabel( _
        name:=COMMAND_LABEL & index _
        , index:=index _
        , caption:=command _
        , width:=CommandLabel_0.width _
        , height:=CommandLabel_0.height _
        , top:=CommandLabel_0.top _
        , left:=CommandLabel_0.left _
    )
    ' TextBox need create first for it can be back of label
    Call createLineTextBox( _
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
    Call createLineLabel( _
        name:=KEYBINDING_LABEL & index _
        , index:=index _
        , caption:=shortcut _
        , width:=KeyBindingLabel_0.width _
        , height:=KeyBindingLabel_0.height _
        , top:=KeyBindingLabel_0.top _
        , left:=KeyBindingLabel_0.left _
    )
    Call createLineLabel( _
        name:=WHEN_LABEL & index _
        , index:=index _
        , caption:=when _
        , width:=WhenLabel_0.width _
        , height:=WhenLabel_0.height _
        , top:=WhenLabel_0.top _
        , left:=WhenLabel_0.left _
    )
    Call createLineLabel( _
        name:=SOURCE_LABEL & index _
        , index:=index _
        , caption:=source _
        , width:=SourceLabel_0.width _
        , height:=SourceLabel_0.height _
        , top:=SourceLabel_0.top _
        , left:=SourceLabel_0.left _
    )
    Call createLineLabel( _
        name:=LINE_1 & index _
        , index:=index _
        , width:=Line1_0.width _
        , height:=Line1_0.height _
        , top:=Line1_0.top _
        , left:=Line1_0.left _
        , hasBorder:=fmBorderStyleSingle _
    )
    Call createLineLabel( _
        name:=LINE_2 & index _
        , index:=index _
        , width:=Line2_0.width _
        , height:=Line2_0.height _
        , top:=Line2_0.top _
        , left:=Line2_0.left _
        , hasBorder:=fmBorderStyleSingle _
    )
    Call createLineLabel( _
        name:=LINE_3 & index _
        , index:=index _
        , width:=Line3_0.width _
        , height:=Line3_0.height _
        , top:=Line3_0.top _
        , left:=Line3_0.left _
        , hasBorder:=fmBorderStyleSingle _
    )
    Call createLineLabel( _
        name:=LINE_4 & index _
        , index:=index _
        , width:=Line4_0.width _
        , height:=Line4_0.height _
        , top:=Line4_0.top _
        , left:=Line4_0.left _
        , hasBorder:=fmBorderStyleSingle _
    )
End Sub

Private Sub createLineLabel( _
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
    Dim lineLabel As MSForms.label
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

Private Sub createLineTextBox( _
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
    Dim lineTextbox As MSForms.control
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
    ' Loop thourgh all controls for find suitable
    For Each ctrl In Me.KeyboardFrameContainer.controls
        With ctrl
        ' Skip reseting pickingLabel
        If TypeOf ctrl Is MSForms.label _
            And Not ctrl Is pickingLabel _
            And pickingIndex <> Replace(.Tag, LINE_TAG, vbNullString) _
        Then
            'Titles Reset
            If _
                .Tag = TITLE_TAG _
                And .backColor <> vbMenuBar _
            Then
                 Let .backColor = vbMenuBar
            'Lines Reset
            ElseIf _
                .Tag Like (LINE_TAG & ASTERISK) _
                And .backColor <> COLOR.window_background _
            Then
                Let .backColor = COLOR.window_background
            End If
            'Keybind reset
            If _
                .name Like (KEYBINDING_LABEL & ASTERISK) _
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

Private Sub addLabelEvent(ByRef ctrl As MSForms.control)
    With ctrl
    If .Tag = TITLE_TAG Then eventColl.add createLabelEvent(ctrl)
    If .Tag Like (LINE_TAG & ASTERISK) Then eventColl.add createLabelEvent(ctrl)
    End With
End Sub

Private Sub addTextBoxEvent(ByRef ctrl As MSForms.control)
    With ctrl
    If .Tag Like (LINE_TAG & ASTERISK) Then
        eventColl.add createTextBoxEvent(ctrl)
    End If
    End With
End Sub

Private Sub storeCustomEvent()
    'Must store control with event inside a collection by adding functions return custom class in that collection.
    Set eventColl = New Collection
    ' Loop thourgh all controls for find suitable
    For Each ctrl In Me.KeyboardFrameContainer.controls
        If TypeOf ctrl Is MSForms.label Then Call addLabelEvent(ctrl)
        If TypeOf ctrl Is MSForms.textBox Then Call addTextBoxEvent(ctrl)
    Next ctrl
End Sub

' CUSTOM EVENTS

Private Function createLabelEvent(ByRef ctrl As MSForms.control) As CustomLabelEvent
    Dim labelE As CustomLabelEvent: Set labelE = New CustomLabelEvent
    Set labelE.setLabel = ctrl
    Set createLabelEvent = labelE
    Set labelE = Nothing
End Function

Private Function createTextBoxEvent(ByRef ctrl As MSForms.control) As CustomTextBoxEvent
    Dim textBoxE As CustomTextBoxEvent: Set textBoxE = New CustomTextBoxEvent
    Set textBoxE.setTextBox = ctrl
    Set createTextBoxEvent = textBoxE
    Set textBoxE = Nothing
End Function

Public Sub labelMoveOn(ByRef label As MSForms.label)
    Dim keybindingLabel As MSForms.label
    Call initLabel
    ' Titles Hover
    If label.Tag = TITLE_TAG Then
        Let label.backColor = COLOR.title_hover
    ' Lines Hover
    ElseIf label.Tag Like (LINE_TAG & ASTERISK) Then
        ' Hightlight line
        For Each ctrl In Me.KeyboardFrameContainer.controls
            With ctrl
            If _
                TypeOf ctrl Is MSForms.label _
                And .Tag = label.Tag _
                And Not ctrl Is pickingLabel _
                And pickingIndex <> Replace(.Tag, LINE_TAG, vbNullString) _
            Then
                If .backColor <> COLOR.line_hover Then
                    Let .backColor = COLOR.line_hover
                End If
            End If
            End With
        Next ctrl
        ' Highlight label
        Set keybindingLabel = Me.KeyboardFrame.controls(KEYBINDING_LABEL & Replace(label.Tag, LINE_TAG, vbNullString))
        With keybindingLabel
        If _
            Not label Is pickingLabel _
            And .foreColor <> COLOR.highlight _
        Then
            Let .foreColor = COLOR.highlight
            Let .FontItalic = True
            Let .FontBold = True
        End If
        End With
    End If
    Set keybindingLabel = Nothing
End Sub

Public Sub labelClick(ByRef label As MSForms.label)
    With label
    ' Titles Click
    If .Tag = TITLE_TAG Then
        MsgBox ("TODO: Sort By" & .caption)
    ' Lines Click
    ElseIf .Tag Like (LINE_TAG & ASTERISK) Then
        Call hidePicking
        Call hideEditing
        ' Assign picked Label
        Let pickingIndex = Replace(label.Tag, LINE_TAG, vbNullString)
        Set pickingLabel = Me.KeyboardFrame.controls(KEYBINDING_LABEL & pickingIndex)
        Call showPicking
    End If
    End With
End Sub

Public Sub labelDbClick(ByRef label As MSForms.label)
    If label.Tag Like (LINE_TAG & ASTERISK) Then
        Call hideEditing
        Call showEditing
    End If
End Sub

Public Sub textBoxKeyDown( _
    ByRef textBox As MSForms.textBox _
    , ByRef KeyCode As MSForms.ReturnInteger _
    , ByRef Shift As Integer _
)
    Let editingShortcut = shortcutC.convertKeyToName(KeyCode, Shift)
    If editingShortcut = shortcutC.getEnterKey() Then
        Call hideEditing
    ElseIf editingShortcut = shortcutC.getEscKey() Then
        ' Load before edit shotcut
        Let textBox.text = LTrim(editingLabel.caption)
        Call hideEditing
    ElseIf editingShortcut = shortcutC.getBackspaceKey() Then
        Let textBox.text = vbNullString
    Else
        Let textBox.text = editingShortcut
    End If
    'Prevent default keyDown
    Let KeyCode = 0
End Sub

Public Sub textBoxChange(ByRef textBox As MSForms.textBox)
'    If editingLabel Is Nothing Then Exit Sub
'    Let editingLabel.caption = Space(1) & textBox.text
End Sub

Private Sub showEditing()
    Let editingIndex = Replace(pickingLabel.Tag, LINE_TAG, vbNullString)
    'Assign editing shortcut object
    Set editingLabel = Me.KeyboardFrame.controls(KEYBINDING_LABEL & editingIndex)
    Set editingTextBox = Me.KeyboardFrame.controls(KEYBINDING_TEXTBOX & editingIndex)
    Let editingTextBox.text = LTrim(editingLabel.caption)
    'Show display
    Let editingTextBox.visible = True
    Let editingLabel.visible = False
    Call editingTextBox.SetFocus
End Sub

Private Sub hideEditing()
    'Check for the 1st time
    If editingLabel Is Nothing Then Exit Sub
    If editingTextBox Is Nothing Then Exit Sub
    'Save textbox to label
    Let editingLabel.caption = Space(1) & editingTextBox.text
    'Compare with origin shortcut
    Call checkEditedShortcut( _
        key:=Replace(editingTextBox.Tag, LINE_TAG, vbNullString) _
        , value:=editingTextBox.text _
    )
    Call xxx
    'Hide display
    Let editingTextBox.visible = False
    Let editingLabel.visible = True
    'Clean editing
    Set editingTextBox = Nothing
    Set editingLabel = Nothing
End Sub

Private Sub checkEditedShortcut(ByRef key As String, ByRef value As String)
    Dim keybind As String
    Dim lineIndex As String
    Dim shortcut As String
    For Each row In shortcutTb.ListRows
        Let keybind = row.Range(1, shortcutC.getCustomKeybindColumn())
        Let shortcut = shortcutC.convertCodeToName(keybind)
        Let lineIndex = row.Range(1, shortcutC.getNoColumn())
        If lineIndex = key Then
            Let editedArr(lineIndex - 1) = vbNullString
            If value <> shortcut Then
                Let editedArr(lineIndex - 1) = value
            End If
        Exit For ' row loop
        End If
    Next row
End Sub

' Todo: create system update collection.exist

Private Sub xxx()
    Dim i As Integer
    For i = LBound(editedArr) To UBound(editedArr)
        Debug.Print i & ": " & editedArr(i)
    Next i
End Sub

Private Sub showPicking()
    'Highlight line
    For Each ctrl In Me.KeyboardFrame.controls
        With ctrl
        If _
            TypeOf ctrl Is MSForms.label _
            And pickingIndex = Replace(.Tag, LINE_TAG, vbNullString) _
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

Private Sub hidePicking()
    For Each ctrl In Me.KeyboardFrame.controls
        With ctrl
        If _
            TypeOf ctrl Is MSForms.label _
            And pickingIndex = Replace(.Tag, LINE_TAG, vbNullString) _
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
    Let Me.SourceLabel_0.visible = False
End Sub

Public Sub closeForm()
    Call Unload(Me)
    End
End Sub

Private Sub cleanUp()
    ' Clear Objects
    Set info = Nothing
    Set eventColl = Nothing
    Set editingLabel = Nothing
    Set editingTextBox = Nothing
    Set pickingLabel = Nothing
    Set shortcutTb = Nothing
    Set ctrl = Nothing
    Set row = Nothing
    Set shortcutC = Nothing
    ' Clear Arrays
    Erase editedArr
End Sub


