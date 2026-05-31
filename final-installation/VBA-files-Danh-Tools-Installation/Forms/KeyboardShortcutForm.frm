VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} KeyboardShortcutForm 
   Caption         =   "Settings"
   ClientHeight    =   5607
   ClientLeft      =   105
   ClientTop       =   392
   ClientWidth     =   8113
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
'Declare Variables
Private userResponse As VbMsgBoxResult
Private info As InfoConstants
Private ctrl As MSForms.control
Private coll As Collection
Private editingLabel As MSForms.label
Private editingTextBox As MSForms.textBox
Private editingIndex As String 'Find index by replace so use String is better
Private editingShortcut As String
Private orgiginShortcut As String
Private pickingLabel As MSForms.label
Private pickingIndex As String 'Find index by replace so use String is better
Private Const FILTER_PLACEHOLDER As String = "<Type to filter text>"
Private Const TITLE_TAG As String = "title"
Private Const TITLE_HOVER_COLOR As Long = 16378841 'RGB(229, 243, 255)
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
Private Const LINE_HOVER_COLOR As Long = 16774117 'RGB(217, 235, 249)
Private Const LINE_SELECTED_COLOR As Long = 16772040 'RGB(200, 235, 255)
Private Const LINE_BORDER_COLOR As Long = 14935011 'RGB(227, 227, 227)
Private Const LINE_HEIGHT As Byte = 13.5
Private Const LINE_FONT_SIZE As Byte = 8.5
Private Const MAX_LINE As Byte = 15
Private Const PROG_ID_LABEL As String = "Forms.Label.1"
Private Const PROG_ID_TEXTBOX As String = "Forms.Textbox.1"
Private Const ASTERISK As String = "*"
'EVENTS

Private Sub KeyboardFrameContainer_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Call InitLabel
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Call InitLabel
End Sub

Private Sub UserForm_Deactivate()
    Call InitLabel
End Sub

Private Sub UserForm_Initialize()
    Set info = New InfoConstants
    Call InvisiblePattern
    Call InitComboBox
    Call InitLine
    Call InitLabel
    Call StoreCustomEvent
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
    Set info = Nothing
    Set coll = Nothing
End Sub

Private Sub UserForm_Terminate()
    Call MouseScroll.DisableMouseScroll(Me) ' Remove the mousewheel scrolling
    Set info = Nothing
    Set coll = Nothing
End Sub

Private Sub ApplyAndCloseButton_Click()
    Call KeyboardShortcutForm.CloseForm
End Sub

Private Sub CancelButton_Click()
    Call KeyboardShortcutForm.CloseForm
End Sub

Private Sub FilterPlaceTextBox_Enter()
    'Place Holder Handel
    If FilterPlaceTextBox.Value = FILTER_PLACEHOLDER Then
        FilterPlaceTextBox.Value = vbNullString
        FilterPlaceTextBox.foreColor = vbWindowText
    End If
End Sub

Private Sub FilterPlaceTextBox_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    'Place Holder Handel
    If FilterPlaceTextBox.Value = vbNullString Then
        FilterPlaceTextBox.Value = FILTER_PLACEHOLDER
        FilterPlaceTextBox.foreColor = vbButtonShadow
    End If
End Sub

' ACCESSORS

' FUNCTIONS

Private Sub InitComboBox()
    With ProfilesComboBox
        .Style = fmStyleDropDownList
        .AddItem "Default"
        .ListIndex = "0" '0 Or False
    End With
End Sub
    
Private Sub InitLine()
    Dim shortcutC As ShortcutController: Set shortcutC = New ShortcutController
    Dim shortcutTb As ListObject: Set shortcutTb = shortcutC.getShortcutTable()
    Dim row As ListRow
    Dim keybind As String
    Dim keybindDefault As String
    Dim defaultMark As String
    For Each row In shortcutTb.ListRows
        Let keybind = row.Range(1, shortcutC.getCustomKeybindColumn())
        Let keybindDefault = row.Range(1, shortcutC.getDefaultKeybindColumn())
        Let defaultMark = IIf(keybind = keybindDefault, " *", vbNullString)
        Call CreateLine( _
            index:=row.index _
            , command:=row.Range(1, shortcutC.getCommandColumn()) _
            , keybinding:=shortcutC.convertCodeToName(keybind) _
            , when:=row.Range(1, shortcutC.getWhenColumn()) _
            , Source:=row.Range(1, shortcutC.getSourceCoulmn()) & defaultMark _
        )
    Next row
    Call CreateScrollBar(shortcutTb.DataBodyRange.Rows.Count) ' Max Row
    Set shortcutC = Nothing
    Set shortcutTb = Nothing
End Sub

Private Sub CreateScrollBar(ByRef totalLines As Long)
    If totalLines > MAX_LINE Then
        Me.KeyboardFrame.ScrollBars = fmScrollBarsVertical
        Me.KeyboardFrame.ScrollHeight = totalLines * LINE_HEIGHT
    Else
        Me.KeyboardFrame.ScrollBars = fmScrollBarsNone
    End If
End Sub

Private Sub CreateLine( _
    ByRef index As Long _
    , ByRef command As String _
    , ByRef keybinding As String _
    , ByRef when As String _
    , ByRef Source As String _
)
    Call CreateLineLabel( _
        name:=COMMAND_LABEL & index _
        , index:=index _
        , caption:=command _
        , width:=CommandLabel_0.width _
        , height:=CommandLabel_0.height _
        , top:=CommandLabel_0.top _
        , left:=CommandLabel_0.left _
    )
    ' TextBox need create first for it can be back of label
    Call CreateLineTextBox( _
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
    Call CreateLineLabel( _
        name:=KEYBINDING_LABEL & index _
        , index:=index _
        , caption:=keybinding _
        , width:=KeyBindingLabel_0.width _
        , height:=KeyBindingLabel_0.height _
        , top:=KeyBindingLabel_0.top _
        , left:=KeyBindingLabel_0.left _
    )
    Call CreateLineLabel( _
        name:=WHEN_LABEL & index _
        , index:=index _
        , caption:=when _
        , width:=WhenLabel_0.width _
        , height:=WhenLabel_0.height _
        , top:=WhenLabel_0.top _
        , left:=WhenLabel_0.left _
    )
    Call CreateLineLabel( _
        name:=SOURCE_LABEL & index _
        , index:=index _
        , caption:=Source _
        , width:=SourceLabel_0.width _
        , height:=SourceLabel_0.height _
        , top:=SourceLabel_0.top _
        , left:=SourceLabel_0.left _
    )
    Call CreateLineLabel( _
        name:=LINE_1 & index _
        , index:=index _
        , width:=Line1_0.width _
        , height:=Line1_0.height _
        , top:=Line1_0.top _
        , left:=Line1_0.left _
        , hasBorder:=fmBorderStyleSingle _
    )
    Call CreateLineLabel( _
        name:=LINE_2 & index _
        , index:=index _
        , width:=Line2_0.width _
        , height:=Line2_0.height _
        , top:=Line2_0.top _
        , left:=Line2_0.left _
        , hasBorder:=fmBorderStyleSingle _
    )
    Call CreateLineLabel( _
        name:=LINE_3 & index _
        , index:=index _
        , width:=Line3_0.width _
        , height:=Line3_0.height _
        , top:=Line3_0.top _
        , left:=Line3_0.left _
        , hasBorder:=fmBorderStyleSingle _
    )
    Call CreateLineLabel( _
        name:=LINE_4 & index _
        , index:=index _
        , width:=Line4_0.width _
        , height:=Line4_0.height _
        , top:=Line4_0.top _
        , left:=Line4_0.left _
        , hasBorder:=fmBorderStyleSingle _
    )
End Sub

Private Sub CreateLineLabel( _
    ByRef name As String _
    , ByRef index As Long _
    , ByRef width As Single _
    , ByRef height As Single _
    , ByRef top As Single _
    , ByRef left As Single _
    , Optional ByRef caption As String = vbNullString _
    , Optional ByRef backColor As String = vbWindowBackground _
    , Optional ByRef foreColor As String = vbWindowText _
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
    .bordercolor = LINE_BORDER_COLOR
    .Font.size = LINE_FONT_SIZE
    End With
End Sub

Private Sub CreateLineTextBox( _
    ByRef name As String _
    , ByRef index As Long _
    , ByRef width As Single _
    , ByRef height As Single _
    , ByRef top As Single _
    , ByRef left As Single _
    , Optional ByRef backColor As String = vbWindowBackground _
    , Optional ByRef foreColor As String = vbWindowText _
    , Optional ByRef hasBorder As Byte = fmBorderStyleNone _
    , Optional ByRef effect As Byte = fmSpecialEffectFlat _
    , Optional ByRef visible As Boolean = True _
    , Optional ByRef italic As Boolean = False _
    , Optional ByRef bold As Boolean = False _
)
    Dim lineTextbox As MSForms.textBox
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
    Let .bordercolor = LINE_BORDER_COLOR
    Let .Font.size = LINE_FONT_SIZE
    Let .FontItalic = italic
    Let .FontBold = bold
    Let .SpecialEffect = effect
    Let .visible = visible
    End With
End Sub

'INIT

Private Sub InitLabel()
    ' Loop thourgh all controls for find suitable
    For Each ctrl In Me.KeyboardFrameContainer.controls
        With ctrl
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
                And .backColor <> vbWindowBackground _
            Then
                Let .backColor = vbWindowBackground
            End If
            'Keybind reset
            If _
                .name Like (KEYBINDING_LABEL & ASTERISK) _
                And .foreColor <> vbMenuText _
            Then
                Let .foreColor = vbMenuText
                Let .FontItalic = False
                Let .FontBold = False
            End If
        End If
        End With
    Next ctrl
End Sub

Private Sub AddLabelEvent(ByRef ctrl As MSForms.control)
    With ctrl
    If .Tag = TITLE_TAG Then coll.add createLabelEvent(ctrl)
    If .Tag Like (LINE_TAG & ASTERISK) Then coll.add createLabelEvent(ctrl)
    End With
End Sub

Private Sub AddTextBoxEvent(ByRef ctrl As MSForms.control)
    With ctrl
    If .Tag Like (LINE_TAG & ASTERISK) Then coll.add createTextBoxEvent(ctrl)
    End With
End Sub

Private Sub StoreCustomEvent()
    'Must store control with event inside a collection by adding functions return custom class in that collection.
    Set coll = New Collection
    ' Loop thourgh all controls for find suitable
    For Each ctrl In Me.KeyboardFrameContainer.controls
        If TypeOf ctrl Is MSForms.label Then Call AddLabelEvent(ctrl)
        If TypeOf ctrl Is MSForms.textBox Then Call AddTextBoxEvent(ctrl)
    Next ctrl
End Sub

' CUSTOM EVENTS

Private Function createLabelEvent(ByRef ctrl As MSForms.control) As CustomLabelEvent
    Dim labelE As CustomLabelEvent: Set labelE = New CustomLabelEvent
    Set labelE.setLabel = ctrl
    Set createLabelEvent = labelE
End Function

Private Function createTextBoxEvent(ByRef ctrl As MSForms.control) As CustomTextBoxEvent
    Dim textBoxE As CustomTextBoxEvent: Set textBoxE = New CustomTextBoxEvent
    Set textBoxE.setTextBox = ctrl
    Set createTextBoxEvent = textBoxE
End Function

Public Sub LabelMoveOn(ByRef label As MSForms.label)
    Dim keybindingLabel As MSForms.label
    Call InitLabel
    ' Titles Hover
    If label.Tag = TITLE_TAG Then
        Let label.backColor = TITLE_HOVER_COLOR
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
                If .backColor <> LINE_HOVER_COLOR Then
                    Let .backColor = LINE_HOVER_COLOR
                End If
            End If
            End With
        Next ctrl
        ' Highlight label
        Set keybindingLabel = Me.KeyboardFrame.controls(KEYBINDING_LABEL & Replace(label.Tag, LINE_TAG, vbNullString))
        With keybindingLabel
        If _
            Not label Is pickingLabel _
            And .foreColor <> vbHighlight _
        Then
            Let .foreColor = vbHighlight
            Let .FontItalic = True
            Let .FontBold = True
        End If
        End With
    End If
End Sub

Public Sub LabelClick(ByRef label As MSForms.label)
    With label
    ' Titles Click
    If .Tag = TITLE_TAG Then
        MsgBox ("TODO: Sort By" & .caption)
    ' Lines Click
    ElseIf .Tag Like (LINE_TAG & ASTERISK) Then
        Call DisplayPicking(False)
        Call DisplayEditing(False)
         ' Assign picked Label
        Let pickingIndex = Replace(label.Tag, LINE_TAG, vbNullString)
        Set pickingLabel = Me.KeyboardFrame.controls(KEYBINDING_LABEL & pickingIndex)
        Call DisplayPicking(True)
    End If
    End With
End Sub

Public Sub LabelDbClick(ByRef label As MSForms.label)
    If label.Tag Like (LINE_TAG & ASTERISK) Then
        Call DisplayEditing(False)
        ' Assign editing shortcut
        Let editingIndex = CInt(Replace(label.Tag, LINE_TAG, vbNullString))
        Set editingLabel = Me.KeyboardFrame.controls(KEYBINDING_LABEL & editingIndex)
        Set editingTextBox = Me.KeyboardFrame.controls(KEYBINDING_TEXTBOX & editingIndex)
        Call DisplayEditing(True)
    End If
End Sub

Public Sub TextBoxKeyDown( _
    ByRef textBox As MSForms.textBox _
    , ByRef keycode As MSForms.ReturnInteger _
    , ByRef Shift As Integer _
)
    Dim beforeEditShortcut As String: Let beforeEditShortcut = textBox.text
    Dim shortcutC As ShortcutController: Set shortcutC = New ShortcutController
    Let editingShortcut = shortcutC.convertKeyToName(keycode, Shift)
    If editingShortcut = shortcutC.getEnterKey() Then
        'TODO: tonight
        Call DisplayEditing(False)
    ElseIf editingShortcut = shortcutC.getEscKey() Then
        Let textBox.text = beforeEditShortcut
        Call DisplayEditing(False)
    Else
        Let textBox.text = editingShortcut
    End If
    'Prevent default keyDown
    Let keycode = 0
    Set shortcutC = Nothing
End Sub

Public Sub TextBoxChange(ByRef textBox As MSForms.textBox)
    Let editingLabel.caption = Space(1) & textBox.text
    ' Call DisplayEditing(False)
End Sub

Private Sub DisplayEditing(isShow As Boolean)
    ' Check for the 1st time
    If editingTextBox Is Nothing Then Exit Sub
    'Toggle display
    Let editingTextBox.visible = isShow
    Let editingLabel.visible = Not isShow
    If isShow Then
        editingTextBox.SetFocus
    Else
        Set editingTextBox = Nothing
        Set editingLabel = Nothing
    End If
End Sub

Private Sub DisplayPicking(isShow As Boolean)
    ' Check for the 1st time
    If pickingLabel Is Nothing Then Exit Sub
    'Highlight line
    For Each ctrl In Me.KeyboardFrame.controls
        With ctrl
        If _
            TypeOf ctrl Is MSForms.label _
            And pickingIndex = Replace(.Tag, LINE_TAG, vbNullString) _
        Then
            If isShow Then
                'Highlight line
                Let .backColor = LINE_SELECTED_COLOR
                'Highlight label
                Let .foreColor = vbHighlight
                Let .FontItalic = False
                Let .FontBold = True
            Else
                'Highlight line
                Let .backColor = vbWindowBackground
                'Highlight label
                Let .foreColor = vbMenuText
                Let .FontItalic = False
                Let .FontBold = False
            End If
        End If
        End With
    Next ctrl
End Sub

' CLEAN
Private Sub InvisiblePattern()
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

Public Sub CloseForm()
    Unload Me
    End
End Sub

