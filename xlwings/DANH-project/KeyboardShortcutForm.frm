VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} KeyboardShortcutForm 
   Caption         =   "Settings"
   ClientHeight    =   7686
   ClientLeft      =   5488
   ClientTop       =   4011
   ClientWidth     =   10934
   OleObjectBlob   =   "KeyboardShortcutForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "KeyboardShortcutForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Check README.md for more information
Option Explicit
' Declare Variables
' NOTE: 'Find indexes by replace so use String as variable is better
Private userResponse As VbMsgBoxResult
Private info As InfoConstants
Private shortcutTb As ListObject
Private eventColl As Collection
Private hoverLabel As MsForms.label
Private hoverIndex As String
Private editingLabel As MsForms.label
Private editingTextBox As MsForms.textBox
Private editingIndex As String
Private editingShortcut As String
Private editedArr() As String ' TODO: Create custom array CRUD
Private applyArr() As String
Private pickingLabel As MsForms.label
Private pickingIndex As String
Private shortcutC As ShortcutController
Private maxRow As Long
Private Enum COLOR
    title_hover = 16378841 'RGB(229, 243, 255)
    line_hover = 16774117 'RGB(217, 235, 249)
    line_hover_edited = 51450 'RGB(250, 200, 0)
    line_picking = 16772040 'RGB(200, 235, 255)
    line_picking_edited = 51450 'RGB(250, 200, 0)
    line_border = 14935011 'RGB(227, 227, 227)
    line_edited = 61690 'RGB(250, 240, 0)
    line_duplicated = 987135 'RGB(255,15,15)
    ' Default Variables
    highlight = vbHighlight
    window_text = vbWindowText
    button_shadow = vbButtonShadow
    window_background = vbWindowBackground
    menu_text = vbMenuText
    menu_bar = vbMenuBar
End Enum
Private Enum FORM_POSITION
    manual = 0
    center_owner = 1
    center_screen = 2
    windows_default = 3
    height = 410
    width = 550
    top = height / 2
    left = width / 2
End Enum
Private Enum DIRECTION
    up = -1
    down = 1
End Enum
Private ENUM MASK
    none = 0
    shiftKey =  1
    ctrlKey = 2
    altKey = 4
End Enum
Private Const DEFAULT = "<Default>"
Private Const MODIFIED = "<Modified>"
Private Const DUPLICATED = "<Duplicated>"
' TODO: MAKE instruction constants class
Private Const INSTRUCTION_PICKING = "Please pick a line to edit."
Private Const INSTRUCTION_EDITING = "Click Edit button or double click a line to modify."
Private Const INSTRUCTION_MODIFY = "Press desired key combination and then press ENTER."
Private Const FILTER_PLACEHOLDER As String = "<Type to filter text>"
Private Const TITLE_TAG As String = "title_"
Private Const LINE_TAG As String = "line_"
Private Const OVERLAY_TAG As String = "overlay_"
Private Const OVERLAY_FORM As String = "outer"
Private Const OVERLAY_PAGE As String = "inner"
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
Private Const ZERO As String = "0"
Private Const BACKGROUND As String = "BackColor"
Private Const FORE As String = "ForeColor"
Private Const ITALIC As String = "FontItalic"
Private Const BOLD As String = "FontBold"
Private Const EDIT_CAPTION As String = "Edit"
' Loop iterators
Private ctrl As MsForms.control
Private row As ListRow

' NOTE: Used to use Mutators/Accessors as Public for fix bug when cls form as instance but not working

' MUTATORS

Private Sub letUserResponse(ByRef value As VbMsgBoxResult): Let userResponse = value: End Sub
Private Sub setInfo(ByRef value As InfoConstants): Set info = value: End Sub
Private Sub setEventColl(ByRef value As Collection): Set eventColl = value: End Sub
Private Sub setHoverLabel(ByRef value As MsForms.label): Set hoverLabel = value: End Sub
Private Sub letHoverIndex(ByRef value As String): Let hoverIndex = value: End Sub
Private Sub setEditingLabel(ByRef value As MsForms.label): Set editingLabel = value: End Sub
Private Sub setEditingTextBox(ByRef value As MsForms.textBox): Set editingTextBox = value: End Sub
Private Sub letEditingIndex(ByRef value As String): Let editingIndex = value: End Sub
Private Sub letEditingShortcut(ByRef value As String): Let editingShortcut = value: End Sub
Private Sub setPickingLabel(ByRef value As MsForms.label): Set pickingLabel = value: End Sub
Private Sub letPickingIndex(ByRef value As String): Let pickingIndex = value: End Sub
Private Sub setShortcutC(ByRef value As ShortcutController): Set shortcutC = value: End Sub
Private Sub letMaxRow(ByRef value As Long): Let maxRow = value: End Sub

Private Sub setInstruction(ByRef text As String)
    Let AsteriskInstructionLabel.Caption = ASTERISK & Space(1) & text
End Sub

Private Sub setCaption(ByRef ctrl AS MsForms.Control,ByRef text As String)
    Let ctrl.Caption = text
End Sub

' ACCESSORS

Private Function getUserResponse() As VbMsgBoxResult: Let getUserResponse = userResponse: End Function
Private Function getInfo() As InfoConstants: Set getInfo = info: End Function
Private Function getEventColl() As Collection: Set getEventColl = eventColl: End Function
Private Function getHoverLabel() As MsForms.label: Set getHoverLabel = hoverLabel: End Function
Private Function getHoverIndex() As String: Let getHoverIndex = hoverIndex: End Function
Private Function getEditingLabel() As MsForms.label: Set getEditingLabel = editingLabel: End Function
Private Function getEditingTextBox() As MsForms.textBox: Set getEditingTextBox = editingTextBox: End Function
Private Function getEditingIndex() As String: Let getEditingIndex = editingIndex: End Function
Private Function getEditingShortcut() As String: Let getEditingShortcut = editingShortcut: End Function
Private Function getPickingLabel() As MsForms.label: Set getPickingLabel = pickingLabel: End Function
Private Function getPickingIndex() As String: Let getPickingIndex = pickingIndex: End Function
Private Function getShortcutC() As ShortcutController: Set getShortcutC = shortcutC: End Function
Private Function getMaxRow() As Long: Let getMaxRow = maxRow: End Function

' NOTE: Can get line index of both label and textbox
Private Function getLineIndex(ByRef ctrl As MsForms.control) As String
    Let getLineIndex = Replace(ctrl.Tag, LINE_TAG, vbNullString)
End Function

' CHECKS

Private Function isLabel(ByRef ctrl As MsForms.control) As Boolean
    Let isLabel = (TypeOf ctrl Is MsForms.label)
End Function

Private Function isTextBox(ByRef ctrl As MsForms.control) As Boolean
    Let isTextBox = (TypeOf ctrl Is MsForms.textBox)
End Function

Private Function isTitle(ByRef label As MsForms.label) As Boolean
    Let isTitle = (label.Tag Like (TITLE_TAG & ASTERISK))
End Function

Private Function isLine(ByRef ctrl As MsForms.control) As Boolean
    If (isLabel(ctrl) Or isTextBox(ctrl)) Then
        Let isLine = (ctrl.Tag Like (LINE_TAG & ASTERISK))
    Else
        Let isLine = False
    End If
End Function

Private Function isSameLine( _
    ByRef labelFirst As MsForms.label, _
    ByRef labelSecond As MsForms.label _
) As Boolean
    If (labelFirst Is Nothing) Or (labelSecond Is Nothing) Then
        Let isSameLine = False
    Else
        Let isSameLine = (labelFirst.Tag = labelSecond.Tag)
    End If
End Function

Private Function isKeyBinding(ByRef label As MsForms.label) As Boolean
    Let isKeyBinding = (label.name Like (KEYBINDING_LABEL & ASTERISK))
End Function

Private Function isHoverLabel(ByRef label As MsForms.label) As Boolean
    Let isHoverLabel = (label Is getHoverLabel())
End Function

Private Function isHoverLine(ByRef lineIndex As String) As Boolean
    Let isHoverLine = (lineIndex = getHoverIndex())
End Function

Private Function isPickingLabel(ByRef label As MsForms.label) As Boolean
    Let isPickingLabel = (label Is getPickingLabel())
End Function

Private Function isPickingLine(ByRef lineIndex As String) As Boolean
    Let isPickingLine = (lineIndex = getPickingIndex())
End Function

Private Function isEditedLine(ByRef lineIndex As String) As Boolean
    Dim i As Long
    For i = LBound(editedArr) To UBound(editedArr)
        If _
            (lineIndex = CStr(i + 1)) _
            And (editedArr(i) <> DEFAULT) _
        Then
            Let isEditedLine = True
            Exit Function 'Stop if found
        End If
    Next i
End Function

Private Function isDuplicatedLine(ByRef lineIndex As String) As Boolean
    Dim i As Long
    Dim checkIndex As Long: Let checkIndex = CLng(lineIndex) - 1
    For i = LBound(applyArr) To UBound(applyArr)
        If _
            (i <> checkIndex) _
            And (applyArr(i) <> getShortcutC().getNoSet()) _
            And ( _
                (applyArr(i) = applyArr(checkIndex)) _
                Or (InStr(applyArr(checkIndex), getShortcutC().getUnknown()) > 0) _
            ) _
        Then
            Let isDuplicatedLine = True
            Exit Function 'Stop if found
        End If
    Next i
End Function

Private Function hasDuplicated() As Boolean
    Dim i As Long
    For i = LBound(applyArr) To UBound(applyArr)
        If isDuplicatedLine(CStr(i + 1)) Then
            Let hasDuplicated = True
            Exit Function
        End If
    Next i
End Function

Private Function isPicking() As Boolean
    Let isPicking = Not (getPickingLabel() Is Nothing)
End Function

Private Function isEditing() As Boolean
    Let isEditing = Not (getEditingLabel() Is Nothing Or getEditingTextBox() Is Nothing)
End Function

Private Function isOverlay(ByRef label As MsForms.label) As Boolean
    Let isOverlay = (label.Tag Like (OVERLAY_TAG & ASTERISK))
End Function

' NOTE: This function can improve performance very much
Private Function canUpdateFormat( _
    ByRef label As MsForms.label _
    , ByRef propName As String _
    , ByRef format As Long _
) As Boolean
    If CallByName(label, propName, VbGet) <> format Then
        Call CallByName(label, propName, VbLet, format)
        Let canUpdateFormat = True
    Else
        Let canUpdateFormat = False
    End If
End Function

Private Sub markHoverTitle(ByRef label As MsForms.label)
    Call canUpdateFormat(label, BACKGROUND, COLOR.title_hover)
End Sub

Private Sub cleanMarkTitle(ByRef label As MsForms.label)
    Call canUpdateFormat(label, BACKGROUND, COLOR.menu_bar)
End Sub

Private Sub markPickingLine(ByRef label As MsForms.label)
    Call canUpdateFormat(label, BACKGROUND, COLOR.line_picking)
    Call canUpdateFormat(label, FORE, COLOR.highlight)
    Call canUpdateFormat(label, italic, True)
    Call canUpdateFormat(label, bold, True)
End Sub

Private Sub markEditedLine(ByRef label As MsForms.label)
    Call canUpdateFormat(label, BACKGROUND, COLOR.line_edited)
    Call canUpdateFormat(label, FORE, COLOR.menu_text)
    Call canUpdateFormat(label, italic, False)
    Call canUpdateFormat(label, bold, False)
End Sub

Private Sub markDuplicatedLine(ByRef label As MsForms.label)
    Call canUpdateFormat(label, BACKGROUND, COLOR.window_background)
    Call canUpdateFormat(label, FORE, COLOR.line_duplicated)
    Call canUpdateFormat(label, italic, False)
    Call CanUpdateFormat(label, bold, True)
End SUb

Private Sub markPickingEditedLine(ByRef label As MsForms.label)
    Call canUpdateFormat(label, BACKGROUND, COLOR.line_picking_edited)
    Call canUpdateFormat(label, FORE, COLOR.highlight)
    Call canUpdateFormat(label, italic, True)
    Call canUpdateFormat(label, bold, True)
End Sub

Private Sub markPickingDuplicatedLine(ByRef label As MsForms.label)
    Call canUpdateFormat(label, BACKGROUND, COLOR.line_picking)
    Call canUpdateFormat(label, FORE, COLOR.line_duplicated)
    Call canUpdateFormat(label, italic, True)
    Call canUpdateFormat(label, bold, True)
End Sub

Private Sub markPickingEditedDuplicatedLine(ByRef label As MsForms.label)
    Call canUpdateFormat(label, BACKGROUND, COLOR.line_picking_edited)
    Call canUpdateFormat(label, FORE, COLOR.line_duplicated)
    Call canUpdateFormat(label, italic, True)
    Call canUpdateFormat(label, bold, True)
End Sub

Private Sub markEditedDuplicatedLine(ByRef label As MsForms.label)
    Call canUpdateFormat(label, BACKGROUND, COLOR.line_edited)
    Call canUpdateFormat(label, FORE, COLOR.line_duplicated)
    Call canUpdateFormat(label, italic, False)
    Call canUpdateFormat(label, bold, True)
End Sub

Private Sub markHoverLine(ByRef label As MsForms.label)
    Call canUpdateFormat(label, BACKGROUND, COLOR.line_hover)
End Sub

Private Sub markHoverEditedLine(ByRef label As MsForms.label)
    Call canUpdateFormat(label, BACKGROUND, COLOR.line_hover_edited)
End Sub

Private Sub cleanMarkLine(ByRef label As MsForms.label)
    Call canUpdateFormat(label, BACKGROUND, COLOR.window_background)
    Call canUpdateFormat(label, FORE, COLOR.menu_text)
    Call canUpdateFormat(label, italic, False)
    Call canUpdateFormat(label, bold, False)
End Sub

Private Sub markKeybindingDuplicated(ByRef label As MsForms.label)
    Call canUpdateFormat(label,  FORE, COLOR.line_duplicated)
    Call canUpdateFormat(label, italic, True)
    Call canUpdateFormat(label, bold, True)
End Sub

Private Sub markKeybindingPicking(ByRef label As MsForms.label)
    Call canUpdateFormat(label, FORE, COLOR.highlight)
    Call canUpdateFormat(label, italic, False)
    Call canUpdateFormat(label, bold, True)
End Sub

Private Sub markKeybindingHover(ByRef label As MsForms.label)
    Call canUpdateFormat(label, FORE, COLOR.highlight)
    Call canUpdateFormat(label, italic, True)
    Call canUpdateFormat(label, bold, True)
End Sub

Private Sub cleanMarkKeybinding(ByRef label As MsForms.label)
    Call canUpdateFormat(label, FORE, COLOR.menu_text)
    Call canUpdateFormat(label, italic, False)
    Call canUpdateFormat(label, bold, False)
End Sub

' CONSTRUCTOR

Private Sub UserForm_Initialize()
    Call setShortcutC(New ShortcutController)
    Call setInfo(New InfoConstants)
    Call initUserForm
    Call hidePattern
    Call initComboBox
    Call initButton
    Call initRow
    Call updateInstruction
    ' (REJECT: Use exit event seem better solution but keep source 4 future reuse)
    ' Call initOverlay
    Call storeCustomEvent
'    Call InitTabIndexes( _
'        , FindWhatLabel _
'        , FindAreaInput _
'        , ReplaceWithLabel _
'        , ReplaceAreaInput _
'        , WithinLabel _
'        , sub WithInComboBox _
'        , SelectedAreaInput _
'        , SearchLabel _
'        , SearchComboBox _
'        , MatchCaseCheckBox _
'        , MatchByteCheckBox _
'        , MatchContentCheckBox _
'        , LengthOrderCheckBox _
'        , ReplaceAllButton _
'        , CloseButton _
'    )
    ' Apply the mousewheel scrolling (NOTE: Disable scroll zoom)
    Call MouseScroll.EnableMouseScroll( _
        uForm:= Me _
        , passScrollToParentAtMargins:= True _
        , useShiftForPerpendicularScroll:= False _
        , useCtrlToZoom:= False _
    )
End Sub

' DESTRUCTOR

Private Sub UserForm_Terminate()
    Call MouseScroll.DisableMouseScroll(Me) ' Remove the mousewheel scrolling
    Call clearUp
End Sub

' EVENTS

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Call resetHover
    Call formatLabel
End Sub

Private Sub UserForm_Deactivate()
    Call resetHover
    Call formatLabel
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Call MouseScroll.DisableMouseScroll(Me) ' Remove the mousewheel scrolling
    Call clearUp
End Sub

' Cancel Picking And Editing if click outside

Private Sub UserForm_Click()
    Call hidePickingAndEditing(isChange:=False)
End Sub

Private Sub MultiPage_Change()
    Call hidePickingAndEditing(isChange:=False)
End Sub

Private Sub MultiPage_Click(ByVal Index As Long)
    Call hidePickingAndEditing(isChange:=False)
End Sub

Private Sub MultiPage_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Call hideEditing(isChange:=False)
End Sub

Private Sub KeyboardFrameContainer_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Call hideEditing(isChange:=true)
End Sub

Private Sub KeyboardFrame_KeyDown(ByVal KeyCode As MsForms.ReturnInteger, ByVal Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn: If Shift = MASK.none Then Call showEditing
        Case vbKeyEscape: If Shift = MASK.none Then Call hidePickingAndEditing(isChange:=False)
        Case vbKeyUp: IF Shift = MASK.none Then Call movePicking(DIRECTION.up)
        Case vbKeyDown: IF Shift = MASK.none Then Call movePicking(DIRECTION.down)
    End Select
End Sub

Private Sub KeyboardFrameContainer_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Call resetHover
    Call formatLabel
End Sub

Private Sub ApplyAndCloseButton_Click()
    'TODO save
    Call KeyboardShortcutForm.closeForm
End Sub

Private Sub CancelButton_Click()
    ' TODO Ask for confirmation
    Call KeyboardShortcutForm.closeForm
End Sub

Private Sub EditButton_Click()
    If isPicking Then Call showEditing
End Sub

Private Sub ApplyButton_Click()
    Dim i As Long
    Dim applyCodeArr() As String
    IF isEditing Then Call hideEditing(isChange:=True)
    If  hasDuplicated() Then
    ' TODO: Create constant for message
    Call letUserResponse(MsgBox( _
        getInfo().getPrompt & "There are still duplications or errors, please check again !!!", _
        vbOKOnly + vbExclamation, _
        getInfo().getAuthor))
        Exit Sub
    End If
    ' 2D Array start at 1 will store the converted code
    ReDim applyCodeArr(1 To UBound(applyArr) + 1, 1 To 1)
    For i = LBound(applyArr) To UBound(applyArr)
        Let applyCodeArr(i + 1, 1) = getShortcutC().convertNameToCode(applyArr(i))
    Next i
    Call getShortcutC().unInstall
    Call getShortcutC().setColData(applyCodeArr, getShortcutC().getCustomKeybindingCol())
    Call getShortcutC().install
    Call formatLabel
End Sub

Private Sub RestoreDefaultButton_Click()
    'TEST

End Sub

Private Sub FilterPlaceTextBox_Enter()
    'Place Holder Handel
    With FilterPlaceTextBox
    If .value = FILTER_PLACEHOLDER Then
        .value = vbNullString
        .foreColor = COLOR.window_text
    End If
    End With
End Sub

Private Sub FilterPlaceTextBox_Exit(ByVal Cancel As MsForms.ReturnBoolean)
    'Place Holder Handel
    With FilterPlaceTextBox
    If .value = vbNullString Then
        .value = FILTER_PLACEHOLDER
        .foreColor = COLOR.button_shadow
    End If
    End With
End Sub

' Init
Private Sub initUserForm()
    With Me
        .StartUpPosition = FORM_POSITION.manual
        .height = FORM_POSITION.height
        .width = FORM_POSITION.width
        .top = FORM_POSITION.top
        .left = FORM_POSITION.left
    End With
End Sub

Private Sub initComboBox()
    With ProfilesComboBox
        .Style = fmStyleDropDownList
        .AddItem DEFAULT
        .ListIndex = ZERO '0 Or False
    End With
End Sub

Private Sub initButton()
    ' TODO make order button init
    ' Init EditResetButton
    Call setCaption(ctrl:=Me.EditButton, text:=EDIT_CAPTION)
    Let EditButton.enabled = True
End Sub

Private Sub initRow()
    Dim keybinding As String
    Dim keybindingDefault As String
    Dim lineIndex As String
    Dim defaultMark As String
    Dim shortcut As String
    Call letMaxRow(getShortcutC().getRows().Count)
    ReDim editedArr(getMaxRow() - 1)
    ReDim applyArr(getMaxRow() -1)
    Call resetEdited
    For Each row In getShortcutC().getRows()
        Let keybinding = row.Range(1, getShortcutC().getCustomKeybindingCol())
        Let keybindingDefault = row.Range(1, getShortcutC().getDefaultKeybindingCol())
        ' Let lineIndex = row.Range(1, getShortcutC().getNoCol())
        Let lineIndex = row.index
        Let shortcut = getShortcutC().convertCodeToName(keybinding)
        Let defaultMark = IIf(keybinding = keybindingDefault, DEFAULT, vbNullString)
        Call createRow( _
            index:=CLng(lineIndex) _
            , command:=row.Range(1, getShortcutC().getCommandCol()) _
            , shortcut:=shortcut _
            , when:=row.Range(1, getShortcutC().getWhenCol()) _
            , status:=row.Range(1, getShortcutC().getStatusCol()) & defaultMark _
        )
        Let editedArr(lineIndex - 1) = DEFAULT
        Let applyArr(lineIndex - 1) = shortcut
    Next row
    Call createScrollBar(getMaxRow())
End Sub

Private Sub initOverLay()
    ' Outer overlay
    Call createOverlayLabel( _
        parentCtrl:=Me.Controls _
        , width:=Me.InsideWidth _
        , height:=Me.InsideHeight _
        , top:=0 _
        , left:=0 _
        , name:=OVERLAY_FORM _
        , backStyle:=1 _
        , visible:=True _
    )
    ' Inner overlay
    Call createOverlayLabel( _
        parentCtrl:=Me.MultiPage.KeyboardShortcutsPage.Controls _
        , width:=Me.MultiPage.KeyboardShortcutsPage.InsideWidth _
        , height:=Me.MultiPage.KeyboardShortcutsPage.InsideHeight _
        , top:=0 _
        , left:=0 _
        , name:=OVERLAY_PAGE _
        , backStyle:=1 _
        , visible:=True _
    )
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
        , caption:=status _
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
    , Optional ByRef visible As Boolean = True _
)
    Dim lineLabel As MsForms.label
    Set lineLabel = Me.KeyboardFrame.add( _
        bstrProgId:=PROG_ID_LABEL _
        , name:=name _
        , visible:=visible)
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
    .visible = visible
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
        , visible:=visible)
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

Private Sub createOverlayLabel( _
    ByRef parentCtrl As MsForms.controls _
    , ByRef width As Single _
    , ByRef height As Single _
    , ByRef top As Single _
    , ByRef left As Single _
    , ByRef name As String _
    , Optional ByRef caption As String = vbNullString _
    , Optional ByRef backColor As String = COLOR.highlight _
    , Optional ByRef backStyle As String = fmBackStyleTransparent _
    , Optional ByRef hasBorder As Byte = fmBorderStyleNone _
    , Optional ByRef zOrder As Byte = 0 _
    , Optional ByRef visible As Boolean = False _
)
    Dim overlayLabel As MsForms.label
    Set overlayLabel = parentCtrl.add( _
        bstrProgId:=PROG_ID_LABEL _
        , name:=name _
        , visible:= visible) 'init will hide
    With overlayLabel
        .caption = caption
        .Tag = OVERLAY_TAG & name
        .width = width
        .height = height
        .top = top
        .left = left
        .backColor = backColor
        .BackStyle = backStyle
        .BorderStyle = hasBorder
        .visible = visible
        .ZOrder zOrder ' 0: Bring to front
    End With
    Set overlayLabel = Nothing
End Sub

Private Sub addEvent(ByRef ctrl As MSForms.Control)
    If isLabel(ctrl) Then
        If isTitle(ctrl) Then getEventColl().Add createLabelEvent(ctrl)
        If isLine(ctrl) Then getEventColl().Add createLabelEvent(ctrl)
        If isOverlay(ctrl) Then getEventColl().Add createLabelEvent(ctrl)
    ElseIf isTextBox(ctrl) Then
        If isLine(ctrl) Then getEventColl().Add createTextBoxEvent(ctrl)
    End If
End Sub

Private Sub storeCustomEvent()
    'Must store control with event inside a collection by adding functions return custom class in that collection.
    Call setEventColl(New Collection)
    ' Loop though all controls for find suitable
    For Each ctrl In Me.controls
        Call addEvent(ctrl)
    Next ctrl
End Sub

' CUSTOM EVENTS

Private Function createLabelEvent(ByRef label As MsForms.label) As CustomLabelEvent
    Dim labelE As CustomLabelEvent: Set labelE = New CustomLabelEvent
    Set labelE.setLabel = label
    Set createLabelEvent = labelE
    Set labelE = Nothing
End Function

Private Function createTextBoxEvent(ByRef textBox As MsForms.textBox) As CustomTextBoxEvent
    Dim textBoxE As CustomTextBoxEvent: Set textBoxE = New CustomTextBoxEvent
    Set textBoxE.setTextBox = textBox
    Set createTextBoxEvent = textBoxE
    Set textBoxE = Nothing
End Function

Public Sub labelMoveOn(ByRef label As MsForms.label)
    '  Dim keybindingLabel As MsForms.label
    Dim lineIndex As String
    ' Check still in same label do nothing (Performance issues)
    If isSameLine(label, getHoverLabel()) Then Exit Sub
    ' Skip if picking line is hover Line
    If isPickingLine(getLineIndex(label)) Then 
        Call resetHover()
    Else
        Call updateHover(label)
    End If
    Call formatLabel
End Sub

Public Sub labelClick(ByRef label As MsForms.label)
    ' Skip Editing
    If isSameLine(getEditingLabel(), label) Then
        Exit Sub
    ' Click 2 times
    Elseif isSameLine(getPickingLabel(), label) Then
        Call showEditing
        Exit Sub
    ' Titles Click
    ElseIf isTitle(label) Then
        Call hideEditing(isChange:=False)
        MsgBox ("TODO: Sort By" & label.caption)
    ' Lines Click
    ElseIf isLine(label) Then
        Call hidePickingAndEditing(true)
        Call showPicking(label)
    ElseIf isOverlay(label) Then
        Call hidePickingAndEditing(False)
    End If
    ' Update instruction
    Call updateInstruction
End Sub

Public Sub labelDbClick(ByRef label As MsForms.label)
    If Not isLine(label) Then Exit Sub
    Call showEditing
End Sub

Public Sub textBoxKeyDown( _
    ByRef textBox As MsForms.textBox _
    , ByRef KeyCode As MsForms.ReturnInteger _
    , ByRef Shift As Integer _
)
    Call letEditingShortcut(getShortcutC().convertKeyToName(KeyCode, Shift))
    ' Enter press save editing
    If (getEditingShortcut() = getShortcutC().getEnterKey()) Then
        Call hideEditing(isChange:=True)
    ' Esc press cancel editing
    ElseIf (getEditingShortcut() = getShortcutC().getEscKey()) Then
        Call hideEditing(isChange:=False)
    ' Backspace press clear content
    ElseIf getEditingShortcut() = getShortcutC().getBackspaceKey() Then
        Let textBox.text = vbNullString
    ' Assign shortcut keybinding
    Else
        Let textBox.text = getEditingShortcut()
    End If
    'Prevent default keyDown
    Let KeyCode = 0
End Sub

Public Sub textBoxChange(ByRef textBox As MsForms.textBox)
    ' If GetEditingLabel() Is Nothing Then Exit Sub
    ' Let GetEditingLabel().caption = Space(1) & textBox.text
End Sub

Private Sub showEditing()
    ' Check if are editing hide it
    If isEditing Then Call hideEditing(isChange:=False)
    Call letEditingIndex(getLineIndex(getPickingLabel()))
    ' Assign editing shortcut object
    Call setEditingLabel(Me.KeyboardFrame.controls(KEYBINDING_LABEL & getEditingIndex()))
    Call setEditingTextBox(Me.KeyboardFrame.controls(KEYBINDING_TEXTBOX & getEditingIndex()))
    ' If No Set keybinding set textbox to Blank
    Let getEditingTextBox().text = IIf( _
        LTrim(getEditingLabel().caption) = getShortcutC().getNoSet() _
        , vbNullString _
        , LTrim(getEditingLabel().caption) _
    )
    ' Show display
    Call displayTextBox(isDisplay:=True)
    Call getEditingTextBox().SetFocus
    ' Disable Edit Button
    Let EditButton.enabled = False
    Call updateInstruction
End Sub

Private Sub hideEditing(Optional ByRef isChange As Boolean = True)
    ' Check for the 1st time
    If Not isEditing Then Exit Sub
    ' Restore before edit shortcut (No Change)
    If Not isChange Then Let getEditingTextBox().text = LTrim(getEditingLabel().caption)
    ' If textBox blank set label to No Set
    Dim lineIndex As String
    Let getEditingLabel().caption = IIf( _
        getEditingTextBox().text = vbNullString _
        , getShortcutC().getNoSet() _
        , Space(1) & getEditingTextBox().text _
    )
    Let lineIndex = getLineIndex(getEditingTextBox())
    ' Update edited shortcut
    Call updateEdited( _
        key:=lineIndex _
        , value:=LTrim(getEditingLabel().caption) _
    )
    ' Update apply shortcut
    Call updateApply( _
        key:=lineIndex _
        , value:=LTrim(getEditingLabel().caption) _
    )
    ' Highlight Line
    Call formatLabel
    ' Hide display
    Call displayTextBox(isDisplay:=False)
    Call resetEditing
    ' Disable Edit Button
    Let EditButton.enabled = True
    Call updateInstruction
End Sub

Private Sub resetEditing()
    If Not isEditing Then Exit Sub
    Call setEditingTextBox(Nothing)
    Call setEditingLabel(Nothing)
End Sub

Private Sub updateHover(ByRef label As MsForms.label)
    Call letHoverIndex(getLineIndex(label))
    Call setHoverLabel(label)
End Sub

Private Sub resetHover()
    Call letHoverIndex(vbNullString)
    Call setHoverLabel(Nothing)
End Sub

Private Sub showPicking(ByRef label As MsForms.label)
    Call updatePicking(label)
    Call formatLabel
End Sub

Private Sub hidePicking()
    If Not isPicking Then Exit Sub
    Call resetPicking
    Call formatLabel
End Sub

Private Sub updatePicking(ByRef label As MsForms.label)
    Call letPickingIndex(getLineIndex(label))
    ' NOTE: Actually, pickingLabel is not necessary but for future maybe use
    Call setPickingLabel(Me.KeyboardFrame.controls(KEYBINDING_LABEL & getPickingIndex()))
End Sub

Private Sub resetPicking()
    If Not isPicking Then Exit Sub
    Call letPickingIndex(vbNullString)
    Call setPickingLabel(Nothing)
End Sub

Private Sub movePicking(ByRef direction As Integer)
    if Not isPicking Then Exit Sub
    Dim nextIndex As String
    Dim nextLabel As MsForms.label
    ' Update Next picking by index
    Let nextIndex = getPickingIndex() + direction
    If nextIndex > getMaxRow() Then
        Let nextIndex = getMaxRow()
    ElseIf nextIndex < 1 Then ' Min Row
        Let nextIndex = 1
    End If
    Set nextLabel = Me.KeyboardFrame.controls(KEYBINDING_LABEL & nextIndex)
    Call scrollPicking(nextLabel)
    Call showPicking(nextLabel)
    Set nextLabel = Nothing
End Sub

Private Sub scrollPicking(ByRef label As MsForms.label)
    With label
    Dim top as Single: Let top = .top
    Dim bottom as Single: Let bottom = top + LINE_HEIGHT
        'Frame scroll
        With .parent
        Dim scrollTop as Single: Let scrollTop = .ScrollTop
        Dim scrollBottom as Single: Let scrollBottom = scrollTop + .InsideHeight
        ' Scroll down
        If bottom > scrollBottom Then
            ' .ScrollTop = .ScrollTop + (bottom - scrollBottom)
            .ScrollTop = bottom - .InsideHeight
        ' Scroll up
        ElseIf top < scrollTop Then
            .ScrollTop = top
        End If
        End With ' .parent
    End With ' label
End Sub

Private Sub hidePickingAndEditing(Optional ByRef isChange As Boolean = True)
    If isPicking Then Call hidePicking
    If isEditing Then Call hideEditing(isChange)
End Sub

Private Sub updateEdited(ByRef key As String, ByRef value As String)
    Dim keybinding As String
    Dim shortcut As String
    Let keybinding = getShortcutC().getRows()(key).Range(1, getShortcutC().getCustomKeybindingCol())
    Let shortcut = getShortcutC().convertCodeToName(keybinding)
    ' DEFAULT
    If shortcut = value Then
        Let editedArr(key - 1) = DEFAULT
    ' EDITED
    Else
        Let editedArr(key - 1) = value
    End If
End Sub

Private Sub resetEdited()
    Dim i As Long
    For i = LBound(editedArr) To UBound(editedArr)
        editedArr(i) = DEFAULT
    Next i
End Sub

Private Sub updateApply(ByRef key As String, ByRef value As String)
    If applyArr(key - 1) <> value Then applyArr(key - 1) = value
End Sub

Private Sub updateInstruction()
    If Not isPicking Then
        Call setInstruction(INSTRUCTION_PICKING)
    ElseIf Not isEditing Then
        Call setInstruction(INSTRUCTION_EDITING)
    ElseIf isEditing Then
        Call setInstruction(INSTRUCTION_MODIFY)
    End If
End Sub

Private Sub formatLabel()
    Dim lineIndex As String
    ' Loop all line labels
    For Each ctrl In Me.KeyboardFrameContainer.controls
        If Not isLabel(ctrl) Then GoTo NextCtrl ' Continue
        Let lineIndex = getLineIndex(ctrl)
        If lineIndex = ZERO Then GoTo NextCtrl ' Skip Pattern Line
        If isKeyBinding(ctrl) Then
            Select Case True
                ' Highlight Duplicate
                Case isPickingLine(lineIndex)
                    Call markKeybindingDuplicated(ctrl)
                ' Highlight Picking
                Case isPickingLine(lineIndex)
                    Call markKeybindingPicking(ctrl)
                ' Highlight Hover
                Case isHoverLine(lineIndex)
                    Call markKeybindingHover(ctrl)
                ' Clean
                Case Else
                    Call cleanMarkKeybinding(ctrl)
            End Select
        End If
        ' NOTE: VBA Select Case auto break if matching
        ' Label format
        Select Case True
            ' Highlight Title
            Case isTitle(ctrl) And isHoverLabel(ctrl)
                Call markHoverTitle(ctrl)
            ' Clean Title
            Case isTitle(ctrl)
                Call cleanMarkTitle(ctrl)
            ' Highlight picking + edited + duplicated
            Case ( _
                isPickingLine(lineIndex) _
                And isEditedLine(lineIndex) _
                And isDuplicatedLine(lineIndex) _
            )
                Call markPickingEditedDuplicatedLine(ctrl)
            ' Highlight picking + duplicated
            Case isPickingLine(lineIndex) And isDuplicatedLine(lineIndex)
                Call markPickingDuplicatedLine(ctrl)
            ' Highlight edited + duplicated
            Case isEditedLine(lineIndex) And isDuplicatedLine(lineIndex)
                Call markEditedDuplicatedLine(ctrl)
            ' Highlight picking + edited
            Case isPickingLine(lineIndex) And isEditedLine(lineIndex)
                Call markPickingEditedLine(ctrl)
            ' Highlight hover + edited
            Case isHoverLine(lineIndex) And isEditedLine(lineIndex)
                Call markHoverEditedLine(ctrl)
            ' Highlight duplicated
            Case isDuplicatedLine(lineIndex)
                Call markDuplicatedLine(ctrl)
            ' Highlight picking
            Case isPickingLine(lineIndex)
                Call markPickingLine(ctrl)
            ' Highlight edited
            Case isEditedLine(lineIndex)
                Call markEditedLine(ctrl)
            ' Highlight hover
            Case isHoverLine(lineIndex)
                Call markHoverLine(ctrl)
            ' Default: Clean mark line
            Case Else
                Call cleanMarkLine(ctrl)
        End Select
NextCtrl:
    Next ctrl
End Sub

Private Sub displayTextBox(Optional ByRef isDisplay As Boolean = True)
    Let getEditingTextBox().visible = isDisplay
    Let getEditingLabel().visible = Not isDisplay
End Sub

' CLEANING

Private Sub hidePattern()
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

Private Sub clearUp()
    ' Clear Objects
    Call setInfo(Nothing)
    Call setEventColl(Nothing)
    Call setHoverLabel(Nothing)
    Call setEditingLabel(Nothing)
    Call setEditingTextBox(Nothing)
    Call setPickingLabel(Nothing)
    Call setShortcutC(Nothing)
    ' Clear Arrays
    Erase editedArr
End Sub
