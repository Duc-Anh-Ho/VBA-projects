VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MultipleReplaceForm 
   Caption         =   "Multiple Replace"
   ClientHeight    =   2955
   ClientLeft      =   119
   ClientTop       =   462
   ClientWidth     =   7336
   OleObjectBlob   =   "MultipleReplaceForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MultipleReplaceForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Check README.md for more information

Option Explicit
'Declare Variables
Private userResponse As VbMsgBoxResult
Private info As InfoConstants
Private withinIndex As Byte
Private selectedArea As Range
Private findArea As Range
Private replaceArea As Range

'EVENTS

Private Sub UserForm_Initialize()
    Set info = New InfoConstants
    'FindAreaInput.SetFocus
    SelectedAreaInput.Visible = False
    MatchCaseCheckBox.value = False
    MatchByteCheckBox.value = False
    MatchContentCheckBox.value = False
    With WithInComboBox
        .Style = fmStyleDropDownList
        .AddItem "Selection" ' 0 Or False
        .AddItem "Sheet" ' 1
        .AddItem "Workbook" ' 2
        .ListIndex = "1"
    End With
    With SearchComboBox
        .Style = fmStyleDropDownList
        .AddItem "By Rows"
        .AddItem "By Columns"
        .ListIndex = "0"
    End With
End Sub

Private Sub SearchLabel_Click()
    SearchComboBox.DropDown
End Sub

Private Sub FindWhatLabel_Click()
    FindAreaInput.SetFocus
End Sub

Private Sub ReplaceWithLabel_Click()
    ReplaceAreaInput.SetFocus
End Sub

Private Sub WithinLabel_Click()
    WithInComboBox.DropDown
End Sub

Private Sub FindUnderlineLabel_Click()
    Call FindWhatLabel_Click
End Sub

Private Sub ReplaceWithUnderlineLabel_Click()
    Call ReplaceWithLabel_Click
End Sub

Private Sub WithinUnderline_Click()
    Call WithinLabel_Click
End Sub

Private Sub SearchUnderlineLabel_Click()
    Call SearchLabel_Click
End Sub

Private Sub MatchCaseUnderlineLabel_Click()
    MatchCaseCheckBox.value = Not MatchCaseCheckBox.value
End Sub

Private Sub MatchByteUnderLineLabel_Click()
    MatchByteCheckBox.value = Not MatchByteCheckBox.value
End Sub

Private Sub MatchContentUnderLIneLabel_click()
    MatchContentCheckBox.value = Not MatchContentCheckBox.value
End Sub

Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call shortcut(KeyCode, Shift)
End Sub

Private Sub MultiPage_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call shortcut(KeyCode, Shift)
End Sub

Private Sub ReplaceAllUnderLineLabel_Click()
    Call ReplaceAllButton_Click
End Sub

Private Sub CloseUnderLineLabel_Click()
    Call CloseButton_Click
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Call closeForm
End Sub

Private Sub UserForm_Terminate()
    Call closeForm
End Sub

Private Sub CloseButton_Click()
    Call closeForm
End Sub

Private Sub WithInComboBox_Change()
    SelectedAreaInput.text = vbNullString
    Let withinIndex = WithInComboBox.ListIndex
    ' Within Selection
    If Not CBool(withinIndex) Then
        SelectedAreaInput.Visible = True
        SelectedAreaInput.SetFocus
    Else
        SelectedAreaInput.Visible = False
        ReplaceAllButton.SetFocus
    End If
End Sub

Private Sub FindAreaInput_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Trim(FindAreaInput.text & vbNullString) = vbNullString Then Exit Sub
    On Error Resume Next
    Set findArea = Application.Evaluate(FindAreaInput.text)
    On Error GoTo 0
    If typeName(findArea) <> "Range" Then
        Let userResponse = MsgBox( _
            Prompt:=info.getPrompt & "Please choose 'Find What:' as Range!", _
            Buttons:=vbOKOnly + vbExclamation, _
            Title:=info.getAuthor)
        Let Cancel = True ' Keep focus
        ' Select all text
        Let FindAreaInput.SelStart = 0
        Let FindAreaInput.SelLength = Len(FindAreaInput.text)
    End If
End Sub

Private Sub ReplaceAreaInput_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Trim(ReplaceAreaInput.text & vbNullString) = vbNullString Then Exit Sub
    On Error Resume Next
    Set replaceArea = Application.Evaluate(ReplaceAreaInput.text)
    On Error GoTo 0
    If typeName(replaceArea) <> "Range" Then
        Let userResponse = MsgBox( _
            Prompt:=info.getPrompt & "Please choose 'Replace With:' as Range!", _
            Buttons:=vbOKOnly + vbExclamation, _
            Title:=info.getAuthor)
         Let Cancel = True ' Keep focus
         ' Select all text
         Let ReplaceAreaInput.SelStart = 0
         Let ReplaceAreaInput.SelLength = Len(ReplaceAreaInput.text)
    End If
End Sub

Private Sub SelectedAreaInput_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Trim(SelectedAreaInput.text & vbNullString) = vbNullString Then Exit Sub
    On Error Resume Next
    Set selectedArea = Application.Evaluate(SelectedAreaInput.text)
    On Error GoTo 0
    If typeName(selectedArea) <> "Range" Then
        Let userResponse = MsgBox( _
            Prompt:=info.getPrompt & "Please choose a selected area as Range!", _
            Buttons:=vbOKOnly + vbExclamation, _
            Title:=info.getAuthor)
         Let Cancel = True ' Keep focus
         Let SelectedAreaInput.text = vbNullString 'Clear text
    End If
End Sub

Private Sub ReplaceAllButton_Click()
    Dim rangeC As New RangesController
    If Trim(FindAreaInput.text & vbNullString) = vbNullString Then
        Let userResponse = MsgBox( _
            Prompt:=info.getPrompt & "Please choose 'Find What'!", _
            Buttons:=vbOKOnly + vbExclamation, _
            Title:=info.getAuthor)
            FindAreaInput.SetFocus
    ElseIf Trim(ReplaceAreaInput.text & vbNullString) = vbNullString Then
        Let userResponse = MsgBox( _
            Prompt:=info.getPrompt & "Please choose 'Replace With'!", _
            Buttons:=vbOKOnly + vbExclamation, _
            Title:=info.getAuthor)
            ReplaceAreaInput.SetFocus
    ' Within Selection
    ElseIf Not CBool(withinIndex) _
        And Trim(SelectedAreaInput.text & vbNullString) = vbNullString Then
        Let userResponse = MsgBox( _
            Prompt:=info.getPrompt _
                & vbNewLine & "Within Selection can not be empty." _
                & vbNewLine & "Please choose a selected area!", _
            Buttons:=vbOKOnly + vbExclamation, _
            Title:=info.getAuthor)
            SelectedAreaInput.SetFocus
    ElseIf findArea.Count <> replaceArea.Count Then
        Let userResponse = MsgBox( _
            Prompt:=info.getPrompt _
                & vbNewLine & "Number of 'Find What' and 'Replace With' does not match." _
                & vbNewLine & "Please check again!", _
            Buttons:=vbOKOnly + vbExclamation, _
            Title:=info.getAuthor)
            ReplaceAreaInput.SetFocus
    Else
        Call rangeC.multipleReplace( _
            findArea:=findArea _
            , replaceArea:=replaceArea _
            , withinIndex:=withinIndex _
            , selectedArea:=selectedArea _
            , isMatchCase:=MatchCaseCheckBox.value _
            , isMatchByte:=MatchByteCheckBox.value _
            , isMatchContent:=MatchContentCheckBox.value _
            , searchOrderCd:=SearchComboBox.value _
        )
    End If
End Sub

' FUNCTIONS

Private Sub shortcut( _
    ByVal KeyAscii As Integer, _
    ByVal Shift As Byte _
)
    Dim altCode As Byte: Let altCode = 4 ' fmAltMask
    Select Case True
        Case KeyAscii = vbKeyN And Shift = altCode ' Alt + N
            FindWhatLabel_Click
        Case KeyAscii = vbKeyE And Shift = altCode ' Alt + E
            ReplaceWithLabel_Click
        Case KeyAscii = vbKeyH And Shift = altCode ' Alt + H
            WithinLabel_Click
        Case KeyAscii = vbKeyA And Shift = altCode ' Alt + A
            ReplaceAllButton.SetFocus
        Case KeyAscii = vbKeyC And Shift = altCode ' Alt + C
            CloseButton.SetFocus
        Case KeyAscii = vbKeyEscape ' Ecs
            Call closeForm
    End Select
End Sub

Public Sub closeForm()
    FindAreaInput.text = ""
    ReplaceAreaInput.text = ""
    FindAreaInput.Enabled = False
    ReplaceAreaInput.Enabled = False
    Unload Me
    End
End Sub


