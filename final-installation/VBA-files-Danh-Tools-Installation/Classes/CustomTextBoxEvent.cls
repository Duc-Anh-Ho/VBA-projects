VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CustomTextBoxEvent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents textBoxE As MSForms.textBox
Attribute textBoxE.VB_VarHelpID = -1

' MUTATORS

Public Property Set setTextBox(ByVal textBox As MSForms.textBox)
    Set textBoxE = textBox
End Property

Public Property Get getTextBox() As MSForms.textBox
    Set getTextBox = textBoxE
End Property

' EVENTS

Private Sub textBoxE_KeyDown(ByVal keycode As MSForms.ReturnInteger, ByVal Shift As Integer)
     Call KeyboardShortcutForm.TextBoxKeyDown(textBoxE, keycode, Shift)
End Sub

Private Sub textBoxE_Change()
    Call KeyboardShortcutForm.TextBoxChange(textBoxE)
End Sub
