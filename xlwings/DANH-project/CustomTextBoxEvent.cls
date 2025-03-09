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

Private WithEvents textBoxE As MsForms.textBox
' Attribute textBoxE.VB_VarHelpID = -1

' MUTATORS

Public Property Set setTextBox(ByRef textBox As MsForms.textBox): Set textBoxE = textBox: End Property

' ACCESSORS

Public Property Get getTextBox() As MsForms.textBox: Set getTextBox = textBoxE: End Property

' Constructor
Private Sub Class_Initialize()
    '
End Sub

' Destructor
Private Sub Class_Terminate()
    Set textBoxE = Nothing
End Sub

' EVENTS

Private Sub textBoxE_KeyDown(ByVal KeyCode As MsForms.ReturnInteger, ByVal Shift As Integer)
    Call KeyboardShortcutForm.textBoxKeyDown(textBoxE, KeyCode, Shift)
End Sub

Private Sub textBoxE_Change()
    Call KeyboardShortcutForm.textBoxChange(textBoxE)
End Sub
