VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CustomLabelEvent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents labelE As MSForms.label
Attribute labelE.VB_VarHelpID = -1

' MUTATORS

Public Property Set setLabel(ByVal label As MSForms.label)
    Set labelE = label
End Property

Public Property Get getLabel() As MSForms.label
    Set getLabel = labelE
End Property

' EVENTS

Private Sub LabelE_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Call KeyboardShortcutForm.labelMoveOn(labelE)
End Sub

Private Sub LabelE_Click()
    Call KeyboardShortcutForm.labelClick(labelE)
End Sub


Private Sub labelE_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Call KeyboardShortcutForm.labelDbClick(labelE)
End Sub

