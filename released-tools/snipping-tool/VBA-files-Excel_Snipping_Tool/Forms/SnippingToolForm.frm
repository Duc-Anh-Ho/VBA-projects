VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SnippingToolForm 
   Caption         =   "Snipping tool"
   ClientHeight    =   2884
   ClientLeft      =   119
   ClientTop       =   462
   ClientWidth     =   3661
   OleObjectBlob   =   "SnippingToolForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SnippingToolForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CheckBoxLockRadio_Change()
    Call SnippingToolModule.setLockRadio
End Sub

Private Sub CloseButton_Click()
    Unload SnippingToolForm
    ThisWorkbook.Close savechanges:=False
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Unload SnippingToolForm
        Call CloseButton_Click
        Cancel = True
        Application.ScreenUpdating = True
        ThisWorkbook.Close savechanges:=False
        End '<<TODO: Tim hieu tai sao
    End If
End Sub

Private Sub Scisors_Icon_Click()
    Call SnippingToolModule.snip
 '   Unload SnippingToolForm
End Sub

'Color

Private Sub Scisors_Icon_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Scisors_Icon.SpecialEffect = fmSpecialEffectBump
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    CloseButton.BackColor = vbButtonFace
    Scisors_Icon.SpecialEffect = fmSpecialEffectFlat
End Sub
