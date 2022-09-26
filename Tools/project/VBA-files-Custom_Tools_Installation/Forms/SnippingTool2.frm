VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SnippingTool2 
   Caption         =   "Snipping tool (Ctrl + q)"
   ClientHeight    =   2415
   ClientLeft      =   119
   ClientTop       =   462
   ClientWidth     =   3689
   OleObjectBlob   =   "SnippingTool2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SnippingTool2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Button_finish_Click()
    ThisWorkbook.Close SaveChanges:=False
End Sub

Private Sub Image_Scisors_Click()
    Call Snipping
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ThisWorkbook.Close SaveChanges:=False
End Sub
