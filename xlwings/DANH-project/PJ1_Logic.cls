VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PJ1_Logic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Check README.md for more information
Option Explicit
' --- Dependencies ---
Private Utils As Utils_Test
Private Addr As Utils_Address
' Private IAs InfoConstants
' Private System As SystemUpdate
' --- Object ---
Private userResponse As VbMsgBoxResult
' --- Constant ---
' --- Constructor ---
Friend Sub Init( _
    Optional ByRef targetWorksheet As Worksheet _
)
    If targetWorksheet Is Nothing Then
        Set targetWorksheet = Excel.Application.ActiveSheet
    End If
    Set Utils = New Utils_Test
    Set Addr = New Utils_Address
    Call Addr.Init(targetWorksheet, New PJ1_Address)
    Call Me.CheckFileName
End Sub
' --- Destructor ---
Private Sub Class_Terminate()
    Set Utils = Nothing
    Set Addr = Nothing
End Sub
' --- Fuctions / Subroutines ---
Friend Sub CheckFileName()
    Dim testCasePath As String: Let testCasePath = Addr.GetValue("EVIDENCE")
    Dim checkListPath As String: Let checkListPath = Addr.GetValue("CHECKLIST")
    If Utils.IsExistPath(testCasePath) Then
        Call Addr.SetValue("EVIDENCE_EXISTS_CHECK", Utils.PASS)
    Else
        Call Addr.SetValue("EVIDENCE_EXISTS_CHECK", Utils.FAIL)
    End If
    If Utils.IsExistPath(checkListPath) Then
        Call Addr.SetValue("CHECKLIST_EXISTS_CHECK", Utils.PASS)
    Else
        Call Addr.SetValue("CHECKLIST_EXISTS_CHECK", Utils.FAIL)
    End If
End Sub



