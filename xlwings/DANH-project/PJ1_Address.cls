VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PJ1_Address"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Implements Interface_Address
' Check README.md for more information
Option Explicit
' --- Dependencies ---
' --- Object ---
Friend Sub Interface_Address_Pair(ByRef Addr As Utils_Address)
    Call Addr.add("EVIDENCE", "E6")
    Call Addr.add("CHECKLIST", "E7")
    Call Addr.add("EVIDENCE_EXISTS_CHECK", "F6")
    Call Addr.add("CHECKLIST_EXISTS_CHECK", "F7")
End Sub

