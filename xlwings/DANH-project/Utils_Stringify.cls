VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Utils_Stringify"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Check README.md for more information
Option Explicit
' Constants
Private Const OPEN_BRACER As String = "{{"
Private Const CLOSE_BRACER As String = "}}"
' --- Fuctions / Subroutines ---
Friend Function SheetName(ByRef targetWorksheet As Worksheet) As String
    Dim name As String: Let name = targetWorksheet.name
    Let SheetName = "'" & VBA.Strings.Replace(name, "'", "''") & "'!"
End Function

Friend Function ReplacePlaceholder(ByRef templateText As String, ByRef dict As Object) As String
    Dim resultText As String
    Let resultText = templateText
    Dim key As Variant: For Each key In dict.Keys
        Let resultText = VBA.Strings.Replace( _
            resultText _
            , OPEN_BRACER & key & CLOSE_BRACER _
            , VBA.Conversion.CStr(dict(key)) _
        )
    Next key
    Let ReplacePlaceholder = resultText
End Function

