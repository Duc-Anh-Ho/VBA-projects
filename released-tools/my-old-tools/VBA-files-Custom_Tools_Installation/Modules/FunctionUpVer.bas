Attribute VB_Name = "FunctionUpVer"
' Source: internet
' Use for Excel version don't have functions below
' 0: CHECKVER (pending)
' 1: TEXTJOIN

Option Explicit

'Public Function CHECKVER() As String
'    Public version_no As Byte
'    version_no = CByte(Application.version)
'    Select Case version
'        Case 8
'            CHECKVER = 2007
'
'End Function

Public Function TEXTJOIN(Delimiter As String, Ignore_Empty As Boolean, ParamArray Text1() As Variant) As String
    
    Dim Cell As Variant, RangeArea As Variant
    Dim x As Long
    
    'Loop Through Each Cell in Given Input
    For Each RangeArea In Text1
        If TypeName(RangeArea) = "Range" Then
            For Each Cell In RangeArea
                If Len(Cell.Value) <> 0 Or Ignore_Empty = False Then
                    TEXTJOIN = TEXTJOIN & Delimiter & Cell.Value
                End If
            Next Cell
        ElseIf TypeName(RangeArea) = "Variant()" Then
            For Each Cell In RangeArea
                If Len(Cell) <> 0 Or Ignore_Empty = False Then
                    TEXTJOIN = TEXTJOIN & Delimiter & Cell
                End If
            Next
        Else
        'Text String was Entered
            If Len(RangeArea) <> 0 Or Ignore_Empty = False Then
                TEXTJOIN = TEXTJOIN & Delimiter & RangeArea
            End If
        End If
    Next RangeArea
    TEXTJOIN = Mid(TEXTJOIN, Len(Delimiter) + 1)
End Function
