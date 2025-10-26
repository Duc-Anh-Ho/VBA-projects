VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Utils_Address"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Check README.md for more information
Option Explicit
' --- Dependencies ---
Private Script As Utils_Scripts
Private Stringify As Utils_Stringify
' --- Object ---
Private AddressDictionary As Object
Private WithEvents WorksheetEvent As Worksheet
Attribute WorksheetEvent.VB_VarHelpID = -1
' --- Constructor ---
Friend Sub Init( _
    ByRef targetWorksheet As Worksheet _
    , ByRef addressSource As Interface_Address _
)
    Set Me.UseSheet = targetWorksheet
    Call addressSource.Pair(Me)
End Sub

Private Sub Class_Initialize()
    Set Script = New Utils_Scripts
    Set Stringify = New Utils_Stringify
    Set AddressDictionary = Script.Dictionary
End Sub
' --- Destructor ---
Private Sub Class_Terminate()
    Set Script = Nothing
    Set Stringify = Nothing
    Set AddressDictionary = Nothing
    Set WorksheetEvent = Nothing
End Sub
' --- Property Accessor / Mutator ---
Friend Property Set UseSheet(ByRef targetWorksheet As Worksheet)
    Set WorksheetEvent = targetWorksheet
End Property

Friend Property Get UseSheet() As Worksheet
    Set UseSheet = WorksheetEvent
End Property

Friend Property Set Dictionary(ByRef dict As Object)
    Set AddressDictionary = dict
End Property

Friend Property Get Dictionary() As Object
    Set Dictionary = AddressDictionary
End Property
' --- Fuctions / Subroutines ---
Friend Sub Add( _
    ByRef key As String _
    , ByRef value As String _
    , Optional ByRef isNote As Boolean = True _
)
    Let AddressDictionary(key) = Stringify.sheetName(WorksheetEvent) & value
    With WorksheetEvent
    ' Add Note
    Call .Range(AddressDictionary(key)).ClearComments
    If isNote Then
        Call .Range(AddressDictionary(key)).NoteText( _
            text:=key _
            , START:=1 _
            , Length:=False _
        )
    End If
    End With
End Sub

Friend Function GetByKey(ByRef key As String) As String
    Let GetByKey = AddressDictionary(key)
End Function

Friend Sub SetValue(ByRef key As String, ByRef newValue As String)
    Let WorksheetEvent.Range(AddressDictionary(key)).value = newValue
End Sub

Friend Function GetValue(ByRef key As String) As String
    Let GetValue = WorksheetEvent.Range(AddressDictionary(key)).value
End Function

Friend Function GetValueArray(ByRef key As String) As String()
    Dim Array2D As Variant
    Dim Array1D() As String
    Let Array2D = WorksheetEvent.Range(AddressDictionary(key)).value
    ReDim Array1D(1 To UBound(Array2D, 1))
    Dim i As Long: For i = 1 To UBound(Array2D, 1)
        Let Array1D(i) = CStr(Array2D(i, 1))
    Next i
    Let GetValueArray = Array1D
End Function

' TODO: Change input to Variant for 1DArr Or 2DArr Or Collection
Friend Sub SetValueArray(ByRef key As String, ByRef inArray1D() As String)
    Dim rng As Range
    Dim Array2D As Variant
    Set rng = WorksheetEvent.Range(AddressDictionary(key))
    ReDim Array2D(LBound(inArray1D) To UBound(inArray1D), 1 To 1)
    If UBound(inArray1D) - LBound(inArray1D) + 1 <> rng.Rows.count Then
        MsgBox "Error: Array size (" & (UBound(inArray1D) - LBound(inArray1D) + 1) & _
               ") does not match target range rows (" & rng.Rows.count & ")", vbCritical
        Stop
        Exit Sub
    End If
    Dim i As Long: For i = LBound(inArray1D) To UBound(inArray1D)
        Let Array2D(i, 1) = inArray1D(i)
    Next i
    Let rng.value = Array2D
End Sub

Friend Sub AddAllNotes()
    Dim key As Variant: For Each key In AddressDictionary.Keys
        Dim Addr As String: Let Addr = AddressDictionary(key)
        With WorksheetEvent.Range(Addr)
        Call .ClearComments
        Call .NoteText( _
            text:=key _
            , START:=1 _
            , Length:=False _
        )
        End With
    Next key
End Sub

Friend Sub ClearAllNotes()
    Dim key As Variant: For Each key In AddressDictionary.Keys
        Dim Addr As String: Let Addr = AddressDictionary(key)
        Call WorksheetEvent.Range(Addr).ClearComments
    Next key
End Sub




