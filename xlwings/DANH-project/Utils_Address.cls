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
    Set AddressDictionary = Script.Dictionary
End Sub
' --- Destructor ---
Private Sub Class_Terminate()
    Set AddressDictionary = Nothing
    Set WorksheetEvent = Nothing
    Set Script = Nothing
End Sub
' --- Helper Methods ---
Private Function stringifySheetName(ByRef targetWorksheet As Worksheet) As String
    Dim name As String: Let name = targetWorksheet.name
    Let stringifySheetName = "'" & Replace(name, "'", "''") & "'!"
End Function
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
Friend Sub add( _
    ByRef key As String _
    , ByRef value As String _
    , Optional ByRef isNote As Boolean = True _
)
    Let AddressDictionary(key) = stringifySheetName(WorksheetEvent) & value
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

Friend Sub AddAllNotes()
    Dim key As Object
    For Each key In AddressDictionary.Keys
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
    Dim key As Object: For Each key In AddressDictionary.Keys
        Dim Addr As String: Let Addr = AddressDictionary(key)
        With WorksheetEvent.Range(Addr)
        Call .ClearComments
        End With
    Next key
End Sub


