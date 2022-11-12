VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RangesController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' AUTO INVERT COLOR
Option Explicit
'Declare Variables
Private info As Object
Private system As Object
Private fileSystem As Object
Private targetRange As Range ' Object
Private userResponse As VbMsgBoxResult
Private Const INPUT_BOX_RANGE_TYPE As Byte = 8
Private Const START_CHAR As Byte = 1
'Initialize Variables
Private Function hasVariables() As Boolean
On Error GoTo ErrorHandle
    Set info = New MyInfo
    Set system = New SystemUpdate
    Set fileSystem = system.createFileSystem
    If Not system.hasWorkplace Then
        hasVariables = False
        GoTo ExecuteProcedure
    End If
    If TypeName(Selection) = "Range" Then
'    If TypeName(Selection) = "ChartArea" Then
        Set targetRange = Selection
    Else
        MsgBox _
            Prompt:=info.getPrompt & _
                "You only can select ranges!", _
            Buttons:=vbOKOnly + vbExclamation, _
            Title:=info.getAuthor
        hasVariables = False
        GoTo ExecuteProcedure
    End If
    hasVariables = True
    GoTo ExecuteProcedure
ErrorHandle:
    Call system.tackleErrors
ExecuteProcedure:
End Function
'Constructor
Private Sub Class_Initialize()
'
End Sub
'Destructor
Private Sub Class_Terminate()
'
End Sub

'ASSESSORS
'1-0 Inver Color
Private Property Get getInvertColor(ByVal RGBColor As Long) As Long
    Dim redColor As Integer
    Dim greenColor As Integer
    Dim blueColor As Integer
     'Algorithm color
    Let redColor = 255 - RGBColor Mod 256
    Let greenColor = 255 - RGBColor \ 256 Mod 256
    Let blueColor = 255 - RGBColor \ 65536 Mod 256
    Let getInvertColor = RGB(redColor, greenColor, blueColor)
End Property
'2-0 Bold First Line
Private Property Get getNewLine() As String
    If system.app.OperatingSystem Like "Windows*" Then
        getNewLine = Chr(10)
    Else
        getNewLine = vbNewLine
    End If
End Property

'MUTATORS

'METHODS
'2-1 Trigger Bold First Line
Private Sub triggerBoldFirstLine(ByRef cell As Range)
    cell.Characters( _
            Start:=START_CHAR, _
            Length:=InStr( _
                cell.Value, _
                getNewLine)) _
        .Font.Bold = Not cell.Characters( _
            Start:=START_CHAR, _
            Length:=InStr( _
                cell.Value, _
                getNewLine)) _
        .Font.Bold
End Sub
'MAIN
'1. Inver Color
Public Sub invertColor()
On Error GoTo ErrorHandle
    If Not hasVariables Then GoTo ExecuteProcedure
    Call system.speedOn
    Dim cell As Object
    'Loop through range areas
    For Each cell In targetRange
        Let cell.Interior.Color = getInvertColor(cell.Interior.Color)
        Let cell.Font.Color = getInvertColor(cell.Font.Color)
        Let cell.Borders.Color = getInvertColor(cell.Borders.Color)
    Next cell
GoTo ExecuteProcedure
ErrorHandle:
    Call system.tackleErrors
ExecuteProcedure:
    Call system.speedOff
End Sub
'2. Bold First Line
Public Sub BoldFirstLine()
On Error GoTo ErrorHandle
    If Not hasVariables Then GoTo ExecuteProcedure
    Call system.speedOn
    Dim cell As Range
    If targetRange Is Nothing Then GoTo ExecuteProcedure
    'Loop through ranges
    For Each cell In targetRange
        'Trigger Bold First Line
        Call triggerBoldFirstLine(cell)
    Next cell
    GoTo ExecuteProcedure
ErrorHandle:
    Call system.tackleErrors
ExecuteProcedure:
    Call system.speedOff
End Sub


