VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PicturesController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Check README.md for more information
Option Explicit
'Declare Variables
Private info As Object
Private system As Object
Private fileSystem As Object
Private userResponse As VbMsgBoxResult
Private Const DEFAULT_OFFSET As Byte = 10
Private targetPlace As Object
Private picCaptured As Picture
Private offsetTop As Byte
Private offsetLeft As Byte
Private pictureTop As Long
Private pictureLeft As Long
Private pictureHeight As Long
Private pictureWidth As Long
Private isLockRatio As Boolean
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

Public Property Get getOffset(Optional ByVal customOffset = DEFAULT_OFFSET) As Integer
    Let getOffset = customOffset
End Property

'MUTATORS

'METHODS

Private Sub setTargetArea()
    Select Case TypeName(Selection)
        Case "Range"
            Set targetPlace = Selection
        Case "DrawingObjects"
            MsgBox _
                Prompt:="Sorry! we've not supported " & TypeName(Selection) & " yet" _
                    & vbNewLine & "Please select another object type !!!!", _
                Buttons:=vbOKOnly + vbCritical, _
                Title:=info.getAuthor
            Call killProject
        Case Else
            userResponse = MsgBox( _
                Prompt:="You've selected: '" & Selection.Name & _
                    "' as type: " & TypeName(Selection) & vbNewLine & _
                    "Do you want to replace " & Selection.Name & "?", _
                Buttons:=vbYesNoCancel + vbQuestion, _
                Title:=info.getAuthor)
            'yes pressed
            If userResponse = vbYes Then
                Set targetPlace = Selection
            'no pressed
            ElseIf userResponse = vbNo Then
                Set targetPlace = Selection
            'X or cancel Pressed
            Else
                Call killProject
            End If
    End Select
End Sub

Private Sub setPicturePosition()
     Select Case TypeName(targetPlace)
        Case "Range"
            Let offsetTop = DEFAULT_OFFSET
            Let offsetLeft = DEFAULT_OFFSET
        Case Else
            Let offsetTop = 0
            Let offsetLeft = 0
    End Select
    Let pictureTop = targetPlace.Top + offsetTop
    Let pictureLeft = targetPlace.Left + offsetLeft
    Let pictureHeight = targetPlace.Height
    Let pictureWidth = targetPlace.Width
End Sub

Private Sub deleteOldPicture()
    If userResponse = vbYes And TypeName(targetPlace) <> "Range" Then
        targetPlace.delete
    End If
End Sub

Private Sub capturePuture()
    With system.app
        .ScreenUpdating = False
        .CommandBars.ExecuteMso "ScreenClipping" 'important key
    End With
    If TypeName(Selection) = "Picture" Then
        Set picCaptured = Selection
'        picCaptured.Cut
'    'Paste to A1 as default
'    picCaptured = ws.Paste Destination:=ws.Range("A1")
    Else
        MsgBox _
            Prompt:=info.getPrompt & "Failed to capture!!!" _
                & vbNewLine & "Please try again", _
            Buttons:=vbOKOnly + vbExclamation, _
            Title:=info.getAuthor
        Call capturePuture 'recursion
    End If
End Sub

Private Sub scalePicture(ByVal isLockRatio As Boolean)
    'Scale up when too small
    If pictureHeight < DEFAULT_OFFSET * 2 Or pictureWidth < DEFAULT_OFFSET * 2 Then
        pictureHeight = pictureHeight + DEFAULT_OFFSET * 2
        pictureWidth = pictureWidth + DEFAULT_OFFSET * 2
    End If
    'Set top and left
    With picCaptured
        .ShapeRange.LockAspectRatio = isLockRatio
        .Left = pictureLeft
        .Top = pictureTop
    End With
    'Check Lock Ratio and set height and width
    If isLockRatio Then
        If pictureWidth > pictureHeight Then
            picCaptured.Height = pictureHeight - offsetTop * 2
        Else
            picCaptured.Width = pictureWidth - offsetLeft * 2
        End If
    End If
    If Not isLockRatio Then
        With picCaptured
            .Width = pictureWidth - offsetLeft * 2
            .Height = pictureHeight - offsetTop * 2
        End With
    End If
    'Reset default settings
    picCaptured.ShapeRange.LockAspectRatio = True
    system.app.ScreenUpdating = True
End Sub

Private Sub killProject()
    system.app.ScreenUpdating = True
    End
End Sub

Public Sub setLockRadio()
     Let isLockRatio = SnippingToolForm.CheckBoxLockRadio.Value
End Sub
'2-0  Move or  Size With Cell
Private Sub MoveAndSize(Optional ByVal placementStt As Byte = xlMoveAndSize)
    Dim sh As Object
    For Each sh In ActiveSheet.Shapes
        sh.Placement = placementStt
    Next sh
End Sub

'MAIN
'1. Snipping Tool
Public Sub snip()
On Error GoTo ErrorHandle
    If Not hasVariables Then GoTo ExecuteProcedure
    Call system.speedOn
    Call setTargetArea
    Call setPicturePosition
    Call deleteOldPicture
    Call capturePuture
    Call setLockRadio
    Call scalePicture(isLockRatio) 'TODO: Optional choose
GoTo ExecuteProcedure
ErrorHandle:
    Call system.tackleErrors
ExecuteProcedure:
    Call system.speedOff
End Sub
'2. Move Or Size with Cell all shape
Public Sub setPlacement(Optional ByVal placementStt As Byte = xlMoveAndSize)
On Error GoTo ErrorHandle
    If Not hasVariables Then GoTo ExecuteProcedure
    Call system.speedOn
    Call MoveAndSize(placementStt)
GoTo ExecuteProcedure
ErrorHandle:
    Call system.tackleErrors
ExecuteProcedure:
    Call system.speedOff
End Sub
