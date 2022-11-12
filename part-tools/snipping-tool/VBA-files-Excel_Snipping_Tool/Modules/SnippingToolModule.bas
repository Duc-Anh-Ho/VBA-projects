Attribute VB_Name = "SnippingToolModule"
' AUTO ADD ADDIN TOOLS
' Author: DANH
' Version: 2.0.0
' Update: 2022/10/09
' Check README.md for more information

Option Explicit

Private Const VERSION = "v1.1.0"
Private Const AUTHOR_PROMPT As String = "AUTHOR: "
Private Const AUTHOR_NAME As String = "DANH"
Private Const DEFAULT_OFFSET As Byte = 10
Private fileSystem As Object
Private userResponse As VbMsgBoxResult
Private app As Application
Private wb As Workbook
Private ws As Worksheet
Private targetPlace As Object
Private picCaptured As Picture
Private offsetTop As Byte
Private offsetLeft As Byte
Private pictureTop As Long
Private pictureLeft As Long
Private pictureHeight As Long
Private pictureWidth As Long
Private isLockRatio As Boolean

'Functions

Private Static Sub initializeVariables()
    'ThisWorkbook.Author = AUTHOR_NAME
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    Set app = Application
    If app.ActiveWorkbook Is Nothing Or app.ActiveSheet Is Nothing Then
        MsgBox _
            Prompt:=AUTHOR_PROMPT & _
                "You have to select target position to place picture first", _
            Buttons:=vbOKOnly + vbExclamation, _
            Title:=AUTHOR_NAME
        Call killProject
    Else
        Set wb = app.ActiveWorkbook
        Set ws = wb.ActiveSheet
    End If
End Sub

Private Static Sub setTargetArea()
    Select Case TypeName(Selection)
        Case "Range"
            Set targetPlace = Selection
        Case "DrawingObjects"
            MsgBox _
                Prompt:="Sorry! we've not supported " & TypeName(Selection) & " yet" _
                    & vbNewLine & "Please select another object type !!!!", _
                Buttons:=vbOKOnly + vbCritical, _
                Title:=AUTHOR_NAME
            Call killProject
        Case Else
            userResponse = MsgBox( _
                Prompt:="You've selected: '" & Selection.Name & _
                    "' as type: " & TypeName(Selection) & vbNewLine & _
                    "Do you want to replace " & Selection.Name & "?", _
                Buttons:=vbYesNoCancel + vbQuestion, _
                Title:=AUTHOR_NAME)
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

Private Static Sub setPicturePosition()
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

Private Static Sub deleteOldPicture()
    If userResponse = vbYes And TypeName(targetPlace) <> "Range" Then
        targetPlace.Delete
    End If
End Sub

Private Static Sub capturePuture()
    With app
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
            Prompt:=AUTHOR_PROMPT & "Failed to capture!!!" _
                & vbNewLine & "Please try again", _
            Buttons:=vbOKOnly + vbExclamation, _
            Title:=AUTHOR_NAME
        Call capturePuture 'recursion
    End If
End Sub

Private Static Sub scalePicture(ByVal isLockRatio As Boolean)
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
    app.ScreenUpdating = True
End Sub

Private Static Sub killProject()
    app.ScreenUpdating = True
    End
End Sub

'Options

Public Static Sub setLockRadio()
     Let isLockRatio = SnippingToolForm.CheckBoxLockRadio.Value
End Sub

Public Static Function getOffset(Optional ByVal customOffset = DEFAULT_OFFSET) As Integer
    Let getOffset = customOffset
End Function

'MAIN
Public Static Sub snip()
On Error GoTo ErrorHandle
    Call initializeVariables
    Call setTargetArea
    Call setPicturePosition
    Call deleteOldPicture
    Call capturePuture
    Call setLockRadio
    Call scalePicture(isLockRatio) 'TODO: Optional choose
GoTo ExecuteProcedure
ErrorHandle:
    Call tackleErrors
ExecuteProcedure:
End Sub

'Errors
Private Static Sub tackleErrors()
    Select Case Err.Number
        Case 0
        'Can't enable Addin
        Case 1004
        'VBA file have password
        Case 50289
            MsgBox _
                Prompt:=Err.Description & _
                    vbNewLine & _
                    AUTHOR_PROMPT & _
                    "Can not access this VBA file because of been protected by a password!", _
                Buttons:=vbCritical + vbOKOnly, _
                Title:=AUTHOR_NAME
        'Un-handled Error
        Case Else
            Call errorDisplay
    End Select
    Application.ScreenUpdating = True
    On Error GoTo 0
End Sub

Private Static Sub errorDisplay()
    Dim errorMessage As String: Let errorMessage = _
        "Error # " & Str(Err.Number) & _
        " was generated by " & Err.Source & _
        Chr(13) & "Error Line: " & Erl & _
        Chr(13) & Err.Description
    MsgBox _
        Prompt:=errorMessage, _
        Title:="Error", _
        HelpFile:=Err.HelpFile, _
        Context:=Err.HelpContext
End Sub

' PerCls tool rework
Private Static Sub speedOn()
    With Application
        .ScreenUpdating = False
        .AskToUpdateLinks = False
        .DisplayAlerts = False
    End With
End Sub

Private Static Sub speedOff()
    With Application
        .ScreenUpdating = True
        .AskToUpdateLinks = True
        .DisplayAlerts = True
    End With
End Sub
