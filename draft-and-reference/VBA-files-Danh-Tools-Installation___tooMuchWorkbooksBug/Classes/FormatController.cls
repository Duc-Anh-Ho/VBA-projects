VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FormatController"
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
Private targetObject As Object
'Initialize Variables
Private Function hasVariables() As Boolean
On Error GoTo ErrorHandle
    Set info = New MyInfo
    Set system = New SystemUpdate
    Set fileSystem = system.createFileSystem
    If Not system.hasWorkplace Then
        Let hasVariables = False
        GoTo ExecuteProcedure
    End If
    hasVariables = True
    Set targetObject = Selection
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
    
End Sub

'ASSESSORS
    
'MUTATORS
'METHODS

Private Sub copyFormatRanges()
    targetObject.Copy
End Sub

Private Sub pasteFormatRanges()
    If Application.CutCopyMode = xlCopy Then
    targetObject.PasteSpecial _
        Paste:=xlPasteFormats, _
        Operation:=xlNone, _
        SkipBlanks:=False, _
        Transpose:=False
    Application.CutCopyMode = False
    Else
    MsgBox _
            Prompt:=info.getPrompt & _
                "Can not paste formating" & _
                vbNewLine & _
                "Please copy formating by Ctrl + Shift + C first!!", _
            Buttons:=vbOKOnly + vbExclamation, _
            Title:=info.getAuthor
    End If
End Sub

Private Sub copyFormatDrawingObjects()
    'system.ws.Shapes(targetObject.name).PickUp
End Sub

Private Sub pasteFormatDrawingObjects()
    Dim targetObjectItem As Variant
    For Each targetObjectItem In targetObject
        system.ws.Shapes(targetObjectItem.Name).Apply
    Next targetObjectItem
End Sub

Private Sub copyFormatOtherObject()
    system.ws.Shapes(targetObject.Name).PickUp
End Sub
Private Sub pasteFormatOtherObject()
    system.ws.Shapes(targetObject.Name).Apply
End Sub
'MAIN
Public Sub copyFormat()
On Error GoTo ErrorHandle
    If Not hasVariables Then GoTo ExecuteProcedure
'    Call system.speedOn
    Select Case TypeName(targetObject)
        Case "Range"
            Call copyFormatRanges
        Case "DrawingObjects"
            Call copyFormatDrawingObjects 'Group of objects
        Case Else
            Call copyFormatOtherObject ' picture, shape, chart...
    End Select
GoTo ExecuteProcedure
ErrorHandle:
    Call system.tackleErrors
ExecuteProcedure:
'   Call system.speedOff(, , False) ' TO-DO: research why update link clear CutCopyMode
End Sub

Public Sub pasteFormat()
On Error GoTo ErrorHandle
    If Not hasVariables Then GoTo ExecuteProcedure
'    Call system.speedOn
    Select Case TypeName(targetObject)
        Case "Range"
            Call pasteFormatRanges
        Case "DrawingObjects"
            Call pasteFormatDrawingObjects 'Group of object
        Case Else
            Call pasteFormatOtherObject ' picture, shape, chart...
    End Select
GoTo ExecuteProcedure
ErrorHandle:
    Call system.tackleErrors
ExecuteProcedure:
'   Call system.speedOff
End Sub


