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
Private info As InfoConstants
Private system As SystemUpdate
Private fileSystem As Object
Private userResponse As VbMsgBoxResult
Private targetObject As Object
'Initialize Variables
Private Function hasVariables() As Boolean
On Error GoTo ErrorHandle
    Set info = New InfoConstants
    Set system = New SystemUpdate
    Set fileSystem = system.createFileSystem
    If Not system.hasWorkPlace(hasMsg:=True, workPlaceType:="xlWorksheet") Then
        Let hasVariables = False
        GoTo ExecuteProcedure
    End If
    Let hasVariables = True
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
    '
End Sub

'ASSESSORS
    
'MUTATORS
'METHODS

Private Sub copyFormatRanges()
    targetObject.copy
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

Private Sub pasteValueRanges()
    If Application.CutCopyMode = xlCopy Then
        targetObject.PasteSpecial _
            Paste:=xlPasteValues, _
            Operation:=xlNone, _
            SkipBlanks:=False, _
            Transpose:=False
        Application.CutCopyMode = False
    Else
        MsgBox _
                Prompt:=info.getPrompt & _
                    "Can not paste value" & _
                    vbNewLine & _
                    "Please copy value by Ctrl + C first!!", _
                Buttons:=vbOKOnly + vbExclamation, _
                Title:=info.getAuthor
    End If
End Sub

'Used default function don't need to override
Private Sub copyFormatDrawingObjects()
    'system.ws.Shapes(targetObject.name).PickUp
End Sub

Private Sub pasteFormatDrawingObjects()
    Dim targetObjectItem As Variant
    For Each targetObjectItem In targetObject
        system.ws.Shapes(targetObjectItem.name).Apply
    Next targetObjectItem
End Sub

Private Sub copyFormatOtherObject()
    system.ws.Shapes(targetObject.name).PickUp
End Sub
Private Sub pasteFormatOtherObject()
    system.ws.Shapes(targetObject.name).Apply
End Sub
'2-0  Move or  Size With Cell
Private Sub MoveAndSize(Optional ByVal placementStt As Byte = xlMoveAndSize)
    Dim Sh As Object
    For Each Sh In system.ws.Shapes
        Sh.Placement = placementStt
    Next Sh
End Sub
'3-0 Clear content (default Del)
Public Sub clearContentRanges()
    targetObject.ClearContents
End Sub
'3-1 Clear format
Public Sub clearFormatRanges()
    targetObject.ClearFormats
End Sub
'3-2 Clear all
Public Sub clearFormatContentRange()
    targetObject.Clear
End Sub

'MAIN
'1.1 For override shortcut
Public Sub copyFormat()
On Error GoTo ErrorHandle
    If Not hasVariables Then GoTo ExecuteProcedure
'    Call system.speedOn
    Select Case typeName(targetObject)
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
'1.2 For override shortcut
Public Sub pasteFormat()
On Error GoTo ErrorHandle
    If Not hasVariables Then GoTo ExecuteProcedure
'    Call system.speedOn
    Select Case typeName(targetObject)
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
'1.3 For override shortcut
Public Sub pasteValue()
On Error GoTo ErrorHandle
    If Not hasVariables Then GoTo ExecuteProcedure
'    Call system.speedOn
    Select Case typeName(targetObject)
        Case "Range"
            Call pasteValueRanges
    End Select
GoTo ExecuteProcedure
ErrorHandle:
    Call system.tackleErrors
ExecuteProcedure:
'   Call system.speedOff
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
'3.0 Clear format
Public Sub clearContent()
On Error GoTo ErrorHandle
    If Not hasVariables Then GoTo ExecuteProcedure
    Call system.speedOn
    Select Case typeName(targetObject)
        Case "Range"
            Call clearContentRanges
    End Select
GoTo ExecuteProcedure
ErrorHandle:
    Call system.tackleErrors
ExecuteProcedure:
    Call system.speedOff
End Sub
'3.1 Clear format
Public Sub clearFormat()
On Error GoTo ErrorHandle
    If Not hasVariables Then GoTo ExecuteProcedure
    Call system.speedOn
    Select Case typeName(targetObject)
        Case "Range"
            Call clearFormatRanges
    End Select
GoTo ExecuteProcedure
ErrorHandle:
    Call system.tackleErrors
ExecuteProcedure:
    Call system.speedOff
End Sub
'3.2 Clear content and format
Public Sub clearAll()
On Error GoTo ErrorHandle
    If Not hasVariables Then GoTo ExecuteProcedure
    Call system.speedOn
    Select Case typeName(targetObject)
        Case "Range"
            Call clearFormatContentRange
    End Select
GoTo ExecuteProcedure
ErrorHandle:
    Call system.tackleErrors
ExecuteProcedure:
    Call system.speedOff
End Sub
