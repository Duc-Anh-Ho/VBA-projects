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
' Declare Variables
Private info As InfoConstants
Private system As SystemUpdate
Private fileSystem As Object
Private userResponse As VbMsgBoxResult
Private customOffset As Byte
Private targetPlace As Object
Private targetObject As Object
Private isLockRatio As Boolean
Private Const MARKER_COLOR As Long = vbRed
Private Const FAKE_BORDER_COLOR As Long = vbMagenta
Private Const MARKER_TEAR_COLOR As Long = vbMagenta
Private Const MARKER_TRANSPARENT_RATE As Byte = 80
Private Const MARKER_LINE_WEIGHT As Byte = 5
Private Const MARKER_TEXT As String = "ARRANGING AREA"
Private Const MARKER_TEXT_FONT As String = "Helvetica"
' Initialize Variables
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
    GoTo ExecuteProcedure
ErrorHandle:
    Call system.tackleErrors
ExecuteProcedure:
End Function

'CONSTRUCTOR
Private Sub Class_Initialize()
    '
End Sub
'DESTRUCTOR
Private Sub Class_Terminate()
    '
End Sub

' ASSESSORS

'1,2,3
Public Property Get hasSelectObjectMode() As Boolean
    Let hasSelectObjectMode = CommandBars("Drawing").controls("Select Objects").State
End Property

' MUTATORS

'1,2,3
Public Property Let letOffset(ByRef newOffset As Byte)
    If newOffset > 100 Then newOffset = 100
    Let customOffset = newOffset
End Property
'1,2,3
Public Property Let letLockRatio(ByRef newLockRatio As Boolean)
    Let isLockRatio = newLockRatio
    'TO-DO: Pop-up Form
    'Let isLockRatio = SnippingToolForm.CheckBoxLockRadio.value
End Property
'2.
Public Property Let selectObjectMode(ByRef objectMode As Boolean)
    If objectMode And Not CommandBars("Drawing").controls("Select Objects").State Then
        Call CommandBars("Drawing").controls("Select Objects").Execute
    ElseIf Not objectMode And CommandBars("Drawing").controls("Select Objects").State Then
        Call CommandBars("Drawing").controls("Select Objects").Execute
    End If
End Property
'2
Public Property Let hasFakeBorder(ByRef isFake As Boolean)
    Dim Sh As Shape
    Dim fakeBorderName As String
    Let fakeBorderName = info.getAuthor & "_fakeBorder_"
    If isFake Then
        For Each Sh In system.ws.Shapes
            With system.ws.Shapes.AddShape( _
                Type:=msoShapeRectangle, _
                left:=Sh.left - MARKER_LINE_WEIGHT / 2, _
                top:=Sh.top - MARKER_LINE_WEIGHT / 2, _
                width:=Sh.width + MARKER_LINE_WEIGHT, _
                height:=Sh.height + MARKER_LINE_WEIGHT)
                .name = fakeBorderName & .name
                .Fill.visible = msoTrue
                .Fill.foreColor.RGB = FAKE_BORDER_COLOR
                .Line.visible = msoFalse
                .zOrder msoSendToBack
            End With
        Next Sh
    Else
        For Each Sh In system.ws.Shapes
            If Sh.name Like (fakeBorderName & "*") Then Call Sh.Delete
        Next Sh
    End If
End Property
'2
Public Property Let hasMarkTargetPlace(ByRef isMarked As Boolean)
    Dim Sh As Shape
    Dim markUpName As String
    Let markUpName = info.getAuthor & "_marker"
    If isMarked Then
        With system.ws.Shapes.AddShape( _
            Type:=msoShapeRectangle, _
            left:=targetPlace.left, _
            top:=targetPlace.top, _
            width:=targetPlace.width, _
            height:=targetPlace.height)
            .name = markUpName
            .Fill.visible = msoTrue
            .Fill.foreColor.RGB = MARKER_COLOR
            .Fill.Transparency = MARKER_TRANSPARENT_RATE / 100
            .Fill.Solid
            .Line.visible = msoTrue
            .Line.foreColor.RGB = MARKER_COLOR
            .Line.Weight = MARKER_LINE_WEIGHT
            .Line.DashStyle = msoLineSysDot
            .TextFrame.Characters.text = MARKER_TEXT
            .TextFrame.Characters.Font.name = MARKER_TEXT_FONT
            .TextFrame.Characters.Font.bold = msoTrue
            .TextFrame.Characters.Font.italic = msoFalse
            .TextFrame.Characters.Font.size = .height / 5
            .TextFrame.Characters.Font.ColorIndex = vbBlack
            .TextFrame2.HorizontalAnchor = msoAnchorCenter
            .TextFrame2.VerticalAnchor = msoAnchorMiddle
            .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        End With
    Else
        For Each Sh In system.ws.Shapes
            If Sh.name = markUpName Then Call Sh.Delete
        Next Sh
    End If
End Property
'3
Public Property Let hasMarkTargetObject(ByRef isMarked As Boolean)
    Dim Sh As Shape
    Dim groupSh As Shape
    Dim tearSh As Shape
    Dim tearName As String
    Dim groupName As String
    Let tearName = info.getAuthor & "_tear_"
    Let groupName = info.getAuthor & "_group_"
    If isMarked Then
        For Each Sh In system.ws.Shapes
            Set tearSh = system.ws.Shapes.AddShape( _
                Type:=msoShapeTear, _
                left:=Sh.left, _
                top:=Sh.top, _
                width:=Sh.width / 20, _
                height:=Sh.width / 20)
            With tearSh
                .Line.visible = msoFalse
                .Fill.foreColor.RGB = MARKER_TEAR_COLOR
                .Adjustments.item(1) = 1.85
                .IncrementRotation 270
                .name = tearName & .name 're-name for marker_delete
            End With
            'Group and rename tear with target object
            Set groupSh = system.ws.Shapes.Range(Array(tearSh.name, Sh.name)).Group
            Let groupSh.name = groupName & groupSh.name
        Next Sh
    Else
        'Loop all shapes in sheet
        For Each groupSh In system.ws.Shapes
            'Check if shape is tear group
            If _
                groupSh.Type = msoGroup _
                And groupSh.name Like (groupName & "*") Then
                'Loop all shapes in group
                For Each Sh In groupSh.GroupItems
                    If Sh.name Like (tearName & "*") Then Call Sh.Delete
                Next Sh
            End If
        Next groupSh
    End If
End Property
'METHODS
'1,2,3
Public Sub lockRatio(ByRef targetObject As Object)
    Select Case TypeName(targetObject)
        Case "Shape"
            Let targetObject.LockAspectRatio = isLockRatio
            Call targetObject.zOrder(msoBringToFront)
        Case Else
            Let targetObject.ShapeRange.LockAspectRatio = isLockRatio
            Call targetObject.ShapeRange.zOrder(msoBringToFront)
    End Select
End Sub
'1,2
Private Sub setTargetPlace(targetOj As Object, Optional hasConfirm As Boolean = True)
    Select Case TypeName(targetOj)
        Case "Range"
            Set targetPlace = targetOj
        Case "DrawingObjects"
            MsgBox _
                Prompt:="Sorry! we've not supported " & TypeName(targetOj) & " yet" _
                    & vbNewLine & "Please select another object type !!!!", _
                Buttons:=vbOKOnly + vbCritical, _
                Title:=info.getAuthor
        Case Else
            'Optional user response
            If hasConfirm Then
                userResponse = MsgBox( _
                    Prompt:="You've selected: '" & targetOj.name & _
                        "' as type: " & TypeName(targetOj) & vbNewLine & _
                        "Do you want to replace " & targetOj.name & "?", _
                    Buttons:=vbYesNoCancel + vbQuestion, _
                    Title:=info.getAuthor)
            Else
                userResponse = vbNo
            End If
            'yes pressed
            If userResponse = vbYes Then
                Set targetPlace = targetOj
            'no pressed
            ElseIf userResponse = vbNo Then
                Set targetPlace = targetOj
            'X or cancel Pressed
            Else
                Set targetPlace = Nothing
            End If
    End Select
End Sub
'1
Private Sub deleteOldPicture()
    If targetPlace Is Nothing Then Exit Sub
    If userResponse = vbYes Then
        Select Case TypeName(targetPlace)
            Case "ChartArea"
                MsgBox _
                    Prompt:="The captured picture is inside the " & TypeName(Selection) _
                        & vbNewLine & "Can't delete it automatically", _
                    Buttons:=vbOKOnly + vbExclamation, _
                    Title:=info.getAuthor
            Case "Range"
                'Do notthing
            Case Else
                Call targetPlace.Delete
        End Select
    End If
End Sub
'2.
Private Sub setTargetObject(ByRef objectName As String)
    Dim Sh As Shape
    Dim shGItem As Shape
        For Each Sh In system.ws.Shapes
            If Sh.Type = msoGroup Then
                For Each shGItem In Sh.GroupItems
                    If shGItem.name = objectName Then Set targetObject = Sh
                Next shGItem
            End If
            If Sh.name = objectName Then Set targetObject = Sh
        Next Sh
End Sub
'1.
Private Sub capturePicture()
    With system.app
        .ScreenUpdating = False
        .CommandBars.ExecuteMso "ScreenClipping" 'important key
    End With
    If TypeName(Selection) = "Picture" Then
        Set targetObject = Selection
'        targetObject.Cut
'    'Paste to A1 as default
'    targetObject = ws.Paste Destination:=ws.Range("A1")
    Else
        MsgBox _
            Prompt:=info.getPrompt & "Failed to capture!!!" _
                & vbNewLine & "Please wait a sec and capture again after clicking OK !!!", _
            Buttons:=vbOKOnly + vbExclamation, _
            Title:=info.getAuthor
        Call capturePicture 'recursion
    End If
End Sub
'2
Private Sub assignOnAction(ByRef script As String)
    Dim Sh As Shape
    For Each Sh In system.ws.Shapes
        Let Sh.OnAction = script
    Next Sh
End Sub
'1,2,3
Private Sub scalePicture(ByVal isLockRatio As Boolean)
    'Scale up when too small
    If targetPlace.height < customOffset * 2 Or _
    targetPlace.width < customOffset * 2 Then
        If TypeName(targetPlace) <> "Range" Then
            Let targetPlace.height = targetPlace.height + customOffset * 2
            Let targetPlace.width = targetPlace.width + customOffset * 2
        End If
    End If
    'Set Lock Rate
    Call lockRatio(targetObject)
    'Algorithm When Lock Rate
    If isLockRatio Then
        If targetPlace.height < targetPlace.width Then
            targetObject.top = targetPlace.top + customOffset
            targetObject.height = targetPlace.height - customOffset * 2
            targetObject.left = targetPlace.left + (targetPlace.width - targetObject.width) / 2
        Else
            targetObject.left = targetPlace.left + customOffset
            targetObject.height = targetPlace.height - customOffset * 2
            targetObject.top = targetPlace.top + (targetPlace.height - targetObject.height) / 2
        End If
        If targetObject.height > targetPlace.height Then
            targetObject.height = targetPlace.height - customOffset * 2
            targetObject.top = targetPlace.top + customOffset
            targetObject.left = targetPlace.left + (targetPlace.width - targetObject.width) / 2
        End If
        If targetObject.width > targetPlace.width Then
            targetObject.width = targetPlace.width - customOffset * 2
            targetObject.left = targetPlace.left + customOffset
            targetObject.top = targetPlace.top + (targetPlace.height - targetObject.height) / 2
        End If
    'Algorithm When Not Lock Rate
    Else
        targetObject.left = targetPlace.left + customOffset
        targetObject.top = targetPlace.top + customOffset
        targetObject.width = targetPlace.width - customOffset * 2
        targetObject.height = targetPlace.height - customOffset * 2
    End If
    'Reset default settings
'    targetObject.ShapeRange.LockAspectRatio = True
    Let letLockRatio = False
    Call lockRatio(targetObject)
    system.app.ScreenUpdating = True
End Sub

'MAIN
'1. Snipping Tool
Public Sub snip()
On Error GoTo ErrorHandle
    If Not hasVariables Then GoTo ExecuteProcedure
    Call system.speedOn
    Call setTargetPlace(Selection)
    If targetPlace Is Nothing Then GoTo ExecuteProcedure
    Call capturePicture
    Call scalePicture(isLockRatio)
    Call deleteOldPicture
GoTo ExecuteProcedure
ErrorHandle:
    Call system.tackleErrors
ExecuteProcedure:
    Let Application.ScreenUpdating = True
    Call system.speedOff
End Sub
'2.1 Assign macro to shape
Public Sub assign()
On Error GoTo ErrorHandle
    If Not hasVariables Then GoTo ExecuteProcedure
    Call system.speedOn
    Call setTargetPlace(Selection, hasConfirm:=False)
    If targetPlace Is Nothing Then GoTo ExecuteProcedure
    Call assignOnAction("ThisWorkbook.Auto_Arrange_Shape")
    Let hasFakeBorder = True
    Let hasMarkTargetPlace = True
    Let selectObjectMode = True
GoTo ExecuteProcedure
ErrorHandle:
    Call system.tackleErrors
ExecuteProcedure:
    Call system.speedOff
End Sub
'2.2 Arrrage Object
Public Sub arrange(ByRef objectName As String)
On Error GoTo ErrorHandle
    If Not hasVariables Then GoTo ExecuteProcedure
    If Not hasSelectObjectMode Then GoTo ExecuteProcedure
    Call system.speedOn
    Call setTargetPlace(Selection, hasConfirm:=False)
    If targetPlace Is Nothing Then GoTo ExecuteProcedure
    Call setTargetObject(Application.Caller)
    Call scalePicture(isLockRatio)
    Call deleteOldPicture
GoTo ExecuteProcedure
ErrorHandle:
    Call system.tackleErrors
ExecuteProcedure:
    Call system.speedOff
End Sub
'2.3
Public Sub clearArrange()
On Error GoTo ErrorHandle
    If Not hasVariables Then GoTo ExecuteProcedure
    Call system.speedOn
    Let hasMarkTargetPlace = False
    Call assignOnAction("")
    Let selectObjectMode = False
    Let hasFakeBorder = False
GoTo ExecuteProcedure
ErrorHandle:
    Call system.tackleErrors
ExecuteProcedure:
    Call system.speedOff
End Sub
' 3. Auto Arrange
Public Sub autoArrange(ByRef isOn As Boolean)
On Error GoTo ErrorHandle
    If Not hasVariables Then GoTo ExecuteProcedure
    Call system.speedOn
    Let hasMarkTargetObject = isOn
'    system.app.Run "PicturesController.Auto_Run_Continuously"
GoTo ExecuteProcedure
ErrorHandle:
    Call system.tackleErrors
ExecuteProcedure:
    Call system.speedOff
End Sub
'3.
Public Sub arrangeToMerge()
    On Error GoTo ErrorHandle
    If Not hasVariables Then GoTo ExecuteProcedure
    Call system.speedOn
    If TypeName(Selection) = "GroupObject" Then
        Call setTargetPlace(Selection.TopLeftCell.MergeArea)
        Call setTargetObject(Selection.name)
        Call scalePicture(isLockRatio)
    End If
GoTo ExecuteProcedure
ErrorHandle:
    Call system.tackleErrors
ExecuteProcedure:
    Call system.speedOff
End Sub




