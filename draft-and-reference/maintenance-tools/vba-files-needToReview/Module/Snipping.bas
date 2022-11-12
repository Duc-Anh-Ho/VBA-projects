Attribute VB_Name = "Snipping"
Option Explicit
' Copyright - BuiDanhVN indie software
' Ver.2020.0.3.1
Public Sub Snipping()
'0.initialization
'0.1.offset
    Dim left_offset As Integer
    Dim top_offset As Integer
    Let left_offset = 10 'default option
    Let top_offset = 10 'default option
'0.2.main object
    Dim pic_captured_screen As Picture 'main object
'0.3.target adress
    Dim wb_target As Workbook
    Dim ws_target As Worksheet
    Dim wb_name As String
    Set wb_target = Application.ActiveWorkbook 'picked option
    On Error GoTo error_handle 'No picked
    Set ws_target = wb_target.ActiveSheet 'picked option
    
'Case_1.1 selected range
    'Range setup
    If TypeName(Selection) = "Range" Then
        Dim rg_target As Range
        Set rg_target = Selection 'picked option
        Let wb_name = ThisWorkbook.Name
    '1.Cutting screen
        With Application
            .ScreenUpdating = False
            .CommandBars.ExecuteMso "ScreenClipping" 'important key
        End With
        Selection.Cut
    '2.Pasting screen
        ws_target.Paste Destination:=rg_target
    '3.Resizing captured pictureon
        'On Error GoTo error_handle
        Set pic_captured_screen = Selection 'initilize captured picture
    ''3.1.Check merged cell
    '    If rg_target.MergeCells Then
    '        Set rg_target = ws_target.Range(rg_target.MergeArea.Address)
    '    Else
    '        'pass
    '    End If
    '3.2.Set size (if not move with cell)
        With pic_captured_screen
            .Left = rg_target.Left + left_offset
            .Top = rg_target.Top + top_offset
        End With
        With pic_captured_screen.ShapeRange
            If .Height < .Width Then
                .Width = rg_target.Width - left_offset * 2
            Else
                .Height = rg_target.Height - top_offset * 2
            End If
        End With
'Case_1.2 selected DrawingObjects (Group)
    ElseIf TypeName(Selection) = "DrawingObjects" Then
        MsgBox "Sorry! we do not support " & TypeName(Selection) & " yet" _
            & vbNewLine & "Please select another object type !!!!", _
            vbOKOnly + vbCritical, _
            "SNIPPING TOOL"
'Case_1.3 selected object (picture, shape, chart, ...)
    Else
        'Confirm notification
        Dim replace_confirm As VbMsgBoxResult
        replace_confirm = MsgBox("You selected: '" & Selection.Name & "' as type: " & TypeName(Selection) _
            & vbNewLine & "Your new sniped picture will scale and place above old object", _
            vbOKCancel + vbQuestion, _
            "SNIPPING TOOL")
        'Ok Pressed
        If replace_confirm = vbOK Then
            'Object size setup
            Dim pic_top As Integer: Let pic_top = Selection.Top
            Dim pic_left As Integer: Let pic_left = Selection.Left
            Dim pic_height As Integer: Let pic_height = Selection.Height
            Dim pic_width As Integer: Let pic_width = Selection.Width
            '1.Cutting screen
            On Error GoTo error_handle
            With Application
                .ScreenUpdating = False
                .CommandBars.ExecuteMso "ScreenClipping" 'important key
            End With
            Selection.Cut
            '2.Pasting screen
            ws_target.Paste Destination:=ws_target.Range("A1")
            '3.Resizing captured pictureon
            Set pic_captured_screen = Selection 'initilize captured picture
            '3.1.Set size (if replace object)
            With pic_captured_screen
                .ShapeRange.LockAspectRatio = msoFalse
                .Left = pic_left
                .Top = pic_top
                .Width = pic_width
                .Height = pic_height
                .ShapeRange.LockAspectRatio = msoTrue
            End With
        'Cancel pressed
        Else
            End
        End If
    End If
'XX.Closing
    With Application
        '.CutCopyMode = False
        .ScreenUpdating = True
    End With
error_handle:
    If Err.Number = 91 Then
        MsgBox "Nothing selected!" & vbNewLine & _
        "Please select range or object!", vbOKOnly + vbCritical
        Exit Sub
    Else
        'MsgBox Err.Description, vbOKOnly + vbCritical
        Exit Sub
    End If
End Sub
'Shortcut_create
Public Sub shortcut_add()
    Application.OnKey "^q", "snipping.snipping"  'Assign Ctrl + q to run Snipping Tool
End Sub
'Shortcut_delete
Public Sub shortcut_delete()
    Application.OnKey "^q"  'Delete assigned Ctrl + q
End Sub
