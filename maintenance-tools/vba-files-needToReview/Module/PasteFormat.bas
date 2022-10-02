Attribute VB_Name = "PasteFormat"
' Copyright - BuiDanhVN indie software
' Ver.2020.0.1.1

Option Explicit
'Tutorial: Add with shortcut (like word and power point)
Public copied_range As Range
Sub Paste_Format()
Attribute Paste_Format.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim apply_object As Object
    Dim apply_objects As Object
    Application.ScreenUpdating = False
'Case_1 Selected as Ranges
    If TypeName(Selection) = "Range" Then
        If Application.CutCopyMode Then 'case copy
            On Error GoTo Err_handler
            Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
'            Application.CutCopyMode = False 'Optional: Stop copy format
        End If
'Case_2 Selected as DrawingObject (Group of object)
    ElseIf TypeName(Selection) = "DrawingObjects" Then
        Set apply_objects = Selection
        ActiveSheet.Paste
        Selection.ShapeRange.PickUp
        Selection.Delete
        For Each apply_object In apply_objects
            'On Error GoTo Err_handler
            ActiveSheet.Shapes(apply_object.Name).Apply
        Next apply_object
'Case_3 Selected as another object (picture, shape, chart...)
    Else
        Set apply_object = Selection
        ActiveSheet.Paste
        On Error GoTo Err_handler
        Selection.ShapeRange.PickUp
        Selection.Delete
        ActiveSheet.Shapes(apply_object.Name).Apply
    End If
    Application.ScreenUpdating = True

Err_handler:
    If Err.Number = 1004 Then
        MsgBox "Format painter function still not working on cut mode. Update in next version."
        Exit Sub
    ElseIf Err.Number = 438 Then
    ' Case: format painter choosen is not a shape
        Resume Next
    Else
        Application.Speech.Speak (Err.Description)
        'Debug.Print ("Error No." & Err.Number)
    End If
End Sub

'Shortcut_create
Sub shortcut_add()
    Application.OnKey "+^v", "PasteFormat.Paste_Format" 'Shortcut: ctrl + shift +v
End Sub

'Shortcut_delete
Sub shortcut_delete()
    Application.OnKey "+^v"  'Delete assigned ctrl + shift +v
End Sub
