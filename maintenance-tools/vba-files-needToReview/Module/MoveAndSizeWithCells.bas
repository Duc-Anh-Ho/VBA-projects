Attribute VB_Name = "MoveAndSizeWithCells"
' Copyright - BuiDanhVN indie software
' Ver.2020.0.1.1
Option Explicit

Sub Move_And_Size_With_Cells()
Attribute Move_And_Size_With_Cells.VB_ProcData.VB_Invoke_Func = " \n14"
  Dim shp As Shape
  For Each shp In ActiveSheet.Shapes
    shp.Placement = xlMoveAndSize
  Next shp
End Sub
'Shortcut_add
Sub shortcut_add()
    Application.OnKey "+^{m}", "MoveAndSizeWithCells.Move_And_Size_With_Cells" 'Shortcut: ctrl + shift +m
End Sub
'Shortcut_delete
Sub shortcut_delete()
    Application.OnKey "+^{m}"  'Delete assigned ctrl + shift +m
End Sub


'' code cu~'
'Sub Set_Move_And_Size_With_Cells_Old()
'    Dim xPic As Picture
'    On Error Resume Next
'    Application.ScreenUpdating = False
'    For Each xPic In ActiveSheet.Pictures
'        xPic.Placement = xlMoveAndSize
'    Next
'    Application.ScreenUpdating = True
'
'End Sub



