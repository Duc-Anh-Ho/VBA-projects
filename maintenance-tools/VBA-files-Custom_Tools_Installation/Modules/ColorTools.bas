Attribute VB_Name = "ColorTools"
Option Explicit
' Copyright - BuiDanhVN indie software
' Ver.2020.0.1.1
Sub Invert_Color()
Attribute Invert_Color.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim SelectedRange   As Range
    'Dim FillColor       As Long
    'Dim FontColor       As Long
    'Dim HexColor        As String
    Dim RGBColor        As Long
    Dim RedColor        As Integer
    Dim GreenColor      As Integer
    Dim BlueColor       As Integer
    
    Application.ScreenUpdating = False
    For Each SelectedRange In Selection
        'FillColor = SelectedRange.Interior.Color
        'FontColor = SelectedRange.Font.Color
        
        'HexColor = Right("000000" & Hex(SelectedRange.Interior.Color), 6)
        'MsgBox Right("000000" & Hex(SelectedRange.Interior.Color), 6)
        
        
        RGBColor = SelectedRange.Interior.Color
        'MsgBox (Color)
        
        RedColor = RGBColor Mod 256
        GreenColor = RGBColor \ 256 Mod 256
        BlueColor = RGBColor \ 65536 Mod 256
        'MsgBox (RedColor & " " & GreenColor & " " & BlueColor)
        
        RedColor = 255 - RedColor
        GreenColor = 255 - GreenColor
        BlueColor = 255 - BlueColor
        'MsgBox (RedColor & " " & GreenColor & " " & BlueColor)
        
        SelectedRange.Interior.Color = RGB(RedColor, GreenColor, BlueColor)
        
        '---------------
        RGBColor = SelectedRange.Font.Color
        RedColor = RGBColor Mod 256
        GreenColor = RGBColor \ 256 Mod 256
        BlueColor = RGBColor \ 65536 Mod 256
        RedColor = 255 - RedColor
        GreenColor = 255 - GreenColor
        BlueColor = 255 - BlueColor
        SelectedRange.Font.Color = RGB(RedColor, GreenColor, BlueColor)
        
        '---------------
        RGBColor = SelectedRange.Borders.Color
        RedColor = RGBColor Mod 256
        GreenColor = RGBColor \ 256 Mod 256
        BlueColor = RGBColor \ 65536 Mod 256
        RedColor = 255 - RedColor
        GreenColor = 255 - GreenColor
        BlueColor = 255 - BlueColor
        SelectedRange.Borders.Color = RGB(RedColor, GreenColor, BlueColor)
    Next SelectedRange
    Application.ScreenUpdating = True
End Sub

'Shortcut_add
Sub shortcut_add()
    Application.OnKey "^{i}", "ColorTools.Invert_Color" 'Shortcut: ctrl + shift + I
End Sub
'Shortcut_delete
Sub shortcut_delete()
    Application.OnKey "^{i}"  'Delete assigned ctrl + shift + I
End Sub

