Attribute VB_Name = "SheetsController"
' Copyright - BuiDanhVN indie software
' Ver.2020.0.1.1

' Including 3 sub:
' 1 - Add sheets as selection
' 2 - Delete all except active sheet
' 3 - List all sheet names and add hyperlinks for each sheet
' 4 - Unhide all sheet
' 5 - Hide all sheet
' 6 - Very Hide all sheet (can't unhide normally)
' 7 - Shortcut
Option Explicit
Dim Sh As Worksheet ' public variable
Dim reponse As VbMsgBoxResult 'for yes/no/cancel msgbox
'for list_sheet and delete
Const title_no As String = "No."
Const title_name As String = "SHEET NAMES"
Public previous_sheet As Worksheet 'for deleting listsheet
'1-Tutorial: Select a range then run, the names of new sheets is value in that range
Public Sub AddSheets()
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    Dim sh_current As Worksheet
    Dim TenSheet As Range
    Dim VungTen As Range
    Set sh_current = ActiveSheet
    Set VungTen = Selection
    For Each TenSheet In VungTen
        If TenSheet.Value <> "" Then ' exclude blank range
            On Error GoTo ErrorHandler ' catch  error
            Sheets.Add(After:=Sheets(Sheets.count)).name = TenSheet 'add sheet and change name
            'ActiveWindow.DisplayGridlines = False 'turn off grid (optional)
        End If
    Next TenSheet
    sh_current.Activate 'return to origin sheet
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
ErrorHandler:
        If Err.Number = 1004 Then '1004: dupliticate sheet name
            reponse = MsgBox("Sheet """ & TenSheet & """ already existed." & _
            vbNewLine & "OR:" & _
            vbNewLine & Err.Description, vbOKCancel + vbExclamation)
            If reponse = vbOK Then
                ActiveSheet.Delete 'delete if can't name the she
                Resume Next
            Else
                ActiveSheet.Delete
                sh_current.Activate 'return to origin sheet
                Application.DisplayAlerts = True
                Application.ScreenUpdating = True
            End If
        End If
End Sub

'2-Turtorial: Delete all sheets except activate one
Sub DeleteSheets()
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    reponse = MsgBox("All sheets except activative one will be delete." & vbNewLine & _
    "Excel CAN NOT UNDO and will autosave." & vbNewLine & _
    "Please check carefully before click OK.", _
    vbOKCancel + vbExclamation, _
    "Confirmation")
    If reponse = vbOK Then
        ActiveWorkbook.Save 'Save
        For Each Sh In Worksheets 'loop thought all sheets
            If Sh.name <> ActiveSheet.name Then 'exclude activesheet
                Sh.Delete 'delete sheet
            End If
        Next Sh
    End If
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub

'3-Turtorial: ON/OFF - List all sheet names and add hyperlink
Sub ListSheets()
    Application.ScreenUpdating = False
    Dim sheet_no As Integer
    Set Sh = ActiveSheet
    With Sh
        If .Cells(1, 1) = title_no And .Cells(1, 2) = title_name Then
            ''Xoa neu o A1 la STT
            .Columns(1).Delete
            .Columns(1).Delete
            Exit Sub
        Else
            'Them 2 cot va dat 2 tieu de neu chua ton tai
            'Them 2 cot
            .Columns(1).Insert
            .Columns(1).Insert
            'Them 2 tieu de
            .Cells(1, 1) = title_no
            .Cells(1, 2) = title_name
            'Them STT , TENSHEET  va HYPERLINK tat ca cac sheet
            For sheet_no = 1 To Sheets.count ''Vong lap theo sl sheets
                ''Cach 1: them bang dia chi link
                .Cells(sheet_no + 1, 1) = sheet_no
                .Cells(sheet_no + 1, 2) = Sheets(sheet_no).name
                Sheets(sheet_no) _
                .Hyperlinks.Add Anchor:=Cells(sheet_no + 1, 2), _
                Address:="", _
                SubAddress:="'" & Sheets(sheet_no).name & "'" & "!C1", _
                ScreenTip:=CStr("Click to go to Sheet No." & sheet_no & ": " & Sheets(sheet_no).name)
    '            'Cach 2: dung ham Hyperlink (**de select dc sheet sau do moi co the luu ten trong Workbook_SheetDeactivate)
    '            Cells(sheet_no + 1, 1) = sheet_no
    '            Cells(sheet_no + 1, 2).Formula = "=HYPERLINK(""#MyFunctionkClick()"", ""Run a function..."")"
            Next sheet_no
            With .Range("A1:B1").Font
                .Bold = True
                .Size = 14
                .name = "Calibri"
            End With
            With .Range("B2", Range("B2").End(xlDown))
                .Font.ColorIndex = xlAutomatic
                .Font.Underline = xlUnderlineStyleDouble
                .Font.Italic = True
                .Font.name = "Calibri"
                .Font.Size = 12
                .Font.Bold = False
            End With
            For sheet_no = 1 To Sheets.count ''Vong lap theo sl sheets
                ''Cach 1: them bang dia chi link
                If Sh.Cells(sheet_no + 1, 2) = ActiveSheet.name Then
                    Sh.Cells(sheet_no + 1, 2).Font.Size = 13
                    Sh.Cells(sheet_no + 1, 1).Font.Size = 13
                    Sh.Cells(sheet_no + 1, 2).Font.Bold = True
                    Sh.Cells(sheet_no + 1, 1).Font.Bold = True
                End If
            Next sheet_no
            With .Columns("A:B")
                .AutoFit
            End With
            ''lam mau
        End If
    End With
    
'    ' take this sheet for deleting if wanna delete after click in current sheet
'    Set previous_sheet = ActiveSheet
    Application.ScreenUpdating = True
End Sub
'3.2-Delete list sheet
Public Sub ListSheets_delete(sheet_name As Worksheet)
    With sheet_name
On Error GoTo ErrHandler
        If .Cells(1, 1) = title_no And .Cells(1, 2) = title_name Then
        ''Xoa neu o A1 la STT
            .Columns(1).Delete
            .Columns(1).Delete
        End If
    End With
ErrHandler:
    If Err.Number = -2147221080 Then
    'Bay loi khi xoa sheet khong the tra lai ten sheet da xoa
        Resume Next
    End If
End Sub

'4-Turtorial: Show all hide sheets
Sub Unhide_All_Sheets()
    For Each Sh In ActiveWorkbook.Worksheets
        Sh.Visible = xlSheetVisible
    Next Sh
End Sub

'5-Turtorial: Hide all sheets
Sub Hide_All_Sheet()
    For Each Sh In Application.ActiveWorkbook.Worksheets
        If Sh.name <> Application.ActiveSheet.name Then
            Sh.Visible = xlSheetHidden
        End If
    Next Sh
End Sub

'6-Turtorial: Very Hide all sheets
Sub Very_Hide_All_Sheet()
    For Each Sh In Application.ActiveWorkbook.Worksheets
        If Sh.name <> Application.ActiveSheet.name Then
            Sh.Visible = xlSheetVeryHidden
        End If
    Next Sh
End Sub

'7-Shortcut
'Shortcut_list_sheet
Sub shortcut_add()
    Application.OnKey "^l", "SheetsController.ListSheets"  'Assign "Ctrl + l" to run add sheets
End Sub
'Shortcut_delete_list_sheet
Sub shortcut_delete()
    Application.OnKey "^l" 'Delete assigned "Ctrl + l" to run add sheets
End Sub


    
