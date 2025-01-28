VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "InternetConnector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
    ' Check README.md for more information
Option Explicit
'Declare Variables
Private info As InfoConstants
Private system As SystemUpdate
Private PWShell As PowerShellController
'Private fileSystem As Object
Private userResponse As VbMsgBoxResult
Private Const PING As String = "ping "
Private Const NUM_ECHO As String = "-n 1 " 'Number of echo requests to send
Private Const WAIT_TIME As String = "-w 500 " '1s: Timeout in milliseconds to wait for each reply
Private Const DEFAULT_LINK As String = "google.com"
Private Const WIFI_PROFILE_SCRIPTS As String = "netsh wlan show profile"
Private Const WIFI_NEWWORK_SCRIPTS As String = "netsh wlan show network"
Private Const REGEX_PATTERN As String = "'\:(.+)$'" ' Or """\:(.+)$""" ' """" = Chr(34) = 1 double quote
Private Const WIFI_TABLE_SCRIPTS As String = _
    "(netsh wlan show profiles) " & _
    "| Select-String -Pattern '\:(.+)$' " & _
    "| ForEach-Object -Process { " & _
        "$wifiName = $PSItem.Matches.Groups[1].Value.Trim();" & _
        "Write-Output $wifiName } " & _
    "| ForEach-Object -Process { " & _
        "$wifiFullInfo = (netsh wlan show profiles name = $wifiName key=clear);" & _
        "Write-Output -input $wifiFullInfo } " & _
    "| Select-String -Pattern 'Key Content\W+\:(.+)$' " & _
    "| ForEach-Object -Process { " & _
        "$wifiPW = $PSItem.Matches.Groups[1].Value.Trim();" & _
        "Write-Output $wifiPW } " & _
    "| ForEach-Object -Process { " & _
        "$wifiTable = [PSCustomObject]@{ " & _
            "WIFI_NAMES = $wifiName;" & _
            "WIFI_PASSWORDS = $wifiPW };" & _
         "Write-Output $wifiTable}"
Private Const EXPORT_TXT_SCRIPTS As String = _
    "| Format-Table -AutoSize | Out-File -FilePath  '"
Private Const TXT_FILE_NAME As String = "\wifi.txt'; "
Private Const EXPORT_CSV_SCRIPTS As String = _
    " | Export-Csv -Path '"
Private Const CSV_FILE_NAME As String = "\wifi.csv' -NoTypeInformation;"
Private Const EXPORT_JSON_SCRIPTS As String = _
    " | ConvertTo-Json -Depth 100 | Out-File -FilePath '"
Private Const JSON_FILE_NAME As String = "\wifi.json';"
Enum exportType
    TXT = 0
    CSV = 1
    JSON = 2
End Enum
'Initialize Variables
Private Function hasVariables() As Boolean
On Error GoTo ErrorHandle
    Set info = New InfoConstants
    Set system = New SystemUpdate
    Set PWShell = New PowerShellController
    Let hasVariables = True
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
    Set info = Nothing
    Set system = Nothing
End Sub
'ASSESSORS
Private Property Get checkConnect(Optional ByRef link As String = DEFAULT_LINK) As Boolean
    Dim returnedCode As Byte
    Dim command As String
    Let command = _
            PING & _
            NUM_ECHO & _
            WAIT_TIME & _
            link
    Let returnedCode = PWShell.runScript(command)
    If returnedCode = 0 Then
        Let checkConnect = True
    Else
        Let checkConnect = False
    End If
End Property
'MUTATORS
'METHODS
Private Sub saveWifiAs( _
    Optional ByRef fileType As Byte = exportType.TXT, _
    Optional ByRef useThisFolder As Boolean = False)
    Dim command As String
    Dim extensionScript As String
    Dim folderScript As String
    Dim returnTxt As String
    If useThisFolder Then
        Let folderScript = ThisWorkbook.path
    Else
        Let folderScript = system.getFolder
        'Do nothing when didn't choose folder
        If folderScript = vbNullString Then Exit Sub
    End If
    Select Case fileType
        Case exportType.TXT
            Let extensionScript = _
                EXPORT_TXT_SCRIPTS & _
                folderScript & _
                TXT_FILE_NAME
        Case exportType.CSV
            Let extensionScript = _
                EXPORT_CSV_SCRIPTS & _
                folderScript & _
                CSV_FILE_NAME
        Case exportType.JSON
            Let extensionScript = _
                EXPORT_JSON_SCRIPTS & _
                folderScript & _
                JSON_FILE_NAME
        Case Else
             Exit Sub
    End Select
    Let command = _
            PWShell.createPWShellCommand & _
            WIFI_TABLE_SCRIPTS & _
            extensionScript
   Let returnTxt = PWShell.executeScript(command) 'TO-DO: refactory text return
End Sub
'MAIN
'Check internet connect
Public Function isConnect(Optional ByRef link As String = DEFAULT_LINK) As Boolean
On Error GoTo ErrorHandle
    If Not hasVariables Then GoTo ExecuteProcedure
    Call system.speedOn
    Let isConnect = checkConnect(link)
GoTo ExecuteProcedure
ErrorHandle:
    Call system.tackleErrors
ExecuteProcedure:
    Call system.speedOff
End Function

Public Sub saveWifiAsTxt(Optional ByRef useThisFolder As Boolean = False)
On Error GoTo ErrorHandle
    If Not hasVariables Then GoTo ExecuteProcedure
    Call system.speedOn
    Call saveWifiAs(exportType.TXT, useThisFolder)
GoTo ExecuteProcedure
ErrorHandle:
    Call system.tackleErrors
ExecuteProcedure:
    Call system.speedOff
End Sub

Public Sub saveWifiAsJson(Optional ByRef useThisFolder As Boolean = False)
On Error GoTo ErrorHandle
    If Not hasVariables Then GoTo ExecuteProcedure
    Call system.speedOn
    Call saveWifiAs(exportType.JSON, useThisFolder)
GoTo ExecuteProcedure
ErrorHandle:
    Call system.tackleErrors
ExecuteProcedure:
    Call system.speedOff
End Sub

Public Sub saveWifiAsCsv(Optional ByRef useThisFolder As Boolean = False)
On Error GoTo ErrorHandle
    If Not hasVariables Then GoTo ExecuteProcedure
    Call system.speedOn
    Call saveWifiAs(exportType.CSV, useThisFolder)
GoTo ExecuteProcedure
ErrorHandle:
    Call system.tackleErrors
ExecuteProcedure:
    Call system.speedOff
End Sub
