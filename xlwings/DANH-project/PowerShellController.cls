VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PowerShellController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Check README.md for more information
Option Explicit
'Declare Variables
Private system As SystemUpdate
Private powerShell As Object
Private fileStream As Object 'Stream Object
Private userResponse As VbMsgBoxResult
Private Const CREATE_COMMAND_SCRIPT As String = _
    "PowerShell.exe -NoLogo -WindowStyle Hidden -Command "
Enum PowerShellStt
    Running = 0
    Successful = 1
    Failed = 2
End Enum
Private Const WINDOW_STYLE As Integer = 0 ' 0: Hides the window and activates another window.
Private Const WAIT_ON_RETURN As Boolean = True
'Constructor
Private Sub Class_Initialize()
    Set system = New SystemUpdate
    Set powerShell = system.createPowerShell
End Sub
'Destructor
Private Sub Class_Terminate()
    Set system = Nothing
    Set powerShell = Nothing
End Sub
Public Property Get createPWShellCommand() As String
    Let createPWShellCommand = CREATE_COMMAND_SCRIPT
End Property

Public Property Get runScript(ByRef scrpit As String) As Byte
    Let runScript = powerShell.Run( _
        scrpit, _
        WINDOW_STYLE, _
        WAIT_ON_RETURN)
End Property

' References: learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/textstream-object
Public Property Get executeScript(ByRef scrpit As String) As String
    Dim executionCmd As Object  ' Object = WshExec: Window Shell Execute
    Dim returnTxt As String
    Set executionCmd = powerShell.Exec(scrpit)
    ' Delay for waiting exucte
    Do While executionCmd.Status = PowerShellStt.Running
        Application.Wait (Now + TimeValue("00:00:01")) '<--1s
    Loop
    ' Check return successfull or not
    Select Case executionCmd.Status
        'Successful
        Case PowerShellStt.Successful
            While Not executionCmd.StdOut.AtEndOfStream
                Let returnTxt = executionCmd.StdOut.ReadAll()
            Wend
        'Failed
        Case PowerShellStt.Failed
            Let returnTxt = executionCmd.StdErr.ReadAll()
     End Select
    Let executeScript = returnTxt
End Property
