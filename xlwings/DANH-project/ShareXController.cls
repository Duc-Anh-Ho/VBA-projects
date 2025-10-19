VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ShareXController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Const SHAREX_EXEC As String = "ShareX.exe"
Private Const SHAREX_PATH As String = "D:\Program Files\ShareX\" ' TODO: Make configurable with ENV
Private Const ARG_SILENT As String = "-silent"
Private Const ARG_PORTABLE As String = "-portable"
Private Const ARG_WORKFLOW As String = "-workflow"
Private Const WMI_SCOPE As String = "winmgmts:\\.\root\cimv2"
Private Const QUERY_PREFIX As String = "Select * from Win32_Process Where ProcessId = "
Private Const SHELL_MODE As Byte = vbHide
Private Const DEFAULT_TIMEOUT_MS As Integer = 5000
Private Const DEFAULT_SLEEP_MS As Integer = 100

Private shellPID As Long
Private workflow As String

Private Property Let letShellPID(ByRef PID As Long): Let shellPID = PID: End Property
Public Property Let letWorkflow(ByVal name As String): Let workflow = name: End Property

Private Property Get getShellPID() As Long: Let getShellPID = shellPID: End Property
Public Property Get getWorkflow() As String: Let getWorkflow = workflow: End Property

' TODO: Create Unitity cls
Private Function addQuote(ByVal text As String) As String
    Let addQuote = """" & text & """"
End Function

Private Function isCaptured() As Boolean
    Dim CMD As String
    Dim exePath As String
    Let exePath = SHAREX_PATH & SHAREX_EXEC
    Let CMD = addQuote(exePath) & Space(1) & _
        ARG_WORKFLOW & Space(1) & addQuote(workflow) & Space(1) & _
        ARG_SILENT & Space(1) & ARG_PORTABLE
    Let letShellPID = Shell(CMD, SHELL_MODE)
    Let isCaptured = (shellPID > 0)
End Function

Public Function START(Optional ByVal timeoutMs As Integer = DEFAULT_TIMEOUT_MS)
    If isCaptured Then Call waitForFinish(timeoutMs)
End Function

Private Sub waitForFinish(Optional ByVal timeoutMs As Integer = DEFAULT_TIMEOUT_MS)
    Dim objWMI As Object, processes As Object, proc As Object
    Dim isRunning As Boolean
    Dim tStart As Single: Let tStart = Timer
    Set objWMI = GetObject(WMI_SCOPE)
    Do
        Let isRunning = False
        Set processes = objWMI.ExecQuery(QUERY_PREFIX & getShellPID)
        For Each proc In processes
            Let isRunning = True
        Next
        DoEvents
        Call Preprocessor.Sleep(DEFAULT_SLEEP_MS)
        If (Timer - tStart) * 1000 > timeoutMs Then Exit Do
    Loop While isRunning
End Sub


