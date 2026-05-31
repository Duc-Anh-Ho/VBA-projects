VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Utils_Scripts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Check README.md for more information
Option Explicit
' --- Dependencies ---
Private info As InfoConstants
' --- Constant ---
Private Const FILE_SYSTEM_SCRIPT As String = "Scripting.FileSystemObject"
Private Const WSHELL_SCRIPT As String = "WScript.Shell"
Private Const ASHELL_SCRIPT As String = "Shell.Application"
Private Const FILE_STREAM_SCRIPT As String = "ADODB.Stream"
Private Const CONNECTION_SCRIPT As String = "ADODB.Connection"
Private Const RECORD_SET_SCRIPT As String = "ADODB.Recordset"
Private Const CDO_CONFIG_SCRIPT As String = "CDO.Configuration"
Private Const CDO_MESS_SCRIPT As String = "CDO.Message"
Private Const HTML_FILE_SCRIPT As String = "HtmlFile"
Private Const CLIPBOARD_SCRIPT As String = "New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}"
Private Const DICTIONARY_SCRIPT As String = "Scripting.Dictionary"
Private Const HTTPS_SCRIPT As String = "WinHttp.WinHttpRequets.5.1"
Private Const WINSCP_SCRIPT As String = "WinSCP.Session"
Private Const REGEX_SCRIPT As String = "VBScript.RegExp"

' --- Member ---
Private m_fileSystem As Object
Private m_wShell As Object
Private m_aShell As Object
Private m_fileStream As Object
Private m_cdoConfig As Object
Private m_cdoMess As Object
Private m_htmlFile As Object
Private m_clipboard As Object
Private m_dictionary As Object
Private m_https As Object
Private m_winscp As Object
Private m_connection As Object
Private m_recordset As Object
Private m_regex As Object

' --- Constructor ---
' --- Destructor ---
Private Sub Class_Terminate()
    ' Clean up all created objects
    Set m_fileSystem = Nothing
    Set m_wShell = Nothing
    Set m_aShell = Nothing
    Set m_fileStream = Nothing
    Set m_cdoConfig = Nothing
    Set m_cdoMess = Nothing
    Set m_htmlFile = Nothing
    Set m_clipboard = Nothing
    Set m_dictionary = Nothing
    Set m_https = Nothing
    Set m_winscp = Nothing
    Set m_connection = Nothing
    Set m_recordset = Nothing
    Set m_regex = Nothing
End Sub
' --- Helper Methods ---
Private Function safeCreate(progId As String, ByRef obj As Object) As Object
    On Error GoTo ErrHandler
    If obj Is Nothing Then Set obj = VBA.Interaction.CreateObject(progId)
    Set safeCreate = obj
    Exit Function
ErrHandler:
    MsgBox "Failed to create COM object: " & progId, vbCritical
    ERR.Clear
    Set safeCreate = Nothing
End Function

' --- Properties ---
Friend Property Get FileSystem() As Object
    Set FileSystem = safeCreate(FILE_SYSTEM_SCRIPT, m_fileSystem)
End Property

Friend Property Get WShell() As Object
    Set WShell = safeCreate(WSHELL_SCRIPT, m_wShell)
End Property

Friend Property Get AShell() As Object
    Set AShell = safeCreate(ASHELL_SCRIPT, m_aShell)
End Property

Friend Property Get FileStream() As Object
    Set FileStream = safeCreate(FILE_STREAM_SCRIPT, m_fileStream)
End Property

Friend Property Get Connection() As Object
    Set Connection = safeCreate(CONNECTION_SCRIPT, m_connection)
End Property

Friend Property Get RecordSet() As Object
    Set RecordSet = safeCreate(RECORD_SET_SCRIPT, m_recordset)
End Property

Friend Property Get CDOConfig() As Object
    Set CDOConfig = safeCreate(CDO_CONFIG_SCRIPT, m_cdoConfig)
End Property

Friend Property Get CDOMess() As Object
    Set CDOMess = safeCreate(CDO_MESS_SCRIPT, m_cdoMess)
End Property

Friend Property Get HTMLFile() As Object
    Set HTMLFile = safeCreate(HTML_FILE_SCRIPT, m_htmlFile)
End Property

Friend Property Get Clipboard() As Object
    Set Clipboard = safeCreate(CLIPBOARD_SCRIPT, m_clipboard)
End Property

Friend Property Get Dictionary() As Object
    Set Dictionary = safeCreate(DICTIONARY_SCRIPT, m_dictionary)
End Property

Friend Property Get HTTPS() As Object
    Set HTTPS = safeCreate(HTTPS_SCRIPT, m_https)
End Property

Friend Property Get WinSCP() As Object
    Set WinSCP = safeCreate(WINSCP_SCRIPT, m_winscp)
End Property

Friend Property Get RegexExpression() As Object
    Set RegexExpression = safeCreate(REGEX_SCRIPT, m_regex)
End Property

