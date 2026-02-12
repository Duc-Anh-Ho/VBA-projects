VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Utils_Error"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Check README.md for more information
Option Explicit
' --- Constant ---
Private Const LOG_FOLDER As String = "C:\DT_Logs"
Private Const LOG_FILE As String = "C:\DT_Logs\AppError.log"
Public Enum ECODE
    ' System level
    Err_Unknown = 1000
    Err_FileNotFound = 1001
    Err_FolderNotFound = 1002
    Err_PathNotFound = 1003
    Err_PermissionDenied = 1004
    Err_OperationNotSupported = 1005
    Err_MissingLibrary = 1006

    ' IO operations
    Err_CopyFailed = 2001
    Err_ExtractFailed = 2002
    Err_SaveFailed = 2003

    ' Validation
    Err_InvalidParameter = 3001
    Err_EmptyValue = 3002
End Enum

' --- Helper Functions ---
Private Function GetMessage(ByVal errorCode As ECODE) As String
    Select Case errorCode
        Case Err_FileNotFound: Let GetMessage = "File not found."
        Case Err_FolderNotFound: Let GetMessage = "Folder not found."
        Case Err_PathNotFound: Let GetMessage = "Path not found."
        Case Err_CopyFailed: Let GetMessage = "Copy operation failed."
        Case Err_InvalidParameter: Let GetMessage = "Invalid parameter provided."
        Case Err_OperationNotSupported: Let GetMessage = "Operation not supported."
        Case Err_MissingLibrary: Let GetMessage = "DLL file not found."
        Case Else: Let GetMessage = "Unknown error occurred."
    End Select
End Function

' --- Main Functions ---
Friend Sub Raise( _
    ByVal errorCode As ECODE _
    , ByVal sourceName As String _
    , ByVal message As String _
)
    Call Log(LOG_FILE, errorCode, sourceName, message)
    Call Err.Raise( _
        Number:=errorCode _
        , Source:=sourceName _
        , Description:=message _
    )
End Sub

Friend Function ToText( _
    ByVal errorCode As ECODE _
    , ByVal sourceName As String _
    , ByVal message As String _
) As String
    Let ToText = "[" & Format(Now, "yyyy-mm-dd hh:nn:ss") & "] " & _
        "[Error " & CStr(errorCode) & "] " & sourceName & " | " & message
End Function

Friend Sub Log( _
    ByVal filePath As String _
    , ByVal errorCode As ECODE _
    , ByVal sourceName As String _
    , ByVal message As String _
)
    Dim scripts As Utils_Scripts: Set scripts = New Utils_Scripts
    Dim fileSystem As Object: Set fileSystem = scripts.FileSystem
    If Not fileSystem.FolderExists(LOG_FOLDER) Then Call fileSystem.CreateFolder(LOG_FOLDER)
    Dim fileNumber As Integer: Let fileNumber = FreeFile
    Open filePath For Append As #fileNumber
    Print #fileNumber, ToText(errorCode, sourceName, message)
    Close #fileNumber
End Sub

Friend Sub Fail(ByVal errorCode As ECODE, ParamArray sourceNames() As Variant)
    Dim fullSource As String
    Dim i As Long: For i = LBound(sourceNames) To UBound(sourceNames)
        If i > LBound(sourceNames) Then Let fullSource = fullSource & vbCrLf
        Let fullSource = fullSource & CStr(sourceNames(i))
    Next i
    Call Raise(errorCode, fullSource, GetMessage(errorCode))
End Sub
