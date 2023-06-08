VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AutoAddin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Check README.md for more information

Option Explicit
'Declare Variables
Private info As InfoConstants
Private system As SystemUpdate
Private fileSystem As Object
Private userResponse As VbMsgBoxResult
Private addinPath As String
Private addinShortName As String
Private addinName As String
Private addinFullName As String
Private IsAddinWb As Boolean
Private isAddinExisted As Boolean

'Initialize Variables
Private Function hasVariables() As Boolean
On Error GoTo ErrorHandle
    Set info = New InfoConstants
    Set system = New SystemUpdate
    Set fileSystem = system.createFileSystem
    'NOTE: Don't need to have workplace for this procedure
    'UserLibraryPath of PC:
    '"C:\Users\Admin\AppData\Roaming\Microsoft\AddIns\"
    Let addinPath = Application.UserLibraryPath
    Let addinShortName = info.getAddinShortName
    Let addinName = info.getAddinName
    Let addinFullName = addinPath & addinName
    Let hasVariables = True
    GoTo ExecuteProcedure
ErrorHandle:
    Call system.tackleErrors
ExecuteProcedure:
End Function

'Constructor
Private Sub Class_Initialize()
End Sub

'Destructor
Private Sub Class_Terminate()
    Set info = Nothing
    Set system = Nothing
    Set fileSystem = Nothing
End Sub

'METHODS
Private Function checkIsAddinWb(ByRef currentWb As Workbook)
    If currentWb.FullName = addinFullName Then
        Let IsAddinWb = True
    ElseIf currentWb.FullName <> addinFullName Then
        Let IsAddinWb = False
    End If
End Function

Private Function checkisAddinExisted(ByRef addinFullName As String)
    If Len(Dir(addinFullName)) = 0 Then
        isAddinExisted = False
    ElseIf Len(Dir(addinFullName)) <> 0 Then
        isAddinExisted = True
    End If
End Function

Private Sub saveFileAsAddin(ByRef currentWb As Workbook)
    currentWb.SaveAs _
        fileName:=addinFullName, _
        FileFormat:=xlOpenXMLAddIn ' xlOpenXMLAddIn = .xlam/55
End Sub

Private Sub enableAddin(ByRef isEnable As Boolean)
    Dim ai As addin
    For Each ai In Application.AddIns
        If ai.FullName = addinFullName Then
            If AddIns("danh-tools").Installed <> isEnable Then
                Let AddIns(addinShortName).Installed = isEnable
            End If
        End If
    Next ai
End Sub

Private Sub changeWBAcessToReadOnly(ByRef wb As Workbook)
    If Not wb.ReadOnly Then wb.ChangeFileAccess MODE:=xlReadOnly
End Sub

Private Sub installAddin()
    'Do Nothing when this workbook is addin
    If IsAddinWb Then Exit Sub
    'Ask user update when addin already installed
    If isAddinExisted Then
         Let userResponse = MsgBox( _
            Prompt:="Add-in: " & addinShortName & " allready exists!" & _
                vbNewLine & _
                "Do you want to update to version " & _
                info.getVersion & " ?", _
            Buttons:=vbYesNo + vbQuestion, _
            Title:=info.getAuthor)
    Else
        Let userResponse = vbYes
    End If
    'Do Nothing when user don't want to update
    If Not userResponse = vbYes Then
        'Close wb (installation) if click no
        ThisWorkbook.Close SaveChanges:=False
        Exit Sub
    End If
    'Disable addin when addin already installed
    If isAddinExisted Then
        Call enableAddin(False)
        Call deleteAddin(False)
    End If
    'Save installation wb to UserLibraryPath of PC and enable it
    Call saveFileAsAddin(ThisWorkbook)
    Call enableAddin(True)
    'Annouce install successfully:
    ' up version or first time
    If isAddinExisted Then
        MsgBox _
            Prompt:=info.getPrompt & _
            "Add-in: " & addinShortName & " successfully updated to version -" & info.getVersion & "-!", _
            Buttons:=vbOKOnly + vbInformation, _
            Title:=info.getAuthor
    Else
        MsgBox _
            Prompt:=info.getPrompt & _
            "Add-in: " & addinShortName & "is successfully installed version -" & info.getVersion & "-!", _
            Buttons:=vbOKOnly + vbInformation, _
            Title:=info.getAuthor
    End If
    'Close wb (installation) after install
    ThisWorkbook.Close SaveChanges:=False
End Sub

Private Sub deleteAddin(Optional ByRef hasConfirm As Boolean = False)
    'Announce and exit when addin don't exist anymore
    'For case user delete manually addin in PC path - UserLibraryPath
    If Not isAddinExisted Then
        MsgBox _
            Prompt:=info.getPrompt & _
            "Add-in: " & addinShortName & " does not exist!", _
            Buttons:=vbOKOnly + vbExclamation, _
            Title:=info.getAuthor
        Exit Sub
    End If
    'Optional - ask user want to delete
    If hasConfirm Then
         Let userResponse = MsgBox( _
            Prompt:=info.getPrompt & _
                "Do you want to delete " & _
                "Add-in: " & addinShortName & _
                " from you computer ?", _
            Buttons:=vbYesNo + vbQuestion, _
            Title:=info.getAuthor)
        'Do Nothing when user don't want remove addin
        If Not userResponse = vbYes Then Exit Sub
    End If
    'Change current wb to read-only when it is installation
    If IsAddinWb Then Call changeWBAcessToReadOnly(ThisWorkbook)
    'Delete addin
    Kill (addinFullName)
    'Optional - announce after delete successfully
    If hasConfirm Then
         MsgBox _
            Prompt:=info.getPrompt & _
                "Add-in: " & addinShortName & _
                " sucessfully removed!", _
            Buttons:=vbOKOnly + vbInformation, _
            Title:=info.getAuthor
    End If
    'Disable addin after delete
    Call enableAddin(False)
End Sub

'MAIN
Public Sub install()
On Error GoTo ErrorHandle
    If Not hasVariables Then GoTo ExecuteProcedure
    Call system.speedOn(, True)
    'Check current file is in addin
    Call checkIsAddinWb(ThisWorkbook)
    'Check first time install
    Call checkisAddinExisted(addinFullName)
    'InstallAddin
    Call installAddin
GoTo ExecuteProcedure
ErrorHandle:
    Call system.tackleErrors
ExecuteProcedure:
    Call system.speedOff
End Sub

Public Sub remove(Optional ByRef hasConfirm As Boolean = False)
On Error GoTo ErrorHandle
    If Not hasVariables Then GoTo ExecuteProcedure
    Call system.speedOn(, True)
    'Check current file is in addin
    Call checkIsAddinWb(ThisWorkbook)
    'Check first time install
    Call checkisAddinExisted(addinFullName)
    'DeleteAddin
    Call deleteAddin(hasConfirm)
GoTo ExecuteProcedure
ErrorHandle:
    Call system.tackleErrors
ExecuteProcedure:
    Call system.speedOff
End Sub

