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
Private info As Object
Private system As Object
Private userResponse As VbMsgBoxResult
Private Const ADDIN_NAME = "Danh-Tools" ' ' <-- ***DON'T ADD _AddIn IN THIS NAME, BECAUSE OF UNKNOW ERROR EX. Danh_Tools_Addin
Private Const ADDIN_EXTENSION As String = ".xlam"
Private AddInApp As Excel.Application
Attribute AddInApp.VB_VarHelpID = -1
Private installFile As Excel.Workbook 'TODO - Switch case for Ppt and word
Attribute installFile.VB_VarHelpID = -1
Private fileSystem As Object
Private addinFile As Excel.Workbook 'TODO - Switch case for Ppt and word
Private addInFilePath As String
Private addInFileName As String
Private addinFileFullName As String
'Initialize Variables
Private Function hasVariables() As Boolean
On Error GoTo ErrorHandle
    Set info = New MyInfo
    Set system = New SystemUpdate
    Set fileSystem = system.createFileSystem
    Set AddInApp = Application
    Select Case Application.Name
        Case "Microsoft Excel"
            Set installFile = AddInApp.ThisWorkbook
        Case "Microsoft Word"
            Set installFile = AddInApp.ThisDocument
        Case "Microsoft PowerPoint"
            Set installFile = AddInApp.ThisPresentation
    End Select
    Let addInFilePath = Application.UserLibraryPath 'Place stored Addin - This PC: "C:\Users\Admin\AppData\Roaming\Microsoft\AddIns\"
    Let addInFileName = ADDIN_NAME & ADDIN_EXTENSION
    Let addinFileFullName = addInFilePath & addInFileName
    hasVariables = True
    GoTo ExecuteProcedure
ErrorHandle:
    Call system.tackleErrors
ExecuteProcedure:
End Function

'Constructor
Private Sub Class_Initialize()
On Error GoTo ErrorHandle
    If Not hasVariables Then GoTo ExecuteProcedure
     Call system.speedOn(, True)
    'Check opened file is addIn (.xlam)
    If isAddInFile(installFile.FullName) Then
    '
    End If
    'Check opened file is installation (not .xlam)
    If Not isAddInFile(installFile.FullName) Then
        Call installAddIn
        Call enableAddIn
        Call system.speedOff
Stop ' TESTING ***
        installFile.Close savechanges:=False
    End If
GoTo ExecuteProcedure
ErrorHandle:
    Call system.tackleErrors
ExecuteProcedure:
End Sub

'Destructor
Private Sub Class_Terminate()
On Error GoTo ErrorHandle
    'Check opened file is addIn (.xlam)
    If isAddInFile(installFile.FullName) Then
    '
    End If
    'Check opened file is installation (not .xlam)
    If Not isAddInFile(installFile.FullName) Then
    '
    End If
    Call system.speedOff
GoTo ExecuteProcedure
ErrorHandle:
    Call system.tackleErrors
ExecuteProcedure:
End Sub

'Assessors

Private Property Get isAddInFile(ByVal fileFullName As String) As Boolean
    If fileFullName = addinFileFullName Then
        isAddInFile = True
    ElseIf fileFullName <> addinFileFullName Then
        isAddInFile = False
    End If
End Property

Private Property Get isExistFile(ByVal fileFullName As String) As Boolean
    If Len(Dir(fileFullName)) = 0 Then
        isExistFile = False
    ElseIf Len(Dir(fileFullName)) <> 0 Then
        isExistFile = True
    End If
End Property

'Mutators

'Events Define
'
'Private Sub installFile_BeforeClose(Cancel As Boolean)
''
'End Sub
'
'Private Sub installFile_AfterSave(ByVal Success As Boolean)
''
'End Sub

'Methods

' SAVEAS (or FILECOPY) --> ADD --> ENABLE
Private Sub saveAddInFile()
    installFile.SaveAs _
        Filename:=addinFileFullName, _
        FileFormat:=xlOpenXMLAddIn ' xlOpenXMLAddIn = .xlam/55
End Sub

Private Sub copyAddInFile()
    Call fileSystem.CopyFile( _
        Source:=installFile.FullName, _
        Destination:=addinFileFullName)
    ' NOTE: Can't use because CopyFile just change extension but don't change to addin type file
End Sub

Private Sub addAddin()
    Dim ai As AddIn
    For Each ai In Application.AddIns
        If ai.FullName = addinFileFullName Then
            On Error Resume Next ' TODO: Resolve this
            Application.AddIns.add Filename:=addinFileFullName
        End If
    Next ai
    Let AddIns(ADDIN_NAME).Installed = True
End Sub

Private Sub enableAddIn()
    Dim ai As AddIn
    For Each ai In Application.AddIns
        If ai.FullName = addinFileFullName Then
            Let Application.AddIns(ADDIN_NAME).Installed = True
        End If
    Next ai
End Sub

' DISABLE --> CLOSE (or REMOVE) --> DELETE
Private Sub disableAddIn()
    Dim ai As AddIn
    For Each ai In Application.AddIns
        If ai.FullName = addinFileFullName Then
            Let Application.AddIns(ADDIN_NAME).Installed = False
        End If
    Next ai
End Sub

Private Sub removeAddin()
'    Application.AddIns.Remove fileName:=addinFileFullName
     ' NOTE: Excel >=2007 Deleted this method.
End Sub

Private Sub closeAddin()
    Dim Wb As Workbook
    'Loop through all workbook
    For Each Wb In Workbooks
        If Wb.Name = addInFileName Then
            Application.Workbooks(ADDIN_NAME).Close savechanges:=False
        End If
    Next Wb
End Sub

Public Sub deleteAddInFile()
    'Check can't find Addin find.
    If Len(Dir(addinFileFullName)) = 0 Then
        MsgBox _
            Prompt:="Add-in: " & ADDIN_NAME & " does not exist!", _
            Buttons:=vbOKOnly + vbExclamation, _
            Title:=info.getAuthor
    End If
    'Check Addin exist
    If Len(Dir(addinFileFullName)) <> 0 Then
        Call closeAddin
        Call disableAddIn
        Kill (addinFileFullName)
        Call closeAddin
    End If
End Sub

Private Sub installAddIn()
   'Already exist in folder
    If isExistFile(addinFileFullName) Then
        'Ask user update
         userResponse = MsgBox( _
            Prompt:="Add-in: " & ADDIN_NAME & " allready exists!" & _
                vbNewLine & _
                "Do you want to update to version " & _
                info.GetVersion & " ?", _
            Buttons:=vbYesNo + vbQuestion, _
            Title:=info.getAuthor)
        'User Chose yes
        If userResponse = vbYes Then
            Call disableAddIn
            Call deleteAddInFile
            Call saveAddInFile
            'Call addAddin
            Call enableAddIn
            MsgBox _
                Prompt:="Add-in: " & ADDIN_NAME & " successfully updated to version -" & info.GetVersion & "-!", _
                Buttons:=vbOKOnly + vbInformation, _
                Title:=info.getAuthor
        End If
    End If
    'Check first time install
    If Not isExistFile(addinFileFullName) Then
        Call saveAddInFile
        'Call addAddin
        Call enableAddIn
        MsgBox Prompt:="New Add-in: " & ADDIN_NAME & " successfully Installed!", _
           Buttons:=vbOKOnly + vbInformation, _
           Title:=info.getAuthor
    End If
End Sub
