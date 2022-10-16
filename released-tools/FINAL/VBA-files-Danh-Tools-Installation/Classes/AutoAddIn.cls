VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AutoAddIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' AUTO ADD ADDIN TOOLS
' Author: DANH
' Version: 2.0.2
' Update: 2022/10/09
' Check README.md for more information

Option Explicit

Private Const VERSION = "v2.0.2"
Private Const ADDIN_NAME = "Danh-Tools" ' ' <-- ***DON'T ADD _AddIn IN THIS NAME, BECAUSE OF UNKNOW ERROR EX. Danh_Tools_Addin
Private Const AUTHOR_PROMPT As String = "AUTHOR: "
Private Const AUTHOR_NAME = "DANH"
Private Const ADDIN_EXTENSION As String = ".xlam"
Private WithEvents AddInApplication As Excel.Application
Attribute AddInApplication.VB_VarHelpID = -1
Private WithEvents installFile As Excel.Workbook 'TODO - Switch case for Ppt and word
Attribute installFile.VB_VarHelpID = -1
Private fileSystem As Object
Private addinFile As Excel.Workbook 'TODO - Switch case for Ppt and word
Private addInFilePath As String
Private addInFileName As String
Private addinFileFullName As String
Private userResponse As VbMsgBoxResult

'Constructor
Private Sub Class_Initialize()
'On Error GoTo ErrorHandle
     Call initializeVariables
     Call speedOn
    'Check opened file is addIn (.xlam)
    If isAddInFile(installFile.FullName) Then
    '
    End If
    'Check opened file is installation (not .xlam)
    If Not isAddInFile(installFile.FullName) Then
        Call installAddIn
        Call enableAddIn
        'Close Excel
        Call speedOff
        installFile.Close savechanges:=False
    End If
GoTo ExecuteProcedure
ErrorHandle:
    Call tackleErrors
ExecuteProcedure:
End Sub

'Destructor
Private Sub Class_Terminate()
'On Error GoTo ErrorHandle
    'Check opened file is addIn (.xlam)
    If isAddInFile(installFile.FullName) Then
    '
    End If
    'Check opened file is installation (not .xlam)
    If Not isAddInFile(installFile.FullName) Then
    '
    End If
    Call speedOff
GoTo ExecuteProcedure
ErrorHandle:
    Call tackleErrors
ExecuteProcedure:
End Sub

'Assessors

Private Static Property Get isAddInFile(ByVal fileFullName As String) As Boolean
    If fileFullName = addinFileFullName Then
        isAddInFile = True
    ElseIf fileFullName <> addinFileFullName Then
        isAddInFile = False
    End If
End Property

Private Static Property Get isExistFile(ByVal fileFullName As String) As Boolean
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

Private Static Sub initializeVariables()
    Set AddInApplication = Application
    Select Case Application.Name
        Case "Microsoft Excel"
            Set installFile = AddInApplication.ThisWorkbook
        Case "Microsoft Word"
            Set installFile = AddInApplication.ThisDocument
        Case "Microsoft PowerPoint"
            Set installFile = AddInApplication.ThisPresentation
    End Select
'   installFile.Author = AUTHOR_NAME
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    Let addInFilePath = Application.UserLibraryPath 'Place stored Addin - This PC: "C:\Users\Admin\AppData\Roaming\Microsoft\AddIns\"
    Let addInFileName = ADDIN_NAME & ADDIN_EXTENSION
    Let addinFileFullName = addInFilePath & addInFileName
End Sub

' SAVEAS (or FILECOPY) --> ADD --> ENABLE
Private Static Sub saveAddInFile()
    installFile.SaveAs _
        fileName:=addinFileFullName, _
        FileFormat:=xlOpenXMLAddIn
        ' xlOpenXMLAddIn = xlam/55
End Sub

Private Static Sub copyAddInFile()
    Call fileSystem.CopyFile( _
        Source:=installFile.FullName, _
        Destination:=addinFileFullName)
    ' NOTE: Can't use because CopyFile just change extension but don't change to addin type file
End Sub

Private Static Sub addAddin()
    Dim ai As AddIn
    For Each ai In Application.AddIns
        If ai.FullName = addinFileFullName Then
            On Error Resume Next ' TODO: Resolve this
            Application.AddIns.Add fileName:=addinFileFullName
        End If
    Next ai
    Let AddIns(ADDIN_NAME).Installed = True
End Sub

Private Static Sub enableAddIn()
    Dim ai As AddIn
    For Each ai In Application.AddIns
        If ai.FullName = addinFileFullName Then
            Let Application.AddIns(ADDIN_NAME).Installed = True
        End If
    Next ai
End Sub

' DISABLE --> CLOSE (or REMOVE) --> DELETE
Private Static Sub disableAddIn()
    Dim ai As AddIn
    For Each ai In Application.AddIns
        If ai.FullName = addinFileFullName Then
            Let Application.AddIns(ADDIN_NAME).Installed = False
        End If
    Next ai
End Sub

Private Static Sub removeAddin()
'    Application.AddIns.Remove fileName:=addinFileFullName
     ' NOTE: Excel >=2007 Deleted this method.
End Sub

Private Static Sub closeAddin()
    Dim wb As Workbook
    'Loop through all workbook
    For Each wb In Workbooks
        If wb.Name = addInFileName Then
            Application.Workbooks(ADDIN_NAME).Close savechanges:=False
        End If
    Next wb
End Sub

Public Static Sub deleteAddInFile()
    'Check can't find Addin find.
    If Len(Dir(addinFileFullName)) = 0 Then
        MsgBox _
            Prompt:="Add-in: " & ADDIN_NAME & " does not exist!", _
            Buttons:=vbOKOnly + vbExclamation, _
            Title:=AUTHOR_NAME
    End If
    'Check Addin exist
    If Len(Dir(addinFileFullName)) <> 0 Then
        Call closeAddin
        Call disableAddIn
        Kill (addinFileFullName)
        Call closeAddin
    End If
End Sub

Private Static Sub installAddIn()
   'Already exist in folder
    If isExistFile(addinFileFullName) Then
        'Ask user update
         userResponse = MsgBox( _
            Prompt:="Add-in: " & ADDIN_NAME & " allready exists!" & _
                vbNewLine & _
                "Do you want to update to version " & _
                VERSION & " ?", _
            Buttons:=vbYesNo + vbQuestion, _
            Title:=AUTHOR_NAME)
        'User Chose yes
        If userResponse = vbYes Then
            Call disableAddIn
            Call deleteAddInFile
            Call saveAddInFile
            'Call addAddin
            Call enableAddIn
            MsgBox _
                Prompt:="Add-in: " & ADDIN_NAME & " successfully updated to version -" & VERSION & "-!", _
                Buttons:=vbOKOnly + vbInformation, _
                Title:=AUTHOR_NAME
        End If
    End If
    'Check first time install
    If Not isExistFile(addinFileFullName) Then
        Call saveAddInFile
        'Call addAddin
        Call enableAddIn
        MsgBox Prompt:="New Add-in: " & ADDIN_NAME & " successfully Installed!", _
           Buttons:=vbOKOnly + vbInformation, _
           Title:=AUTHOR_NAME
    End If
End Sub

'Errors
Private Static Sub tackleErrors()
    Select Case Err.Number
        Case 0
        'Can't enable Addin
        Case 1004
            Resume Next
        'VBA file have password
        Case 50289
            MsgBox _
                Prompt:=Err.Description & _
                    vbNewLine & _
                    AUTHOR_PROMPT & _
                    "Can not access this VBA file because of been protected by a password!", _
                Buttons:=vbCritical + vbOKOnly, _
                Title:=AUTHOR_NAME
        'Un-handled Error
        Case Else
            Call errorDisplay
    End Select
    On Error GoTo 0
End Sub

Private Static Sub errorDisplay()
    Dim errorMessage As String: Let errorMessage = _
        "Error # " & Str(Err.Number) & _
        " was generated by " & Err.Source & _
        Chr(13) & "Error Line: " & Erl & _
        Chr(13) & Err.Description
    MsgBox _
        Prompt:=errorMessage, _
        Title:="Error", _
        HelpFile:=Err.HelpFile, _
        Context:=Err.HelpContext
End Sub

' PerCls tool rework
Private Static Sub speedOn()
    With Application
        .ScreenUpdating = False
        .AskToUpdateLinks = False
        .DisplayAlerts = False
    End With
End Sub

Private Static Sub speedOff()
    With Application
        .ScreenUpdating = True
        .AskToUpdateLinks = True
        .DisplayAlerts = True
    End With
End Sub


