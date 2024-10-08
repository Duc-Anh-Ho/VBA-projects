VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FilesController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' AUTO VBA TOOLS
' Author: DANH
' Version: 1.0.0
' Update: 2022/10/02
' Check README.md for more information
Option Explicit
'Declare Variables
Private info As InfoConstants
Private system As SystemUpdate
Private Const FILE_FILLTER As String = _
    "All VB files, *.bas;*.frm;*.cls," & _
    "Module files,*.bas," & _
    "Form files, *.frm," & _
    "Class files, *.cls," & _
    "Text files, *.txt," & _
    "All files, *.*"
Private Const BROWSER_INDEX As Integer = 1 ' Default All VB files
Private Const BROWSER_TITLE As String = "Import to VBA"
Private Const VBA_FOLDER As String = "\VBA-files-"
Private Const MODULE_FOLDER As String = "\Modules\"
Private Const CLASS_FOLDER As String = "\Classes\"
Private Const FORM_FOLDER As String = "\Forms\"
Private Const ELSE_FOLDER As String = "\Else\"
Private Const MODULE_EXTENSION As String = ".bas"
Private Const CLASS_EXTENSION As String = ".cls"
Private Const FORM_EXTENSION As String = ".frm"
Private Const BINARY_FORM_EXTENSION As String = ".frx"
Private Const TEXT_EXTENSION As String = ".txt"
Private activeFile As Object
Private pathArray(0 To 4) As String
Private fileSystem As Object
Private components As Object
Private component As Object
Private codeModule As Object
Private Enum VBAfile ' Codes for  type of VBAfile
    Module = 1 'vbext_ct_StdModule
    ClassModule = 2 'vbext_ct_ClassModule
    form = 3 'vbext_ct_MSForm
    ActiveX = 11 'vbext_ct_ActiveXDesigner
    Document = 100 'vbext_ct_Document
End Enum
Private Enum VBAfolder ' Codes for path array
    MAIN = 0 'VBA-files
    Module = 1
    Class = 2
    form = 3
    Els = 4
End Enum
Private Enum UserReponse 'Codes for msgBox selection
    OK = 1
    Cancel = 2
    Abort = 3
    Retry = 4
    Ignore = 5
    Yes = 6
    No = 7
End Enum
'Initialize Variables
Private Function hasVariables() As Boolean
On Error GoTo ErrorHandle
    Set info = New InfoConstants
    Set system = New SystemUpdate
    Set fileSystem = system.createFileSystem
    If Not system.hasWorkPlace(hasMsg:=True, workPlaceType:="xlWorksheet") Then
        Let hasVariables = False
        GoTo ExecuteProcedure
    End If
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
    Select Case Application.name
        Case "Microsoft Excel"
            Set activeFile = Application.ActiveWorkbook
        Case "Microsoft Word"
            Set activeFile = Application.ActiveDocument
        Case "Microsoft PowerPoint"
            Set activeFile = Application.ActivePresentation
    End Select
    Set components = activeFile.VBProject.VBComponents
ErrorHandle:
    Call system.tackleErrors
ExecuteProcedure:
End Sub

'Destructor
Private Sub Class_Terminate()
    Set activeFile = Nothing
    Set components = Nothing
    Set fileSystem = Nothing
    Set component = Nothing
    Set codeModule = Nothing
End Sub

'ASSESSORS

Private Property Get getFilePaths() As Variant
    Let getFilePaths = Application.GetOpenFilename( _
        FileFilter:=FILE_FILLTER, _
        Title:=BROWSER_TITLE, _
        FilterIndex:=BROWSER_INDEX, _
        MultiSelect:=True)
End Property

Private Property Get getUserReponse(Optional ByVal defaultReponse As Integer) As Integer
    If defaultReponse = 0 Then
        Let getUserReponse = MsgBox( _
            Prompt:= _
                info.getPrompt & _
                "Do you want to replace " & _
                (component.name) & _
                " VBA file?", _
            Buttons:=vbYesNo, _
            Title:="VBA file already exists") ' TODO : Refactor code
    Else
        Let getUserReponse = defaultReponse
    End If
End Property

Private Property Get isValidVBAFileExtension(ByVal path As String) As Boolean
        Select Case "." & fileSystem.GetExtensionName(path) 'Add dot cause this method get extension only
            Case MODULE_EXTENSION
                Let isValidVBAFileExtension = True
            Case CLASS_EXTENSION
                Let isValidVBAFileExtension = True
            Case FORM_EXTENSION
                Let isValidVBAFileExtension = True
            Case TEXT_EXTENSION
                Let isValidVBAFileExtension = True
            Case BINARY_FORM_EXTENSION
                Let isValidVBAFileExtension = False
            Case Else
                MsgBox _
            Prompt:=info.getPrompt & _
                "Can not import file: " & _
                 fileSystem.getFileName(path) & _
                 vbNewLine & _
                 "Please check again!", _
            Buttons:=vbOKOnly + vbExclamation, _
            Title:=info.getAuthor
                Let isValidVBAFileExtension = False
        End Select
End Property
'METHODS

Private Sub createFolders()
'       MkDir (path) 'old way can't check existed
    Dim path As Variant
    Let pathArray(VBAfolder.MAIN) = activeFile.path & _
        VBA_FOLDER & _
        fileSystem.GetBaseName(activeFile.name)
    Let pathArray(VBAfolder.Module) = pathArray(VBAfolder.MAIN) & MODULE_FOLDER
    Let pathArray(VBAfolder.Class) = pathArray(VBAfolder.MAIN) & CLASS_FOLDER
    Let pathArray(VBAfolder.form) = pathArray(VBAfolder.MAIN) & FORM_FOLDER
    Let pathArray(VBAfolder.Els) = pathArray(VBAfolder.MAIN) & ELSE_FOLDER
    'Loop through my path array and create folder
    For Each path In pathArray
        If Not fileSystem.FolderExists(path) Then
            fileSystem.createFolder path
        End If
    Next path
End Sub

'Replace code (Delete and add new component) be used for Document.
Private Sub replaceCode(ByVal path As String)
    'Check empty moduleCode
    Set codeModule = component.codeModule
    If codeModule.CountOfLines >= 1 Then
        codeModule.DeleteLines 1, codeModule.CountOfLines 'Delete old file content
    End If
    codeModule.AddFromFile (path) 'Add code from path
    codeModule.DeleteLines 1, 4 ' Remove 4 first lines: VERSION/BEGIN/MultiUse/END
End Sub

'Replace VBA file be used for Module/ClassModule/Document
Private Sub replaceVBAFile(ByVal path As String)
    Dim amountVBAfile As Integer
    Let amountVBAfile = components.Count
    'Remove VBA file
    components.remove component
    'Check case can't remove immediately (in use VBA file)
    If amountVBAfile = components.Count Then
        component.name = component.name & info.getAuthor
    End If
    'Add VBA file with path
    If isValidVBAFileExtension(path) Then
        components.Import path
    End If
End Sub

Private Sub importVBAFile(ByVal path As String, Optional ByVal defaultReponse As Integer)
    Dim isExist As Boolean
    'Loop through all VBA files
    For Each component In components
            'Check VBA files already exist
        If component.name = fileSystem.GetBaseName(path) Then
            Let isExist = True: Exit For 'Trick do task when check is true
        End If
    Next component
    If isExist Then
        'Ask user for replacing esixt file
        If getUserReponse(defaultReponse) = UserReponse.Yes Then
            Select Case component.Type
                Case VBAfile.Module
                    Call replaceVBAFile(path)
                Case VBAfile.ClassModule
                    Call replaceVBAFile(path)
                Case VBAfile.form
                    Call replaceVBAFile(path)
'                Case VBAfile.ActiveX ' TO-DO: Research what is this file extension
'                    Call replaceVBAFile(path)
                Case VBAfile.Document
                    Call replaceCode(path) 'Replace code cause can not delete last sheet
                Case Else
                    Call replaceVBAFile(path)
            End Select
            Let isExist = False ' Reset IsExsit
        End If
    Else
        If isValidVBAFileExtension(path) Then
            components.Import path
        End If
    End If
End Sub
'MAIN

Public Sub importSelectedVBAfiles()
On Error GoTo ErrorHandle
    Dim filePaths As Variant
    Dim path As Variant
    Let filePaths = getFilePaths
    ' Check if don't select files
    If typeName(filePaths) = "Boolean" Then GoTo ExecuteProcedure
    'Loop through all paths
    For Each path In filePaths
        Call importVBAFile(path)
    Next path
    MsgBox info.getPrompt & "Successfully import selected VBA files!"
    GoTo ExecuteProcedure
ErrorHandle:
    Call system.tackleErrors
ExecuteProcedure:
End Sub

Public Sub importAllVbaFiles()
On Error GoTo ErrorHandle
    Dim filePaths As Variant
    Dim subFolder As Object
    Dim file As Object
    'Create empty folders
    Call createFolders
    'Check folder exist for case delete manually
    If fileSystem.FolderExists(pathArray(VBAfolder.MAIN)) Then
        'Loop through all child folders in parent folder (VBA-files)
        For Each subFolder In fileSystem.getFolder(pathArray(VBAfolder.MAIN)).Subfolders
            'Loop thorugh all files in child folders (Module/Class/Form/Else/Custom)
            For Each file In subFolder.Files
                Call importVBAFile(file.path, defaultReponse:=UserReponse.Yes)
            Next file
        Next subFolder
    End If
    MsgBox info.getPrompt & "Successfully import all VBA files!"
    GoTo ExecuteProcedure
ErrorHandle:
    Call system.tackleErrors
ExecuteProcedure:
End Sub

Public Sub exportAllVbaFiles()
On Error GoTo ErrorHandle
    Dim path As String
    'Create folder
    Call createFolders
    'Loop through all VBA files
    For Each component In components
        Select Case component.Type
            Case VBAfile.Module
                Let path = pathArray(VBAfolder.Module) & component.name & MODULE_EXTENSION
            Case VBAfile.ClassModule
                Let path = pathArray(VBAfolder.Class) & component.name & CLASS_EXTENSION
            Case VBAfile.form
                Let path = pathArray(VBAfolder.form) & component.name & FORM_EXTENSION
            Case VBAfile.Document
                Let path = pathArray(VBAfolder.Els) & component.name & CLASS_EXTENSION
            Case Else
                Let path = pathArray(VBAfolder.Els) & component.name & TEXT_EXTENSION
        End Select
        Call component.Export(path)
    Next component
    MsgBox info.getPrompt & "Successfully export all VBA files!"
    GoTo ExecuteProcedure
ErrorHandle:
    Call system.tackleErrors
ExecuteProcedure:
End Sub

'File Name to clipboard
Public Sub copyFileName(Optional ByRef typeName As String = "name")
On Error GoTo ErrorHandle
    Dim name As String
    Let name = system.getFileName(typeName)
    If name = vbNullString Then
         MsgBox _
                Prompt:=info.getPrompt & _
                    "This file doesn't have " & typeName & _
                    vbNewLine & _
                    "Please save it first!!", _
                Buttons:=vbOKOnly + vbExclamation, _
                Title:=info.getAuthor
        GoTo ExecuteProcedure
    End If
    Let system.setClipboard = system.getFileName(typeName)
ErrorHandle:
    Call system.tackleErrors
ExecuteProcedure:
End Sub





