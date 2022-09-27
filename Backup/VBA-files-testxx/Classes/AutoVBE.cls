VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AutoVBE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const FILE_FILLTER As String = _
        "All VB files, *.bas;*.frm;*.cls," & _
        "Module files,*.bas," & _
        "Form files, *.frm," & _
        "Class files, *.cls," & _
        "Text files, *.txt," & _
        "All files, *.*"
Private Const INDEX As Integer = 1 ' Default All VB files
Private Const IMPORT As String = "Import to VBA"
Private Const VBA_FOLDER As String = "\VBA-files-"
Private Const MODULE_FOLDER As String = "\Modules\"
Private Const CLASS_FOLDER As String = "\Classes\"
Private Const FORM_FOLDER As String = "\Forms\"
Private Const ELSE_FOLDER As String = "\Else\"
Private Const MODULE_EXTENSION As String = ".bas"
Private Const CLASS_EXTENSION As String = ".cls"
Private Const FORM_EXTENSION As String = ".frm"
Private Const ELSE_EXTENSION As String = ".txt"
Private Const AUTHOR_PROMPT As String = "AUTHOR: "
Private activeFile As Object
Private pathArray(0 To 4) As String
Private fileSystem As Object
Private components As Object
Private component As Object
Private codeModule As Object
Private userReponse As Integer
Private Enum VBAfile ' Codes for  type of VBAfile
        Module = 1
        ClassModule = 2
        Form = 3
        Document = 100
End Enum
Private Enum VBAfolder ' Codes for path array
        Main = 0 'VBA-files
        Module = 1
        Class = 2
        Form = 3
        Els = 4
End Enum
Private Enum UserSelect 'Codes for msgBox selection
        OK = 1
        Cancel = 2
        Abort = 3
        Retry = 4
        Ignore = 5
        Yes = 6
        No = 7
End Enum

'Constructor
Private Static Sub Class_Initialize()
        Select Case Application.name
                Case "Microsoft Excel"
                        Set activeFile = Application.ActiveWorkbook
                Case "Microsoft Word"
                        Set activeFile = Application.ActiveDocument
                Case "Microsoft PowerPoint"
                        Set activeFile = Application.ActivePresentation
        End Select
        Set components = activeFile.VBProject.VBComponents
        Set fileSystem = CreateObject("Scripting.FileSystemObject")
End Sub

'Destructor
Private Sub Class_Terminate()
        Set activeFile = Nothing
        Set components = Nothing
        Set fileSystem = Nothing
        Set component = Nothing
        Set codeModule = Nothing
End Sub

' Method

Private Static Sub tackleErrors()
        'Can not import VBA file
        If Err.Number = 50057 Then
                MsgBox Err.Description
        End If
        'FIle import is not VBA file
        If Err.Number = 50021 Then
                MsgBox Err.Description & _
                        vbNewLine & _
                        AUTHOR_PROMPT & _
                        "This file is not VBA file!"
        End If
        On Error GoTo 0
End Sub

Private Static Sub createFolders()
'       MkDir (path) 'old way can't check existed
        Dim path As Variant
        Let pathArray(VBAfolder.Main) = activeFile.path & _
                VBA_FOLDER & _
                fileSystem.getBaseName(activeFile.name)
        Let pathArray(VBAfolder.Module) = pathArray(VBAfolder.Main) & MODULE_FOLDER
        Let pathArray(VBAfolder.Class) = pathArray(VBAfolder.Main) & CLASS_FOLDER
        Let pathArray(VBAfolder.Form) = pathArray(VBAfolder.Main) & FORM_FOLDER
        Let pathArray(VBAfolder.Els) = pathArray(VBAfolder.Main) & ELSE_FOLDER
        'Loop through my path array and create folder
        For Each path In pathArray
                If Not fileSystem.FolderExists(path) Then
                      fileSystem.createFolder path
                End If
        Next path
End Sub

'Replace code be used for Module/ClassModule/Document
Private Static Sub replaceCode(ByVal path As String)
        'Check empty moduleCode
        Set codeModule = component.codeModule
        If codeModule.CountOfLines >= 1 Then
                codeModule.DeleteLines 1, codeModule.CountOfLines 'Delete old file content
        End If
        codeModule.AddFromFile (path) 'Add code from path
        codeModule.DeleteLines 1, 4 ' Remove 4 first lines: VERSION/BEGIN/MultiUse/END
End Sub

'Replace VBA file (Delete and add new component) be used for For.
Private Static Sub replaceVBAFile(ByVal path As String)
       ' Let component.name = fileSystem.getBaseName(path)
        components.Remove component
        components.IMPORT path
End Sub

Private Static Sub importVBAfile(ByVal path As String)
        Dim isExist As Boolean
        'Loop through all VBA files
        For Each component In components
               'Check VBA files already exist
                If component.name = fileSystem.getBaseName(path) Then
                        Let isExist = True: Exit For 'Trick do task when check is true
                Else
                        Let isExist = False
                End If
        Next component
        If isExist Then
                'Ask user for replacing esixt file
                Let userReponse = MsgBox(Prompt:= _
                                AUTHOR_PROMPT & _
                                "Do you want to replace " & _
                                fileSystem.GetFileName(path) & "?", _
                        Buttons:=vbYesNo, _
                        Title:="VBA file already exists")
                If userReponse = UserSelect.Yes Then
                        Select Case component.Type
                                Case VBAfile.Module
                                        Call replaceVBAFile(path)
                                Case VBAfile.ClassModule
                                        Call replaceVBAFile(path)
                                Case VBAfile.Form
                                        Call replaceVBAFile(path)
                                Case VBAfile.Document
                                        Call replaceCode(path) 'Replace code cause can not delete last sheet
                                Case Else
                                        Call replaceVBAFile(path)
                        End Select
                End If
        Else
                components.IMPORT path
        End If
End Sub

Public Static Sub importSelectedVBAfiles()
        On Error GoTo ErrorHandle
        Dim filePaths As Variant
        Dim path As Variant
        Let filePaths = getFilePaths
        ' Check if don't select files
        If TypeName(filePaths) = "Boolean" Then GoTo ExecuteProcedure
        'Loop through all paths
         For Each path In filePaths
                 Call importVBAfile(path)
        Next path
        MsgBox AUTHOR_PROMPT & "Successfully import selected VBA files!"
        GoTo ExecuteProcedure
ErrorHandle:
        Call tackleErrors
ExecuteProcedure:
End Sub

Public Static Sub importAllVBAfiles()
        On Error GoTo ErrorHandle
        Dim filePaths As Variant
        Dim subFolder As Object
        Dim file As Object
        'Create empty folders
        Call createFolders
        'Check folder exist for case delete manually
        If fileSystem.FolderExists(pathArray(VBAfolder.Main)) Then
                'Loop through all child folders in parent folder (VBA-files)
                For Each subFolder In fileSystem.GetFolder(pathArray(VBAfolder.Main)).Subfolders
                        'Loop thorugh all files in child folders (Module/Class/Form/Else/Custom)
                        For Each file In subFolder.Files
                             importVBAfile (file.path)
                        Next file
               Next subFolder
        End If
        MsgBox AUTHOR_PROMPT & "Successfully import all VBA files!"
        GoTo ExecuteProcedure
ErrorHandle:
        On Error GoTo 0
ExecuteProcedure:
End Sub

Public Static Sub exportAllVBAfiles()
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
                        Case VBAfile.Form
                                Let path = pathArray(VBAfolder.Form) & component.name & FORM_EXTENSION
                        Case VBAfile.Document
                                Let path = pathArray(VBAfolder.Els) & component.name & CLASS_EXTENSION
                        Case Else
                                Let path = pathArray(VBAfolder.Els) & component.name & ELSE_EXTENSION
                 End Select
                Call component.Export(path)
        Next component
        MsgBox AUTHOR_PROMPT & "Successfully export all VBA files!"
        GoTo ExecuteProcedure
ErrorHandle:
        On Error GoTo 0
ExecuteProcedure:
End Sub

' Assessor

Private Static Property Get getFilePaths() As Variant
        Let getFilePaths = Application.GetOpenFilename _
                (fileFilter:=FILE_FILLTER, _
                Title:=IMPORT, _
                filterIndex:=INDEX, _
                MultiSelect:=True)
End Property




