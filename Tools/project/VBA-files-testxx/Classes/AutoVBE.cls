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
Private Const IMPORT As String = "Import to VBE"
Private Const VBE_FOLDER As String = "\VBA-files-"
Private Const MODULES_FOLDER As String = "\Modules\"
Private Const CLASS_FOLDER As String = "\Classes\"
Private Const FORM_FOLDER As String = "\Forms\"
Private Const ELSE_FOLDER As String = "\Else\"
Private activeFile As Object
Private pathArray(0 To 4) As String
Private fileSystem As Object
Private components As Object
Private component As Object
Private Enum VBEfile ' Codes for  type of vbefile
        Module = 1
        ClassModule = 2
        Form = 3
        Document = 100
End Enum
Private Enum VBEfolder ' Codes for path array
        Main = 0
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
End Sub


' Method

Public Static Sub createFolder()
'       MkDir (path) 'old way can't check existed
        Dim path As Variant
        Let pathArray(VBEfolder.Main) = activeFile.path & _
                VBE_FOLDER & _
                fileSystem.getBaseName(activeFile.name)
        Let pathArray(VBEfolder.Module) = pathArray(VBEfolder.Main) & MODULES_FOLDER
        Let pathArray(VBEfolder.Class) = pathArray(VBEfolder.Main) & CLASS_FOLDER
        Let pathArray(VBEfolder.Form) = pathArray(VBEfolder.Main) & FORM_FOLDER
        Let pathArray(VBEfolder.Els) = pathArray(VBEfolder.Main) & ELSE_FOLDER
        For Each path In pathArray
                If Not fileSystem.FolderExists(path) Then
                      fileSystem.createFolder path
                End If
        Next path
End Sub

Public Static Sub replaceCode(ByVal path As String)
        Dim codeModule As Object
        Set codeModule = components.item(fileSystem.getBaseName(path)).codeModule
        If codeModule.CountOfLines > 1 Then codeModule.DeleteLines 1, codeModule.CountOfLines 'Delete old file content
        codeModule.AddFromFile (path) 'Add code from path
        codeModule.DeleteLines 1, 4 ' Remove 4 first lines: VERSION/BEGIN/MultiUse/END
End Sub

Public Static Sub importVBEfiles()
        On Error GoTo ErrorHandle
        Dim filePaths As Variant
        Dim path As Variant
        Dim userReponse As Integer
        Dim isExist As Boolean
        Let filePaths = getFilePaths
        If TypeName(filePaths) = "Boolean" Then GoTo ExecuteProcedure ' Check if don't select files
        'Loop through all paths
         For Each path In filePaths
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
                        Let userReponse = MsgBox(Prompt:="Do you want to replace " & _
                                        fileSystem.GetFileName(path) & "?", _
                                Buttons:=vbYesNo, _
                                Title:="VBA file already exists")
                        If userReponse = UserSelect.Yes Then
                                Call replaceCode(path)
                        End If
                Else
                        Application.VBE.ActiveVBProject.components.IMPORT path
                End If
        Next path
        GoTo ExecuteProcedure
ErrorHandle:
        If Err.Number = 50057 Then
                MsgBox Err.Description
                On Error GoTo 0
        End If
ExecuteProcedure:
End Sub

Public Static Sub exportAllVBEfiles()
        On Error GoTo ErrorHandle
        Dim path As String
        'Create folder
                Call createFolder
        'Loop through all VBA files
        For Each component In components
                Select Case component.Type
                        Case VBEfile.Module
                                Let path = pathArray(VBEfolder.Module) & component.name & ".bas"
                        Case VBEfile.ClassModule
                                Let path = pathArray(VBEfolder.Class) & component.name & ".cls"
                        Case VBEfile.Form
                                Let path = pathArray(VBEfolder.Form) & component.name & ".frm"
                        Case VBEfile.Document
                                Let path = pathArray(VBEfolder.Els) & component.name & ".cls"
                        Case Else
                                Let path = pathArray(VBEfolder.Els) & component.name & ".txt"
                 End Select
                Call component.Export(path)
        Next component
        GoTo ExecuteProcedure
ErrorHandle:
        On Error GoTo 0
ExecuteProcedure:
End Sub

Public Static Sub importAllVBEfiles()
        On Error GoTo ErrorHandle
'        If fileSystem.FolderExists(path) Then
'               fileSystem.GetFolder
'        End If
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

