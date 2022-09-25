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

Private Const FILE_FILLTER As String = "All VB files, *.bas;*.frm;*.cls," & _
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
Private pathArray(0 To 4) As String
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

' Method

Public Static Sub createFolder()
'       MkDir (path) 'old way can't check existed
        Dim path As Variant
        With CreateObject("Scripting.FileSystemObject") 'Without add script runtime
                Let pathArray(VBEfolder.Main) = ActiveWorkbook.path & _
                        VBE_FOLDER & _
                        .GetBaseName(ActiveWorkbook.name)
                Let pathArray(VBEfolder.Module) = pathArray(VBEfolder.Main) & MODULES_FOLDER
                Let pathArray(VBEfolder.Class) = pathArray(VBEfolder.Main) & CLASS_FOLDER
                Let pathArray(VBEfolder.Form) = pathArray(VBEfolder.Main) & FORM_FOLDER
                Let pathArray(VBEfolder.Els) = pathArray(VBEfolder.Main) & ELSE_FOLDER
                For Each path In pathArray
                        If Not .FolderExists(path) Then
                               .createFolder path
                        End If
                Next path
        End With
End Sub

Public Static Sub importVBEfiles()
        On Error GoTo ErrorHandle
        Dim filePaths As Variant
        Dim path As Variant
        Let filePaths = getFilePaths
        If TypeName(filePaths) = "Boolean" Then GoTo ExecuteProcedure ' Check if don't select files
        'Loop through all paths
        For Each path In filePaths
                 Application.VBE.ActiveVBProject.VBComponents.IMPORT path
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
        Dim VBComponent As Object
        Dim path As String
        'Create folder
                Call createFolder
        'Loop through all VBA files
        For Each VBComponent In ActiveWorkbook.VBProject.VBComponents
                Select Case VBComponent.Type
                        Case VBEfile.Module
                                Let path = pathArray(VBEfolder.Module) & VBComponent.name & ".bas"
                        Case VBEfile.ClassModule
                                Let path = pathArray(VBEfolder.Class) & VBComponent.name & ".cls"
                        Case VBEfile.Form
                                Let path = pathArray(VBEfolder.Form) & VBComponent.name & ".frm"
                        Case VBEfile.Document
                                Let path = pathArray(VBEfolder.Els) & VBComponent.name & ".cls"
                        Case Else
                                Let path = pathArray(VBEfolder.Els) & VBComponent.name & ".txt"
                 End Select
                Call VBComponent.Export(path)
        Next VBComponent
        GoTo ExecuteProcedure
ErrorHandle:
        On Error GoTo 0
ExecuteProcedure:
End Sub

Public Static Sub importAllVBEfiles()
        On Error GoTo ErrorHandle
        Dim VBComponent As Object
        Dim path As String
        'ActiveWorkbook.VBComponents.Remove VBComp
  
        For Each VBComponent In ActiveWorkbook.VBProject.VBComponents
                If VBComponent.name = "AutoVBE1" Then
                        VBComponent.name = "AutoVBE"
                End If
        Next VBComponent
'        With CreateObject("Scripting.FileSystemObject") 'Without add script runtime
'                If .FolderExists(path) Then
'                       .GetFolder
'                End If
'        End With
        GoTo ExecuteProcedure
ErrorHandle:
        On Error GoTo 0
ExecuteProcedure:
End Sub

' Assessor

Private Static Property Get getFilePaths() As Variant
        Let getFilePaths = Application.GetOpenFilename _
                (fileFilter:=FILE_FILLTER, _
                title:=IMPORT, _
                filterIndex:=INDEX, _
                MultiSelect:=True)
End Property
