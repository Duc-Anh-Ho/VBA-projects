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
        "All files, *.*"
Const INDEX As Integer = 1 ' Default All VB files
Const TASK As String = "Import to VBE"
        
Public Sub importVBEfiles()
        Dim filePaths As Variant ' For MutilSelect
        Let filePaths = Application.GetOpenFilename _
                (fileFilter:=FILE_FILLTER, _
                title:=TASK, _
                filterIndex:=INDEX, _
                MultiSelect:=True)
        If TypeName(filePaths) = "Boolean" Then GoTo ExecuteProcedure
       
       Debug.Print "+ " & TypeName(filePaths)
       Dim path As Variant
       For Each path In filePaths
                Debug.Print "+ " & path
       Next path

ExecuteProcedure:
End Sub
