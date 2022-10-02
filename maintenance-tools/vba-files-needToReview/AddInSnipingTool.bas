Attribute VB_Name = "AddInSnipingTool"
Option Explicit
''Old code install add in
Sub install_add_in()
    Dim mypath As String
    Dim strfile As String
    Dim fileName As String
    
    Let mypath = ThisWorkbook.Path
    Let fileName = "OLD_Snipping_Tool"   'Add-in filename
    Let strfile = fileName & ".xlam"
    
    file_to_copy = mypath & "\" & strfile
    
    folder_to_copy = Environ("Appdata") & "MicrosoftAddIns"
    
    copied_file = folder_to_copy & "\" & strfile
    
    'Check if add-in is installed
    If Len(Dir(copied_file)) = 0 Then
    
    'if add-in does not exist then copy the file
    FileCopy file_to_copy, copied_file
    AddIns(fileName).Installed = True
    MsgBox "Add-in installed"
    
    Else
    
    'if add-in already exists then the user will decide if will replace it or not
    x = MsgBox("Add-in allready exists ! Replace ?", vbYesNo)
    
        If x = vbNo Then
            Exit Sub
        ElseIf x = vbYes Then
            
            'deactivate the add-in if it is activated
            If AddIns(fileName).Installed = True Then
                AddIns(fileName).Installed = False
            End If
            
            'delete the old file
            Kill copied_file
            
            'copy the new file
            FileCopy file_to_copy, copied_file
            AddIns(fileName).Installed = True
            MsgBox "New Add-in Installed !"
    
        End If
    
    End If

End Sub


 
