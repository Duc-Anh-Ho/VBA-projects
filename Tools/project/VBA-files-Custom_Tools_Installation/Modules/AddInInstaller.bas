Attribute VB_Name = "AddInInstaller"
' Using for add in installer
Public Const ADD_IN_NAME = "VBA_Khong_Chuyen"  '<--Customize name of Add-in
Public Const VERSION = "20.0.3.1"
Public Const TOOLNAME = "Super Tools"

Option Explicit
    Dim thisPath As String
    Dim thisFile As String
    Dim thisPathAndFile As String
    Dim targetPath As String
    Dim targetFile As String
    Dim targetFileNoEx As String
    Public targetPathAndFile As String  'Public because need to recycle at Workbook close event
    Dim YorN As VbMsgBoxResult
    
'Making this sub for shortage code
Sub Initialize_Name()
    thisPath = ThisWorkbook.path & "\"
    thisFile = ThisWorkbook.name 'ex. thisFile = "AddInInstaller.xlam"
    thisFile = Left(thisFile, InStr(thisFile, ".") - 1) 'ex. thisFile = "AddInInstaller"
    thisPathAndFile = thisPath & thisFile
    'targetPath = Environ("Appdata") & "\Microsoft\AddIns" 'ex. targetPath = "C:\Users\Admin\AppData\Roaming\Microsoft\AddIns"
    targetPath = Application.UserLibraryPath 'ex. targetPath = "C:\Users\Admin\AppData\Roaming\Microsoft\AddIns\"
    targetFileNoEx = ADD_IN_NAME
    targetFile = targetFileNoEx & ".xlam"
    targetPathAndFile = targetPath & targetFile
End Sub
'Making this sub for shortage code
Sub install_sub()
    Application.DisplayAlerts = False
    ThisWorkbook.SaveAs fileName:=targetPathAndFile, FileFormat:=xlOpenXMLAddIn ' xlOpenXMLAddIn = 55
    AddIns(targetFileNoEx).Installed = True
    Application.DisplayAlerts = True
End Sub
'Making this sub for shortage code
Sub uninstall_sub()
    Application.DisplayAlerts = False 'Check case delete this disable after deleting this itself.
    If Len(Dir(targetPathAndFile)) = 0 Then
        MsgBox "Add-in: " & targetFileNoEx & " does not exist!", vbOKOnly, "Confirmation"
        End 'End all vba
    Else
        If AddIns(targetFileNoEx).Installed = True Then 'check the case already disabled manually
            AddIns(targetFileNoEx).Installed = False
        End If
         Application.DisplayAlerts = True 'back to default
    End If
End Sub
'Callback for Install onAction
Sub Install_AddIn()
    Call Initialize_Name 'shortage code
    If Len(Dir(targetPathAndFile)) = 0 Then 'if brand new add-in
        Call install_sub
        MsgBox "New Add-in: " & targetFileNoEx & " successfully Installed!", vbOKOnly, "Confirmation"
        ThisWorkbook.Close SaveChanges:=False
    Else 'if Add-in already exists then the user will decide if will replace it or not
        YorN = MsgBox("Add-in: " & targetFileNoEx & " allready exists!" & vbNewLine & "Do you want to update version -" & VERSION & "- ?", vbYesNo)
        If YorN = vbYes Then 'deactivate the Add-in if it is activated
            Call uninstall_sub
            Call install_sub
            MsgBox "Add-in: " & targetFileNoEx & " successfully updated to version -" & VERSION & "-!", vbOKOnly, "Confirmation"
            ThisWorkbook.Close SaveChanges:=False
        Else
            ' End 'Exit vba code
            ThisWorkbook.Close SaveChanges:=False 'Close when don't wanna install
        End If
    End If
End Sub


Sub Uninstall_AddIn()
    Call Initialize_Name 'shortage code
    Call uninstall_sub
    ThisWorkbook.Close SaveChanges:=False
End Sub





