Attribute VB_Name = "Module1"
Option Explicit

Sub test()
'    Let Application.AddIns("VBA_Khong_Chuyen").Installed = True
'    Let Application.AddIns("Danh_Tools_Addin").Installed = True
    Dim Auto As Object
    Dim fileSystem As Object
    Dim installFile As String
    Dim addinFileFullName As String
    Dim ai As AddIn
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    Let installFile = "S:\VBA-projects\released-tools\auto-addIn\Auto_add_VBA_AddIn.xlsb"
    Let addinFileFullName = "C:\Users\Danh1\AppData\Roaming\Microsoft\AddIns\Danh-Tools-test.xlam"
'    Call fileSystem.CopyFile( _
'        Source:=installFile, _
'        Destination:=addinFileFullName)
'    Application.AddIns.Remove fileName:=addinFileFullName

    Debug.Print "OK"
'    For Each ai In Application.AddIns
'        Debug.Print ai.Name
'        Debug.Print ai.Path
'    Next ai
    Set Auto = New AutoAddIn
    Set Auto = Nothing
End Sub
