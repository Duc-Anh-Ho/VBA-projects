VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Utils_VBE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Check README.md for more information
Option Explicit
' --- Constant ---
Private Const PROJECT_WINDOW_PREFIX As String = "project - "
Private Const PROPERTIES_WINDOW_PREFIX As String = "properties - "
Private Const STANDARD_MODULE As String = "module"
Private Const STANDARD_MODULE_EXTENSION As String = "bas"
Private Const STANDARD_MODULE_DOT_EXTENSION As String = ".bas"
Private Const CLASS_MODULE As String = "class"
Private Const CLASS_MODULE_EXTENSION As String = "cls"
Private Const CLASS_MODULE_DOT_EXTENSION As String = ".cls"
Private Const MICROSOFT_FORM As String = "userForm"
Private Const MICROSOFT_FORM_EXTENSION As String = "frm"
Private Const MICROSOFT_FORM_DOT_EXTENSION As String = ".frm"
Private Const ACTIVEX_DESIGNER As String = "activex designer"
Private Const ACTIVEX_DESIGNER_EXTENSION As String = "frx"
Private Const ACTIVEX_DESIGNER_DOT_EXTENSION As String = ".frx"
Private Const DOCUMENT_MODULE As String = "document"
Private Const DOCUMENT_MODULE_EXTENSION As String = "cls"
Private Const DOCUMENT_MODULE_DOT_EXTENSION As String = ".cls"
Private Enum COMPONENT_TYPE
    StandardModule = 1 'vbext_ct_StdModule
    ClassModule = 2 'vbext_ct_ClassModule
    MicrosoftForm = 3 'vbext_ct_MSForm
    ActivexDesigner = 11 'vbext_ct_ActiveXDesigner
    DocumentModule = 100 'vbext_ct_Document
End Enum
' --- Helper Methods ---
Private Sub closeWindowByPrefix(ByRef prefix As String)
    Dim WINDOW As Object
    Dim windowName As String
    For Each WINDOW In Application.VBE.Windows
        Let windowName = LCase$(left$(WINDOW.caption, Len(prefix)))
        If windowName = prefix Then
            Call WINDOW.Close
            Exit For
        End If
    Next WINDOW
End Sub

' --- Main ---
Friend Sub CloseProjectWindow()
    Call closeWindowByPrefix(PROJECT_WINDOW_PREFIX)
End Sub

Friend Sub ClosePropertiesWindow()
    Call closeWindowByPrefix(PROPERTIES_WINDOW_PREFIX)
End Sub

Friend Sub CloseAllWindows()
    Call CloseProjectWindow
    Call ClosePropertiesWindow
End Sub

Friend Sub ListComponentsByTypeName(Optional ByRef typeName As String = vbNullString)
    Dim component As Object
    For Each component In Application.VBE.ActiveVBProject.VBComponents
        If typeName <> vbNullString Then
            Select Case VBA.Strings.LCase$(typeName)
                Case STANDARD_MODULE, STANDARD_MODULE_EXTENSION, STANDARD_MODULE_DOT_EXTENSION
                    If component.Type = COMPONENT_TYPE.StandardModule Then Debug.Print component.name
                Case CLASS_MODULE, CLASS_MODULE_EXTENSION, CLASS_MODULE_DOT_EXTENSION
                    If component.Type = COMPONENT_TYPE.ClassModule Then Debug.Print component.name
                Case MICROSOFT_FORM, MICROSOFT_FORM_EXTENSION, MICROSOFT_FORM_DOT_EXTENSION
                    If component.Type = COMPONENT_TYPE.MicrosoftForm Then Debug.Print component.name
                Case ACTIVEX_DESIGNER, ACTIVEX_DESIGNER_EXTENSION, ACTIVEX_DESIGNER_DOT_EXTENSION
                    If component.Type = COMPONENT_TYPE.ActivexDesigner Then Debug.Print component.name
                Case DOCUMENT_MODULE, DOCUMENT_MODULE_EXTENSION, DOCUMENT_MODULE_DOT_EXTENSION
                    If component.Type = COMPONENT_TYPE.DocumentModule Then Debug.Print component.name
                Case Else
                    ' Do nothing for unknown type
            End Select
        Else
            Select Case component.Type
                Case COMPONENT_TYPE.StandardModule
                    Debug.Print component.name & " - " & STANDARD_MODULE
                Case COMPONENT_TYPE.ClassModule
                    Debug.Print component.name & " - " & CLASS_MODULE
                Case COMPONENT_TYPE.MicrosoftForm
                    Debug.Print component.name & " - " & MICROSOFT_FORM
                Case COMPONENT_TYPE.ActivexDesigner
                    Debug.Print component.name & " - " & ACTIVEX_DESIGNER
                Case COMPONENT_TYPE.DocumentModule
                    Debug.Print component.name & " - " & DOCUMENT_MODULE
                Case Else
                    Debug.Print component.name & " - Unknown Type"
            End Select
        End If
    Next component
End Sub

Friend Sub OpenComponentByName(ByRef name As String)
    Dim component As Object
    For Each component In Application.VBE.ActiveVBProject.VBComponents
        If VBA.Strings.LCase$(component.name) = VBA.Strings.LCase$(name) Then
            Call Application.VBE.ActiveVBProject.VBComponents(name).Activate
            Debug.Print name & " opened."
            Exit Sub
        End If
    Next component
    Debug.Print name & " not found."
End Sub

Friend Sub ClearImmediateWindowUnix()
    Debug.Print String(65535, vbCr) ' Fill Immediate Window with new lines same as Unix
End Sub

Friend Sub ClearImmediateWindow()
    Call VBA.Interaction.SendKeys("^g ^a {BACKSPACE}") ' Ctrl+G, Ctrl+A, Backspace
End Sub

'VBE - DEBUGGER
Public Sub DebugAssert(condition As Boolean)
#If debugger Then
    Debug.Assert condition
#End If
End Sub

Friend Sub PrintParsedDictionary(ByRef dict As Object)
    Dim key As Variant
    For Each key In dict.Keys
        Debug.Print "Key: " & key & " | Value: " & dict(key)
    Next key
End Sub

Friend Sub PrintParsedStringArray(ByRef StrArr As Variant)
    Dim i As Long: For i = LBound(StrArr) To UBound(StrArr)
        Debug.Print "Index: " & i & " | Value: " & StrArr(i)
    Next i
End Sub

Friend Sub PrintParsedCollection(ByRef col As Collection)
    Dim i As Long
    For i = 1 To col.count
        Debug.Print "Index: " & i & " | Value: " & col(i)
    Next i
End Sub
