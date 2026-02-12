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
Private Const IMMEDIATE_WINDOW_PREFIX As String = "immediate"
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
' REF: https://learn.microsoft.com/en-us/office/vba/language/reference/visual-basic-add-in-model/properties-visual-basic-add-in-model
Private Enum COMPONENT_TYPE
    StandardModule = 1 ' vbext_ct_StdModule
    ClassModule = 2 ' vbext_ct_ClassModule
    MicrosoftForm = 3 ' vbext_ct_MSForm
    ActiveXDesigner = 11 ' vbext_ct_ActiveXDesigner
    DocumentModule = 100 ' vbext_ct_Document
End Enum
Private Enum WINDOW_TYPE
	CodeWindow = 0 ' vbext_wt_CodeWindow - Code window
	Designer = 1 ' vbext_wt_Designer - Designer
	Browser = 2 ' vbext_wt_Browser - Object Browser
	Watch = 3 ' vbext_wt_Watch - Immediate window
	Locals = 4 ' vbext_wt_Locals - Project window
	Immediate = 5 ' vbext_wt_Immediate - Properties window
	ProjectWindow = 6 ' vbext_wt_ProjectWindow - Find dialog box
	PropertyWindow = 7 ' vbext_wt_PropertyWindow - Search and Replace dialog box
	Find = 8 ' vbext_wt_Find - Linked window frame
	FindReplace = 9 ' vbext_wt_FindReplace - Main window
	Toolbox = 10 ' vbext_wt_Toolbox - Watch window
	LinkedWindowFrame = 11 ' vbext_wt_LinkedWindowFrame - Locals window
	MainWindow = 12 ' vbext_wt_MainWindow - Toolbox
	ToolWindow = 15 ' vbext_wt_ToolWindow - Tool window
End Enum
' --- Helper Methods ---
Private Sub closeWindowByPrefix(ByRef prefix As String)
    Dim window As Object
    Dim windowName As String
    For Each window In Application.VBE.Windows
        Let windowName = LCase$(left$(window.caption, Len(prefix)))
        If windowName = prefix Then
            Call window.Close
            Exit For
        End If
    Next window
End Sub

' --- Main ---
Friend Sub CloseProjectWindow()
    Call closeWindowByPrefix(PROJECT_WINDOW_PREFIX)
End Sub

Friend Sub ClosePropertiesWindow()
    Call closeWindowByPrefix(PROPERTIES_WINDOW_PREFIX)
End Sub

Friend Sub CloseImmediateWindow()
    Call closeWindowByPrefix(IMMEDIATE_WINDOW_PREFIX)
End Sub

Friend Sub CloseAllWindows()
    Dim window As Object ' VBIDE.Window
    For Each window In Application.VBE.Windows
        If window.Type <> WINDOW_TYPE.Immediate Then window.Close
    Next window
    ' Call CloseProjectWindow
    ' Call ClosePropertiesWindow
    ' Call CloseAllComponents
End Sub

Friend Sub ListComponentsByTypeName(Optional ByRef typeName As String = vbNullString)
    Dim component As Object
    For Each component In Application.VBE.ActiveVBProject.VBComponents
        If typeName <> vbNullString Then
            Select Case VBA.Strings.LCase$(typeName)
                Case STANDARD_MODULE, STANDARD_MODULE_EXTENSION, STANDARD_MODULE_DOT_EXTENSION
                    If component.Type = COMPONENT_TYPE.StandardModule Then Debug.Print component.Name
                Case CLASS_MODULE, CLASS_MODULE_EXTENSION, CLASS_MODULE_DOT_EXTENSION
                    If component.Type = COMPONENT_TYPE.ClassModule Then Debug.Print component.Name
                Case MICROSOFT_FORM, MICROSOFT_FORM_EXTENSION, MICROSOFT_FORM_DOT_EXTENSION
                    If component.Type = COMPONENT_TYPE.MicrosoftForm Then Debug.Print component.Name
                Case ACTIVEX_DESIGNER, ACTIVEX_DESIGNER_EXTENSION, ACTIVEX_DESIGNER_DOT_EXTENSION
                    If component.Type = COMPONENT_TYPE.ActiveXDesigner Then Debug.Print component.Name
                Case DOCUMENT_MODULE, DOCUMENT_MODULE_EXTENSION, DOCUMENT_MODULE_DOT_EXTENSION
                    If component.Type = COMPONENT_TYPE.DocumentModule Then Debug.Print component.Name
                Case Else
                    ' Do nothing for unknown type
            End Select
        Else
            Select Case component.Type
                Case COMPONENT_TYPE.StandardModule
                    Debug.Print component.Name & " - " & STANDARD_MODULE
                Case COMPONENT_TYPE.ClassModule
                    Debug.Print component.Name & " - " & CLASS_MODULE
                Case COMPONENT_TYPE.MicrosoftForm
                    Debug.Print component.Name & " - " & MICROSOFT_FORM
                Case COMPONENT_TYPE.ActiveXDesigner
                    Debug.Print component.Name & " - " & ACTIVEX_DESIGNER
                Case COMPONENT_TYPE.DocumentModule
                    Debug.Print component.Name & " - " & DOCUMENT_MODULE
                Case Else
                    Debug.Print component.Name & " - Unknown Type"
            End Select
        End If
    Next component
End Sub

Friend Sub OpenComponentByName(ByRef name As String)
    Dim component As Object
    For Each component In Application.VBE.ActiveVBProject.VBComponents
        If VBA.Strings.LCase$(component.Name) = VBA.Strings.LCase$(name) Then
            Call Application.VBE.ActiveVBProject.VBComponents(name).Activate
            Exit Sub
        End If
    Next component
End Sub

Friend Sub CloseAllComponents()
    Dim window As Object ' VBIDE.Window
    For Each window In Application.VBE.Windows
        If window.Type = WINDOW_TYPE.CodeWindow Then window.Close
    Next window
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

Friend Sub PrintParsedDictionary(ByRef Dict As Object)
    Dim key As Variant
    For Each key In Dict.Keys
        Debug.Print "Key: " & key & " | Value: " & Dict(key)
    Next key
End Sub

Friend Sub PrintParsedStringArray(ByRef StrArr As Variant)
    Dim i As Long: For i = LBound(StrArr) To UBound(StrArr)
        Debug.Print "Index: " & i & " | Value: " & StrArr(i)
    Next i
End Sub

Friend Sub PrintParsedCollection(ByRef col As Collection)
    Dim i As Long
    For i = 1 To col.Count
        Debug.Print "Index: " & i & " | Value: " & col(i)
    Next i
End Sub

Friend Sub ToggleToolbars()
    Dim vbeCommandBar As Object
    Dim anyVisible As Boolean
    Dim toolbars(0 To 3) As String
    Let toolbars(0) = "Standard"
    Let toolbars(1) = "Edit"
    Let toolbars(2) = "Debug"
    Let toolbars(3) = "UserForm"
    Dim index As Long: For index = 0 To UBound(toolbars)
        Set vbeCommandBar = GetCommandBar(toolbars(index))
        If Not vbeCommandBar Is Nothing Then
            If vbeCommandBar.Visible Then Let anyVisible = True
        End If
    Next index
    For index = 0 To UBound(toolbars)
        Set vbeCommandBar = GetCommandBar(toolbars(index))
        If Not vbeCommandBar Is Nothing Then
            Let vbeCommandBar.Visible = Not anyVisible
        End If
    Next index
End Sub

Friend Function GetCommandBar(ByVal name As String) As Object
    Dim matched As Boolean: Let matched = False
    Dim commandBar As Object: For Each commandBar In Application.VBE.CommandBars
        If StrComp(commandBar.Name, name, vbTextCompare) = 0 Then
            Set GetCommandBar = commandBar
            Let matched = True
            Exit For
        End If
    Next commandBar
    If Not matched Then Set GetCommandBar = Nothing
End Function

