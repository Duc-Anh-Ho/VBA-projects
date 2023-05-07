VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "InfoConstants"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Check README.md for more information
Option Explicit
'Declare Variables
'My Info
Private Const VERSION As String = "v2.2.0"
Private Const UPDATE_DATE  As String = "2022/10/11"
Private Const AUTHOR_PROMPT As String = "AUTHOR: "
Private Const AUTHOR_NAME As String = "DANH"
'Addin Info
Private Const ADDIN_SHORT_NAME As String = "Danh-Tools" ' ' <-- ***DON'T ADD _AddIn IN THIS NAME, BECAUSE OF UNKNOW ERROR EX. Danh_Tools_Addin
Private Const ADDIN_EXTENSION As String = ".xlam"
'BuildIn Properties
'Private Const TOOL_TITLE As String = "" ' "Title"
'Private Const TOOL_SUBJECT As String = "" ' "Subject"
'Private Const TOOL_AUTHOR As String = "" ' "Author"
'Private Const TOOL_KEYWORDS As String = "" ' "Keywords"
'Private Const TOOL_COMMENTS As String = "" ' "Comments"
'Private Const TOOL_TEMPLATE As String = "" ' "Template"
'Private Const TOOL_LAST_AUTHOR As String = "" ' "Last Author"
'Private Const TOOL_REVISION_NUMBER As String = "" ' "Revision Number"
'Private Const TOOL_APPLICATION_NAME As String = "" ' "Application Name"
'Private Const TOOL_CREATION_DATE As String = "" ' "Creation Date"
'Private Const TOOL_LAST_SAVE_TIME As String = "" ' "Last Save Time"
'Private Const TOOL_SECURITY As String = "" ' "Security"
'Private Const TOOL_CATEGORY As String = "" ' "Category"
'Private Const TOOL_FORMAT As String = "" ' "Format"
'Private Const TOOL_MANAGER As String = "" ' "Manager"
'Private Const TOOL_COMPANY As String = "" ' "Company"
'Private Const TOOL_HYPERLINK_BASE As String = "" ' "Hyperlink Base"

'Initialize Variables
Private Function hasVariables() As Boolean
'
End Function
'Constructor
Private Sub Class_Initialize()

End Sub
'Destructor
Private Sub Class_Terminate()
'
End Sub

'ASSESSORS

Public Property Get getVersion() As String
    Let getVersion = VERSION
End Property

Public Property Get getUpdate() As String
    Let getUpdate = UPDATE_DATE
End Property

Public Property Get getPrompt() As String
    Let getPrompt = AUTHOR_PROMPT
End Property

Public Property Get getAuthor() As String
    Let getAuthor = AUTHOR_NAME
End Property

Public Property Get getAddinShortName() As String
    Let getAddinShortName = ADDIN_SHORT_NAME
End Property

Public Property Get getAddinExtension() As String
    Let getAddinExtension = ADDIN_EXTENSION
End Property

Public Property Get getAddinName() As String
    Let getAddinName = ADDIN_SHORT_NAME & ADDIN_EXTENSION
End Property

