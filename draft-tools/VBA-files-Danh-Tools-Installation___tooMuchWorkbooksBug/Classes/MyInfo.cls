VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MyInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Check README.md for more information
Option Explicit
'Declare Variables
Private Const VERSION As String = "v2.0.0"
Private Const UPDATE_DATE  As String = "2022/10/11"
Private Const AUTHOR_PROMPT As String = "AUTHOR: "
Private Const AUTHOR_NAME As String = "DANH"
'Initialize Variables
Private Function hasVariables() As Boolean
'
End Function
'Constructor
Private Sub Class_Initialize()
'
End Sub
'Destructor
Private Sub Class_Terminate()
'
End Sub

'ASSESSORS

Public Property Get GetVersion() As String
    GetVersion = VERSION
End Property

Public Property Get getUpdate() As String
     getUpdate = UPDATE_DATE
End Property

Public Property Get getPrompt() As String
    getPrompt = AUTHOR_PROMPT
End Property

Public Property Get getAuthor() As String
    getAuthor = AUTHOR_NAME
End Property
