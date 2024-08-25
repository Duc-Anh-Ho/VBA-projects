VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CustomUITag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Check README.md for more information
Option Explicit
'Declare Variables
Private id As String
Private size As String
Private Description As String
Private isEnabled As Boolean
Private image As String
Private isShowImage As Boolean
Private keytip As String
Private label As String
Private isShowLabel As Boolean
Private screentip As String
Private supertip As String
Private isVisible As Boolean

'Initialize Variables
'Constructor
Private Sub Class_Initialize()
'
End Sub
'Destructor
Private Sub Class_Terminate()
'
End Sub

'ASSESSORS
Public Property Get getID() As String
    Let getID = id
End Property
Public Property Get getSize() As String
    Let getSize = size
End Property
Public Property Get getDescription() As String
    Let getDescription = Description
End Property
Public Property Get getEnabled() As Boolean
    Let getEnabled = isEnabled
End Property
Public Property Get getImage() As String
    Let getImage = image
End Property
Public Property Get getShowImage() As Boolean
    Let getShowImage = isShowImage
End Property
Public Property Get getKeytip() As String
    Let getKeytip = keytip
End Property
Public Property Get getLabel() As String
    Let getLabel = label
End Property
Public Property Get getShowLabel() As Boolean
    Let getShowLabel = isShowLabel
End Property
Public Property Get getScreentip() As String
    Let getScreentip = screentip
End Property
Public Property Get getSupertip() As String
    Let getSupertip = supertip
End Property
Public Property Get getVisible() As Boolean
    Let getVisible = isVisible
End Property
'MUTATORS
Public Property Let letID(ByVal newID As String)
    Let id = newID
End Property
Public Property Let letSize(ByVal newSize As String)
    Let size = newSize
End Property
Public Property Let letDescription(ByVal newDescription As String)
    Let Description = newDescription
End Property
Public Property Let letEnabled(ByVal newEnabled As Boolean)
    Let isEnabled = newEnabled
End Property
Public Property Let letImage(ByVal newImage As String)
    Let image = newImage
End Property
Public Property Let letShowImage(ByVal newShowImage As Boolean)
    Let isShowImage = newShowImage
End Property
Public Property Let letKeytip(ByVal newKeytip As String)
    Let keytip = newKeytip
End Property
Public Property Let letLabel(ByVal newLabel As String)
    Let label = newLabel
End Property
Public Property Let letShowLabel(ByVal newShowLabel As Boolean)
    Let isShowLabel = newShowLabel
End Property
Public Property Let letScreentip(ByVal newScreentip As String)
    Let screentip = newScreentip
End Property
Public Property Let letSupertip(ByVal newSupertip As String)
    Let supertip = newSupertip
End Property
Public Property Let letVisible(ByVal newVisible As Boolean)
    Let isVisible = newVisible
End Property
'METHODS
'MAIN
