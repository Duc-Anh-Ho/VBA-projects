VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Utils_Clipboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Check README.md for more information
Option Explicit
' --- Constants ---
' REF: https://learn.microsoft.com/en-us/windows/win32/dataxchg/standard-clipboard-formats
Private Enum CLIPBOARD_FORMATS
    TEXT = 1                 ' Plain text (ANSI)
    BITMAP = 2               ' Bitmap image handle
    METAFILEPICT = 3         ' Metafile picture format
    SYLK = 4                 ' Symbolic Link (legacy)
    DIF = 5                  ' Data Interchange Format
    TIFF = 6                 ' Tagged Image File Format
    OEMTEXT = 7              ' Text using OEM character set
    DIB = 8                  ' Device Independent Bitmap
    PALETTE = 9              ' Color palette
    PENDATA = 10             ' Pen data
    RIFF = 11                ' RIFF audio data
    WAVE = 12                ' WAVE audio data
    UNICODETEXT = 13         ' Text (Unicode UTF-16)
    ENHMETAFILE = 14         ' Enhanced metafile
    HDROP = 15               ' File list (used by drag/drop)
    LOCALE = 16              ' Locale identifier
    DIBV5 = 17               ' Device Independent Bitmap v5
End Enum
Public Enum GLOBAL_MEMORY
    FIXED = &H0           ' Memory cannot be moved
    MOVEABLE = &H2        ' Memory block can be moved
    ZEROINIT = &H40       ' Zero-initialize allocated memory
End Enum
' --- Dependencies ---
Private Script As Utils_Scripts
' --- Object ---
Private Clipboard As Object
' --- Constructor ---
Private Sub Class_Initialize()
    Set Script = New Utils_Scripts
    Set Clipboard = Script.Clipboard
End Sub
' --- Destructor ---
Private Sub Class_Terminate()
    Set Script = Nothing
    Set Clipboard = Nothing
End Sub
' --- Properties ---

' --- Main ---
Friend Sub SaveByCOM(Optional ByRef text As String = vbNullString)
    Call Clipboard.SetText(text)
    Call Clipboard.PutInClipboard
End Sub

Friend Function LoadByCOM() As String
    Call Clipboard.GetFromClipboard
    Let LoadByCOM = Clipboard.GetText()
End Function

Friend Sub SaveByAPI(Optional ByRef text As String = vbNullString)
    Dim clipboardMemoryHandle As LongPtr
    Dim lockedMemoryPointer As LongPtr
    Dim byteSize As LongPtr

    If Len(text) = 0 Then Exit Sub
    Let byteSize = (Len(text) + 1) * 2 ' Unicode: 2 bytes per char + null

    Let clipboardMemoryHandle = GlobalAlloc( _
        GLOBAL_MEMORY.MOVEABLE Or GLOBAL_MEMORY.ZEROINIT _
        , byteSize _
    )
    If clipboardMemoryHandle = 0 Then Exit Sub

    Let lockedMemoryPointer = GlobalLock(clipboardMemoryHandle)
    If lockedMemoryPointer = 0 Then Exit Sub

    Call CopyMemory(ByVal lockedMemoryPointer, ByVal StrPtr(text), CLng(byteSize))
    Call GlobalUnlock(clipboardMemoryHandle)

    If OpenClipboard(0) Then
        Call EmptyClipboard
        Call SetClipboardData(CLIPBOARD_FORMATS.UNICODETEXT, clipboardMemoryHandle)
        Call CloseClipboard
    End If
End Sub

Friend Function LoadByAPI() As String
    Dim clipboardMemoryHandle As LongPtr, lockedMemoryPointer As LongPtr
    Dim byteSize As LongPtr
    Dim buffer As String

    If Not OpenClipboard(0) Then Exit Function

    Let clipboardMemoryHandle = GetClipboardData(CLIPBOARD_FORMATS.UNICODETEXT)
    If clipboardMemoryHandle = 0 Then GoTo Cleanup

    Let lockedMemoryPointer = GlobalLock(clipboardMemoryHandle)
    If lockedMemoryPointer = 0 Then GoTo Cleanup

    Let byteSize = GlobalSize(clipboardMemoryHandle)
    Let buffer = String$(CLng(byteSize) \ 2, vbNullChar)
    Call CopyMemory(ByVal StrPtr(buffer), ByVal lockedMemoryPointer, CLng(byteSize))
    Call GlobalUnlock(clipboardMemoryHandle)

Cleanup:
    Call CloseClipboard
    LoadByAPI = Replace(buffer, vbNullChar, vbNullString)
End Function
