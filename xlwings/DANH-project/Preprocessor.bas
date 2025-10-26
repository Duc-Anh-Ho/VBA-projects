Attribute VB_Name = "Preprocessor"
'Preprocessor API Library Declaration
Option Private Module ' No one can access the code in this module from outside (much moduleName.FunctionName)
Option Explicit
' REF: https://learn.microsoft.com/en-us/office/vba/language/concepts/getting-started/compiler-constants?utm_source=chatgpt.com
#If MAC Then ' <-- Macintosh
    MsgBox "MacOS chay khong duoc, cai win di"
#ElseIf VBA7 Then '<-- Window - Office Ver > 2007
    ' Library kernel32
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal Milliseconds As LongPtr)
    Public Declare PtrSafe Function GetTickCount Lib "kernel32" () As LongPtr
    Public Declare PtrSafe Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" ( _
        ByRef Destination As Any _
        , ByRef Source As Any _
        , ByVal Length As Long _
    )
    Public Declare PtrSafe Function GlobalAlloc Lib "kernel32" ( _
        ByVal memoryAllocationFlags As Long _
        , ByVal numberOfBytesToAllocate As LongPtr _
    ) As LongPtr
    Public Declare PtrSafe Function GlobalLock Lib "kernel32" ( _
        ByVal memoryHandle As LongPtr _
    ) As LongPtr
    Public Declare PtrSafe Function GlobalUnlock Lib "kernel32" ( _
        ByVal memoryHandle As LongPtr _
    ) As Long
    Public Declare PtrSafe Function GlobalSize Lib "kernel32" ( _
        ByVal memoryHandle As LongPtr _
    ) As LongPtr
    Public Declare PtrSafe Function lstrcpy Lib "kernel32" ( _
        ByVal destinationPointer As Any _
        , ByVal sourcePointer As Any _
    ) As LongPtr
    ' Library user32
    Public Declare PtrSafe Function OpenClipboard Lib "user32" ( _
        ByVal windowHandleRequestingAccess As LongPtr _
    ) As Long
    Public Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
    Public Declare PtrSafe Function GetClipboardData Lib "user32" ( _
        ByVal clipboardDataFormat As Long _
    ) As LongPtr
    Public Declare PtrSafe Function EmptyClipboard Lib "user32" () As Long
    Public Declare PtrSafe Function SetClipboardData Lib "user32" ( _
        ByVal clipboardDataFormat As Long _
        , ByVal memoryHandle As LongPtr _
    ) As LongPtr
    ' Get Ribbon From Pointer Memory
    Public Function GetRibbon(ByVal ribbonName As name) As IRibbonUI
        Dim objRibbon As Object
        Dim lRibbonPointer As LongPtr
        Let lRibbonPointer = CLngPtr(Replace(ribbonName.RefersTo, "=", ""))
        Call CopyMemory( _
            Destination:=objRibbon _
            , Source:=lRibbonPointer _
            , Length:=LenB(lRibbonPointer) _
        )
        Set GetRibbon = objRibbon
        Set objRibbon = Nothing
    End Function
#ElseIf VBA6 Then  '<-- Window - Office Ver <= 2007
    ' Library kernel32
    Public Declare Sub Sleep Lib "kernel32" (ByVal Milliseconds As Long)
    Public Declare Function GetTickCount Lib "kernel32" () As Long
    Public Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" ( _
        ByRef Destination As Any _
        , ByRef Source As Any _
        , ByVal Length As Long _
    )
    Public Declare Function GlobalAlloc Lib "kernel32" ( _
        ByVal memoryAllocationFlags As Long _
        , ByVal numberOfBytesToAllocate As Long _
    ) As Long
    Public Declare Function GlobalLock Lib "kernel32" ( _
        ByVal memoryHandle As Long
    ) As Long
    Public Declare Function GlobalUnlock Lib "kernel32" ( _
        ByVal memoryHandle As Long) As Long
    Public Declare Function GlobalSize Lib "kernel32" ( _
        ByVal memoryHandle As Long) As Long
    Public Declare Function lstrcpy Lib "kernel32" ( _
        ByVal destinationPointer As Any _
        , ByVal sourcePointer As Any _
    ) As Long
    ' Library user32
    Public Declare Function OpenClipboard Lib "user32" ( _
        ByVal windowHandleRequestingAccess As Long _
    ) As Long
    Public Declare Function CloseClipboard Lib "user32" () As Long
    Public Declare Function GetClipboardData Lib "user32" ( _
        ByVal clipboardDataFormat As Long _
    ) As Long
    Public Declare Function EmptyClipboard Lib "user32" () As Long
    Public Declare Function SetClipboardData Lib "user32" ( _
        ByVal clipboardDataFormat As Long _
        , ByVal memoryHandle As Long _
    ) As Long
    ' Get Ribbon From Pointer Memory
    Public Function GetRibbon(ByVal lRibbonPointer As Long) As IRibbonUI
        Dim objRibbon As Object
        Call CopyMemory( _
            Destination:=objRibbon _
            , Source:=lRibbonPointer _
            , Length:=LenB(lRibbonPointer) _
        )
        Set GetRibbon = objRibbon
        Set objRibbon = Nothing
    End Function
#ElseIf Win64 Then
#ElseIf Win32 Then
#ElseIf Win16 Then
#End If

' Trick: Define dummy LongPtr type for backward compatibility (no native LongPtr support)
#If VBA7 = 0 Then
    Public Enum LongPtr
        [_]
    End Enum
#End If
' Get Ribbon From Pointer Memory
'#If VBA7 Then
'Public Function GetRibbon(ByVal lRibbonPointer As LongPtr) As Object
'#Else
'Public Function GetRibbon(ByVal lRibbonPointer As Long) As Object
'#End If
'        Dim objRibbon As Object
'        CopyMemory _
'            Destination:=objRibbon, _
'            Source:=lRibbonPointer, _
'            Length:=LenB(lRibbonPointer)
'        Set GetRibbon = objRibbon
'        Set objRibbon = Nothing
'End Function






