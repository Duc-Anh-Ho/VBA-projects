Attribute VB_Name = "BrakePasswordVBA"
'PENDDING
'
'Option Explicit
'
'Private Const PAGE_EXECUTE_READWRITE = &H40
'
'#If VBA7 Then 'khai bao an toan cho phien ban 64bit (them PtrSafe va LongPtr)
'    Private Declare PtrSafe Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" _
'            (Destination As LongPtr, Source As LongPtr, ByVal Length As LongPtr)
'
'    Private Declare PtrSafe Function VirtualProtect Lib "kernel32" (lpAddress As LongPtr, _
'            ByVal dwSize As LongPtr, ByVal flNewProtect As LongPtr, lpflOldProtect As LongPtr) As LongPtr
'
'    Private Declare PtrSafe Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As LongPtr
'
'    Private Declare PtrSafe Function GetProcAddress Lib "kernel32" (ByVal hModule As LongPtr, _
'            ByVal lpProcName As String) As LongPtr
'
'    Private Declare PtrSafe Function DialogBoxParam Lib "user32" Alias "DialogBoxParamA" (ByVal hInstance As LongPtr, _
'            ByVal pTemplateName As LongPtr, ByVal hWndParent As LongPtr, _
'            ByVal lpDialogFunc As LongPtr, ByVal dwInitParam As LongPtr) As Integer
'#Else ' khai bao cho phien ban 32bit
'    Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" _
'            (Destination As Long, Source As Long, ByVal Length As Long)
'
'    Private Declare Function VirtualProtect Lib "kernel32" (lpAddress As Long, _
'            ByVal dwSize As Long, ByVal flNewProtect As Long, lpflOldProtect As Long) As Long
'
'    Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
'
'    Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, _
'            ByVal lpProcName As String) As Long
'
'    Private Declare Function DialogBoxParam Lib "user32" Alias "DialogBoxParamA" (ByVal hInstance As Long, _
'            ByVal pTemplateName As Long, ByVal hWndParent As Long, _
'            ByVal lpDialogFunc As Long, ByVal dwInitParam As Long) As Integer
'#End If
'
'Dim HookBytes(0 To 5) As Byte
'Dim OriginBytes(0 To 5) As Byte
'Dim pFunc As Long
'Dim Flag As Boolean
'
'Private Function GetPtr(ByVal Value As Long) As Long
'    GetPtr = Value
'End Function
'
'Public Sub RecoverBytes()
'    If Flag Then MoveMemory ByVal pFunc, ByVal VarPtr(OriginBytes(0)), 6
'End Sub
'
'Public Function Hook() As Boolean
'    Dim TmpBytes(0 To 5) As Byte
'    Dim p As Long
'    Dim OriginProtect As Long
'
'    Hook = False
'
'    pFunc = GetProcAddress(GetModuleHandleA("user32.dll"), "DialogBoxParamA")
'
'
'    If VirtualProtect(ByVal pFunc, 6, PAGE_EXECUTE_READWRITE, OriginProtect) <> 0 Then
'
'        MoveMemory ByVal VarPtr(TmpBytes(0)), ByVal pFunc, 6
'        If TmpBytes(0) <> &H68 Then
'
'            MoveMemory ByVal VarPtr(OriginBytes(0)), ByVal pFunc, 6
'
'            p = GetPtr(AddressOf MyDialogBoxParam)
'
'            HookBytes(0) = &H68
'            MoveMemory ByVal VarPtr(HookBytes(1)), ByVal VarPtr(p), 4
'            HookBytes(5) = &HC3
'
'            MoveMemory ByVal pFunc, ByVal VarPtr(HookBytes(0)), 6
'            Flag = True
'            Hook = True
'        End If
'    End If
'End Function
'
'Private Function MyDialogBoxParam(ByVal hInstance As Long, _
'        ByVal pTemplateName As Long, ByVal hWndParent As Long, _
'        ByVal lpDialogFunc As Long, ByVal dwInitParam As Long) As Integer
'    If pTemplateName = 4070 Then
'        MyDialogBoxParam = 1
'    Else
'        RecoverBytes
'        MyDialogBoxParam = DialogBoxParam(hInstance, pTemplateName, _
'                           hWndParent, lpDialogFunc, dwInitParam)
'        Hook
'    End If
'End Function
'
''Run this sub to unprotected VBA
'Sub unprotected()
'    If Hook Then
'        MsgBox "VBA Project is unprotected!", vbInformation, "*****"
'    End If
'End Sub

