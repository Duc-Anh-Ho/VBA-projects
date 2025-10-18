VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Utils_Test"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Check README.md for more information
Option Explicit
' --- Dependencies ---
Private Script As Utils_Scripts
' --- Object ---
Private WithEvents targetWorksheet As Worksheet
' --- Constant ---
Private Const DEFAULT_PASS As String = "PASS"
Private Const DEFAULT_FAIL As String = "FAIL"
Private Const DEFAULT_SKIP As String = "SKIP"
Private Const CMD As String = "cmd /c "
Private Const WINOS As String = "Windows"
Private Const WIN_REGEX As String = WINOS & "*"
Private Const WIN_WHERE As String = CMD & "where "
Private Const MACOS As String = "Mac"
Private Const MAC_REGEX As String = MACOS & "intosh*"
Private Const MAC_WHERE As String = "which "
Private Const GIT_CMD As String = "git "
' --- Member ---
Private m_pass As String
Private m_fail As String
Private m_skip As String
Private m_gitPath As String
' --- Constructor ---
Private Sub Class_Initialize()
    Let m_pass = DEFAULT_PASS
    Let m_fail = DEFAULT_FAIL
    Let m_skip = DEFAULT_SKIP
    Set Script = New Utils_Scripts
End Sub
' --- Destructor ---
Private Sub Class_Terminate()
    Set Script = Nothing
End Sub
' --- Property Accessor / Mutator ---
Friend Property Let PASS(ByVal value As String)
    If VBA.Strings.Len(value) > 0 Then
        Let m_pass = value
    Else
        Let m_pass = DEFAULT_PASS
    End If
End Property

Friend Property Let FAIL(ByVal value As String)
    If VBA.Strings.Len(value) > 0 Then
        Let m_fail = value
    Else
        Let m_fail = DEFAULT_FAIL
    End If
End Property

Friend Property Let SKIP(ByVal value As String)
    If VBA.Strings.Len(value) > 0 Then
        Let m_skip = value
    Else
        Let m_skip = DEFAULT_SKIP
    End If
End Property

Friend Property Get PASS() As String
    Let PASS = m_pass
End Property

Friend Property Get FAIL() As String
    Let FAIL = m_fail
End Property

Friend Property Get SKIP() As String
    Let SKIP = m_skip
End Property

' TODO Merge With PowerShellController
Friend Property Get OS() As String
    If Excel.Application.OperatingSystem Like WIN_REGEX Then
        Let OS = WINOS
    ElseIf Excel.Application.OperatingSystem Like MAC_REGEX Then
        Let OS = MACOS
    Else
        Let OS = vbNullString
    End If
End Property

Friend Property Get window() As Boolean
    Let window = WINOS
End Property

Friend Property Get MAC() As Boolean
    Let MAC = MACOS
End Property

Friend Property Get GIT() As String
    If VBA.Strings.Len(m_gitPath) = 0 Then
        Let m_gitPath = Me.Where(GIT_CMD)
    End If
    Let GIT = m_gitPath
End Property

' --- Helper Methods ---
' TODO Create sheetUtils or sheetController for this

' --- Fuctions / Subroutines ---
Friend Function IsExistPath(ByRef path As String) As Boolean
    Let IsExistPath = (VBA.fileSystem.Dir(path) <> vbNullString)
End Function

' TODO Merge with PowerShellController
Friend Function Where(ByRef execName As String) As String
    Dim WShell As Object: Set WShell = Script.WShell
    Dim opSys As String: Let opSys = Me.OS
    Dim command As String
    Select Case opSys
        Case WINOS
            Let command = WIN_WHERE & execName
            Let Where = WShell.exec(command).StdOut.ReadLine
        Case MACOS
            Let command = MAC_WHERE & execName
            Let Where = WShell.exec(command).StdOut.ReadLine
        Case Else
            Let Where = vbNullString
    End Select
End Function



