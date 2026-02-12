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
Private Error As Utils_Error
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
Private Const WIN_SPLASH As String = "\"
Private Const MACOS As String = "Mac"
Private Const MAC_REGEX As String = MACOS & "intosh*"
Private Const MAC_WHERE As String = "which "
Private Const MAC_SPLASH As String = "/"
Private Const GIT_CMD As String = "git "
' REF: https://learn.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/locktypeenum
Private Enum LOCK_TYPE
    BatchOptimistic = 4 ' adLockBatchOptimistic
    Optimistic = 3 ' adLockOptimistic
    Pessimistic = 2 ' adLockPessimistic
    ReadOnly = 1 ' adLockReadOnly
    Unspecified = -1 ' adLockUnspecified
End Enum
' REF: https://learn.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/cursortypeenum
Private Enum CURSOR_TYPE
    OpenDynamic = 2 ' adOpenDynamic
    OpenForwardOnly = 0 ' adOpenForwardOnly
    OpenKeyset = 1 ' adOpenKeyset
    OpenStatic = 3 ' adOpenStatic
    OpenUnspecified = -1 ' adOpenUnspecified
End Enum
' REF: https://www.w3schools.com/asp/prop_conn_state.asp
Private Enum OBJECT_STATE
    StateClosed = 0 ' adStateClosed
    StateOpen = 1 ' adStateOpen
    StateConnecting = 2 ' adStateConnecting
    StateExecuting = 4 ' adStateExecuting
    StateFetching = 8 ' adStateFetching
End Enum
' REF: https://learn.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/schemaenum
Private Enum SCHEMA
    Asserts = 0 'adSchemaAsserts
    Catalogs = 1 'adSchemaCatalogs
    CharacterSets = 2 'adSchemaCharacterSets
    CheckConstraints = 5 'adSchemaCheckConstraints
    Collations = 3 'adSchemaCollations
    ColumnPrivileges = 13 'adSchemaColumnPrivileges
    Columns = 4 'adSchemaColumns
    ColumnsDomainUsage = 11 'adSchemaColumnsDomainUsage
    ConstraintColumnUsage = 6 'adSchemaConstraintColumnUsage
    ConstraintTableUsage = 7 'adSchemaConstraintTableUsage
    Cubes = 32 'adSchemaCubes
    DBInfoKeywords = 30 'adSchemaDBInfoKeywords
    DBInfoLiterals = 31 'adSchemaDBInfoLiterals
    Dimensions = 33 'adSchemaDimensions
    ForeignKeys = 27 'adSchemaForeignKeys
    Hierarchies = 34 'adSchemaHierarchies
    Indexes = 12 'adSchemaIndexes
    KeyColumnUsage = 8 'adSchemaKeyColumnUsage
    Levels = 35 'adSchemaLevels
    Measures = 36 'adSchemaMeasures
    Members = 38 'adSchemaMembers
    PrimaryKeys = 28 'adSchemaPrimaryKeys
    ProcedureColumns = 29 'adSchemaProcedureColumns
    ProcedureParameters = 26 'adSchemaProcedureParameters
    Procedures = 16 'adSchemaProcedures
    Properties = 37 'adSchemaProperties
    ProviderSpecific = -1 'adSchemaProviderSpecific
    ProviderTypes = 22 'adSchemaProviderTypes
    ReferentialConstraints = 9 'adSchemaReferentialConstraints
    Schemata = 17 'adSchemaSchemata
    SQLLanguages = 18 'adSchemaSQLLanguages
    Statistics = 19 'adSchemaStatistics
    TableConstraints = 10 'adSchemaTableConstraints
    TablePrivileges = 14 'adSchemaTablePrivileges
    Tables = 20 'adSchemaTables
    Translations = 21 'adSchemaTranslations
    Trustees = 39 'adSchemaTrustees
    UsagePrivileges = 15 'adSchemaUsagePrivileges
    ViewColumnUsage = 24 'adSchemaViewColumnUsage
    Views = 23 'adSchemaViews
    ViewTableUsage = 25 'adSchemaViewTableUsage
End Enum
' REF: https://learn.microsoft.com/en-us/windows/win32/shell/folder-copyhere
Private Enum FILE_OPRATIONS_FLAGS
    MULTIDESTFILES = &H1
    SILENT = &H4
    RENAMEONCOLLISION = &H8
    NOCONFIRMATION = &H10
    WANTMAPPINGHANDLE = &H20
    ALLOWUNDO = &H40
    FILESONLY = &H80
    SIMPLEPROGRESS = &H100
    NOCONFIRMMKDIR = &H200
    NOERRORUI = &H400
    NOCOPYSECURITYATTRIBS = &H800
    NORECURSION = &H1000
    NO_CONNECTED_ELEMENTS = &H2000
    WANTNUKEWARNING = &H4000
    NORECURSEREPARSE = &H8000
End Enum
' REF: https://learn.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/streamtypeenum
Private Enum STREAM_TYPE
    Binary = 1 ' adTypeBinary
    Text = 2 ' adTypeText
End Enum
' --- Member ---
Private m_pass As String
Private m_fail As String
Private m_skip As String
Private m_gitPath As String
Private m_wherePath As String
Private m_splash As String
Private m_os As String
' --- Constructor ---
Private Sub Class_Initialize()
    Let m_pass = DEFAULT_PASS
    Let m_fail = DEFAULT_FAIL
    Let m_skip = DEFAULT_SKIP
    Set Error = New Utils_Error
End Sub
' --- Destructor ---
Private Sub Class_Terminate()
    Set Error = Nothing
End Sub
' --- Property Accessor / Mutator ---
Friend Property Let PASS(ByRef value As String)
    If VBA.Strings.Len(value) > 0 Then
        Let m_pass = value
    Else
        Let m_pass = DEFAULT_PASS
    End If
End Property

Friend Property Let FAIL(ByRef value As String)
    If VBA.Strings.Len(value) > 0 Then
        Let m_fail = value
    Else
        Let m_fail = DEFAULT_FAIL
    End If
End Property

Friend Property Let SKIP(ByRef value As String)
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
    If m_os = vbNullString Then
        If Excel.Application.OperatingSystem Like WIN_REGEX Then
            Let m_os = WINOS
        ElseIf Excel.Application.OperatingSystem Like MAC_REGEX Then
            Let m_os = MACOS
        Else
            Call Error.Fail(ECODE.Err_OperationNotSupported, "Utils_Test.OS")
        End If
    End If
    Let OS = m_os
End Property

Friend Property Get WINDOW() As Boolean
    Let WINDOW = WINOS
End Property

Friend Property Get MAC() As Boolean
    Let MAC = MACOS
End Property

Friend Property Get GIT() As String
    If m_gitPath = vbNullString Then
        Let m_gitPath = Me.Where(GIT_CMD)
    End If
    Let GIT = m_gitPath
End Property

Friend Property Get SPLASH() As String
    If m_splash = vbNullString Then
        Select Case Me.OS
            Case WINOS
                Let m_splash = WIN_SPLASH
            Case MACOS
                Let m_splash = MAC_SPLASH
            Case Else
                Call Error.Fail(ECODE.Err_OperationNotSupported, "Utils_Test.SPLASH")
        End Select
    End If
    Let SPLASH = m_splash
End Property

' TODO Merge with PowerShellController
Friend Property Get Where(ByRef execName As String) As String
    If m_wherePath <> vbNullString Then
        Let Where = m_wherePath
        Exit Property
    End If
    Dim scripts As Utils_Scripts: Set scripts = New Utils_Scripts
    Dim wShell As Object: Set wShell = scripts.WShell
    Dim command As String
    Dim execute As Object
    Select Case Me.OS
        Case WINOS
            Let command = WIN_WHERE & execName
            Set execute = wShell.exec(command)
            If Not execute.StdOut.AtEndOfStream Then Let m_wherePath = execute.StdOut.ReadLine
            ' Resolve .lnk shortcut to target executable
            Const LNK_EXTENSION As String = ".lnk"
            If LCase$(Right$(m_wherePath, 4)) = LNK_EXTENSION  Then
                Dim shellLink As Object: Set shellLink = wShell.CreateShortcut(m_wherePath)
                Let m_wherePath = shellLink.TargetPath
            End If
        Case MACOS
            Let command = MAC_WHERE & execName
            Set execute = wShell.exec(command)
            If Not execute.StdOut.AtEndOfStream Then Let m_wherePath = execute.StdOut.ReadLine
    End Select
Cleanup:
    Set execute = Nothing
    Set wShell = Nothing
    Set scripts = Nothing
    Set shellLink = Nothing
    Let Where = m_wherePath
End Property

' --- Helper Methods ---
' TODO Create sheetUtils or sheetController for this

' --- Fuctions / Subroutines ---
Friend Function IsExistPath(ByRef path As String) As Boolean
    Let IsExistPath = (VBA.fileSystem.Dir(path) <> vbNullString)
End Function

Friend Function IsExistFolder(ByRef folderPath As String) As Boolean
    If VBA.Strings.Trim$(folderPath) = vbNullString Then
        Let IsExistFolder = False
    ElseIf Dir(folderPath, vbDirectory) = vbNullString Then
        Let IsExistFolder = False
    ElseIf (VBA.fileSystem.GetAttr(folderPath) And vbDirectory) = vbDirectory Then
        Let IsExistFolder = True
    Else
        Let IsExistFolder = False
    End If
End Function

Friend Function IsExistFile(ByRef fileName As String) As Boolean
    If VBA.Strings.Trim$(fileName) = vbNullString Then
        Let IsExistFile = False
    ElseIf Dir(fileName, vbNormal) = vbNullString Then
        Let IsExistFile = False
    ElseIf (VBA.fileSystem.GetAttr(fileName) And vbDirectory) = 0 Then
        Let IsExistFile = True
    Else
        Let IsExistFile = False
    End If
End Function

Friend Function CountFilesInFolder( _
    ByRef folderPath As String _
    , Optional ByRef pattern As String = "*.*" _
) As Long
    If Not Me.IsExistFolder(folderPath) Then Exit Function
    Dim fileName As String: Let fileName = Dir(folderPath & WIN_SPLASH & pattern, vbNormal)
    Dim count As Long
    Do While fileName <> ""
        If (VBA.fileSystem.GetAttr(folderPath & WIN_SPLASH & fileName) And vbDirectory) <> vbDirectory Then
            Let count = count + 1
        End If
        Let fileName = Dir
    Loop
    Let CountFilesInFolder = count
End Function

Friend Function CountFoldersInFolder(ByRef folderPath As String) As Long
    If Not Me.IsExistFolder(folderPath) Then Exit Function
    Dim folderName As String: Let folderName = Dir(folderPath & "\*", vbDirectory)
    Dim count As Long
    Do While folderName <> ""
        If folderName <> "." And folderName <> ".." Then
            If (VBA.fileSystem.GetAttr(folderPath & WIN_SPLASH & folderName) And vbDirectory) = vbDirectory Then
                Let count = count + 1
            End If
        End If
        Let folderName = Dir
    Loop
    Let CountFoldersInFolder = count
End Function

Friend Function CountInFolder(ByRef folderPath As String) As Long
    If Not Me.IsExistFolder(folderPath) Then Exit Function
    Dim folderName As String: Let folderName = Dir(folderPath & "\*", vbDirectory)
    Dim count As Long
    Do While folderName <> ""
        If folderName <> "." And folderName <> ".." Then Let count = count + 1
        Let folderName = Dir
    Loop
    Let CountInFolder = count
End Function

Friend Function GetParentFolder(ByVal fileName As String) As String
    Dim posSlash As Long: Let posSlash = InStrRev(fileName, MAC_SPLASH)
    Dim posBack As Long: Let posBack = InStrRev(fileName, WIN_SPLASH)
    Dim pos As Long: Let pos = IIf(posSlash > posBack, posSlash, posBack)
    Do While pos > 0 And (Mid$(fileName, pos, 1) = WIN_SPLASH Or Mid$(fileName, pos, 1) = MAC_SPLASH)
        Let pos = pos - 1
    Loop
    If pos > 0 Then
        Let GetParentFolder = VBA.Strings.left$(fileName, pos)
    Else
        Let GetParentFolder = vbNullString
    End If
End Function

Friend Function GetShortName(ByRef fileName As String) As String
    If Not Me.IsExistFile(fileName) Then
        Call Error.Fail(ECODE.Err_FileNotFound, "Utils_Test.GetShortName")
        Exit Function
    End If
    Dim posSlash As Long: Let posSlash = InStrRev(fileName, MAC_SPLASH)
    Dim posBack As Long: Let posBack = InStrRev(fileName, WIN_SPLASH)
    Dim pos As Long: Let pos = IIf(posSlash > posBack, posSlash, posBack)
    If pos > 0 Then
        Let GetShortName = VBA.Strings.Mid$(fileName, pos + 1)
    Else
        Let GetShortName = fileName
    End If
End Function

Friend Sub CreateFolder(ByRef folderPath As String)
    If Me.IsExistFolder(folderPath) Then: Exit Sub
    Dim parts() As String: Let parts = Split(folderPath, WIN_SPLASH)
    Dim currentPath As String: Let currentPath = vbNullString
    Dim i As Long: For i = LBound(parts) To UBound(parts)
        If parts(i) <> vbNullString Then
            If currentPath = vbNullString Then
                Let currentPath = parts(i)
            Else
                Let currentPath = currentPath & WIN_SPLASH & parts(i)
            End If
            If Not Me.IsExistFolder(currentPath) Then
                Call VBA.fileSystem.MkDir(currentPath)
            End If
        End If
    Next i
End Sub

Friend Sub RenameFile(ByRef oldFileName As String, ByRef newFileName As String)
    If Not Me.IsExistFile(oldFileName) Then
        Call Error.Fail(ECODE.Err_FileNotFound, "Utils_Test.RenameFile")
        Exit Sub
    End If
    If Me.IsExistFile(newFileName) Then
        Call VBA.fileSystem.Kill(newFileName)
    End If
    Name oldFileName As newFileName
End Sub

' NOTE: Old method
Friend Function CopyFileSmall( _
    ByRef sourceFile As String _
    , ByRef outputFile As String _
) As Boolean
    If Not IsExistFile(sourceFile) Then
        Call Error.Fail(ECODE.Err_FileNotFound, "Utils_Test.CopyFileSmall")
        Exit Function
    End If
    Dim outputFolder As String: Let outputFolder = Me.GetParentFolder(outputFile)
    If Not Me.IsExistFolder(outputFolder) Then Call Me.CreateFolder(outputFolder)
    If Me.IsExistFile(outputFile) Then Call VBA.fileSystem.Kill(outputFile)
    Call VBA.fileSystem.FileCopy(sourceFile, outputFile)
    Let CopyFileSmall = Me.IsExistFile(outputFile)
End Function

Friend Sub CopyFile(ByRef sourceFile As String, ByRef outputFile As String)
    Const BLOCKSIZE As Long = 65536 ' 64KB
    If Not IsExistFile(sourceFile) Then
        Call Error.Fail(ECODE.Err_FileNotFound, "Utils_Test.CopyFile")
        Exit Sub
    End If
    Dim outputFolder As String: Let outputFolder = Me.GetParentFolder(outputFile)
    If Not Me.IsExistFolder(outputFolder) Then Call Me.CreateFolder(outputFolder)
    If Me.IsExistFile(outputFile) Then Call VBA.fileSystem.Kill(outputFile)
    Dim sourceNum As Long: Let sourceNum = FreeFile
    Open sourceFile For Binary Access Read As #sourceNum
    Dim targetNum As Long: Let targetNum = FreeFile
    Open outputFile For Binary Access Write As #targetNum
    Dim bytesRead As Long
    Dim buffer() As Byte
    Do While Not EOF(sourceNum)
        ReDim buffer(1 To BLOCKSIZE)
        Let bytesRead = LOF(sourceNum) - Loc(sourceNum) + 1
        If bytesRead > BLOCKSIZE Then bytesRead = BLOCKSIZE
        Get #sourceNum, , buffer
        Put #targetNum, , buffer
    Loop
    Close #sourceNum
    Close #targetNum
End Sub

Friend Sub CopyFolder(ByRef sourceFolder As String, ByRef outputFolder As String)
    If Not Me.IsExistFolder(sourceFolder) Then
        Call Error.Fail(ECODE.Err_FolderNotFound, "Utils_Test.CopyFolder")
        Exit Sub
    End If
    Dim scripts As Utils_Scripts: Set scripts = New Utils_Scripts
    Dim fileSystem As Object: Set fileSystem = scripts.fileSystem
    If Me.IsExistFolder(outputFolder) Then Call Me.DeleteFolder(outputFolder)
    Call fileSystem.CopyFolder(sourceFolder, outputFolder)
End Sub

Friend Sub DeleteFolder(ByRef folderPath As String)
    If Not Me.IsExistFolder(folderPath) Then Exit Sub
    If Right$(folderPath, 1) = "\" Then
        Let folderPath = Left$(folderPath, Len(folderPath) - 1)
    End If
    Dim scripts As Utils_Scripts: Set scripts = New Utils_Scripts
    Dim fileSystem As Object: Set fileSystem = scripts.fileSystem
    Call fileSystem.DeleteFolder(folderPath, True)
End Sub

' HELPER FOR ZIPFLDR.DLL (deprecated)
Friend Function IsZipfldrPresent() As Boolean
    Dim sysPath As String
    ' Office 32-bit on 64-bit Windows
    If VBA.Environ$("PROCESSOR_ARCHITECTURE") = "x86" Then
        Let sysPath = Environ$("SystemRoot") & "\SysWOW64\zipfldr.dll"
    Else
        Let sysPath = Environ$("SystemRoot") & "\System32\zipfldr.dll"
    End If
    Let IsZipfldrPresent = (VBA.FileSystem.Dir(sysPath) <> "")
End Function

' HELPER FOR ZIPFLDR.DLL (deprecated)
Friend Function IsZipfldrRegistered() As Boolean
    Dim scripts As Utils_Scripts: Set scripts = New Utils_Scripts
    Dim wShell As Object: Set wShell = scripts.WShell
    Dim regPath As String: Let regPath = "HKEY_CLASSES_ROOT\CLSID\{E88DCCE0-B7B3-11d1-A9F0-00AA0060FA31}\InProcServer32"
    Dim checkCmd As String: Let checkCmd = CMD & "reg query """ & regPath & """ /ve"
    Dim execute As Object: Set execute = wShell.Exec(checkCmd)
    Dim result As String
    Do While Not execute.StdOut.AtEndOfStream
        result = result & execute.StdOut.ReadLine & vbCrLf
    Loop
    Let IsZipfldrRegistered = (InStr(LCase(result), "zipfldr.dll") > 0)
End Function

' HELPER FOR ZIPFLDR.DLL (deprecated)
Friend Sub RegisterZipfldr()
    Dim scripts As Utils_Scripts: Set scripts = New Utils_Scripts
    Dim wShell As Object: Set wShell = scripts.WShell
    Dim systemPath As String
    If VBA.Environ$("PROCESSOR_ARCHITECTURE") = "x86" Then
        Let systemPath = Environ$("SystemRoot") & "\SysWOW64\zipfldr.dll"
    Else
        Let systemPath = Environ$("SystemRoot") & "\System32\zipfldr.dll"
    End If
    Dim cmd As String: Let cmd = "regsvr32 /s """ & systemPath & """"
    Call wShell.Run(cmd, vbHide, True)
End Sub

Friend Sub ExtractExcelByZip(ByVal excelFileName As String, ByVal outputFolder As String)
    ' Check Input
    If Not Me.IsExistFile(excelFileName) Then
        Call Error.Fail(ECODE.Err_FileNotFound, "Utils_Test.ExtractExcelByZip")
        Exit Sub
    End If
    ' Create folder if not exist
    If Not Me.IsExistFolder(outputFolder) Then Call Me.CreateFolder(outputFolder)
    ' Excel to zip
    Dim sourceFolder As String: Let sourceFolder = Me.GetParentFolder(excelFileName)
    Dim tempFolder As String: tempFolder = sourceFolder & "\_temp"
    If Not Me.IsExistFolder(tempFolder) Then Call Me.CreateFolder(tempFolder)
    Dim tempZip As String: tempZip = tempFolder & "\_temp_file.zip"
    Call Me.CopyFile(excelFileName, tempZip)
    ' Handle Zip with PowerShell Expand-Archive 
    Dim scripts As Utils_Scripts: Set scripts = New Utils_Scripts
    Dim wShell As Object: Set wShell = scripts.WShell
    Const POWERSHELL_CONST As String = "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"
    Dim POWERSHELL As String: Let POWERSHELL = Me.Where("powershell")
    If POWERSHELL = vbNullString Then Let POWERSHELL = POWERSHELL_CONST
    If Not IsExistFile(POWERSHELL) Then
        Call Error.Fail(ECODE.Err_FileNotFound, "Utils_Test.ExtractExcelByZip", POWERSHELL)
    End If
    Dim powershellCommand As String: Let powershellCommand = POWERSHELL _
        & " -NoProfile -Command ""Expand-Archive" _
        & " -Path '" & Replace(tempZip, "'", "''") & "'" _
        & " -DestinationPath '" & Replace(outputFolder, "'", "''") & "'" _
        & " -Force"""
    Dim execute As Object: Set execute = wShell.Exec(powershellCommand)
    Dim output As String
    Dim t As Double: t = Timer
    Const TIMEOUT As Double = 30 ' seconds
    Do
        DoEvents
        Do While Not execute.StdOut.AtEndOfStream
            Let output = output & execute.StdOut.ReadLine & vbCrLf
        Loop
        Do While Not execute.StdErr.AtEndOfStream
            Let output = output & execute.StdErr.ReadLine & vbCrLf
        Loop
        ' Limit Timeout
        If Timer - t > TIMEOUT Then
            Call Error.Fail(ECODE.Err_Unknown, "Utils_Test.ExtractExcelByZip", "Expand-Archive timed out", output)
            Exit Sub
        End If
    Loop While execute.Status = 0
    ' Check output
    If Dir(outputFolder, vbDirectory) = "" Then
        Call Error.Fail(ECODE.Err_Unknown, "Utils_Test.ExtractExcelByZip", "Expand-Archive failed", output)
        Exit Sub
    End If
    ' Delete Zip
    Call Me.DeleteFolder(tempFolder)
End Sub

Friend Function DiffCheck( _
    ByRef fileName1 As String _
    , ByRef fileName2 As String _
) As Boolean
    Const BLOCKSIZE As Long = 65536 ' 64KB
    Dim scripts1 As Utils_Scripts: Set scripts1 = New Utils_Scripts
    Dim scripts2 As Utils_Scripts: Set scripts2 = New Utils_Scripts
    Dim file1 As Object: Set file1 = scripts1.fileStream
    Let file1.Type = STREAM_TYPE.Binary
    Call file1.Open
    Call file1.LoadFromFile(fileName1)
    Dim file2 As Object: Set file2 = scripts2.fileStream
    Let file2.Type = STREAM_TYPE.Binary
    Call file2.Open
    Call file2.LoadFromFile(fileName2)
    Do While Not file1.EOS Or Not file2.EOS
        Dim block1() As Byte: Let block1 = file1.Read(BLOCKSIZE)
        Dim block2() As Byte: Let block2 = file2.Read(BLOCKSIZE)
        Dim bytesRead1 As Long: Let bytesRead1 = UBound(block1) - LBound(block1) + 1
        Dim bytesRead2 As Long: Let bytesRead2 = UBound(block2) - LBound(block2) + 1
        If bytesRead1 <> bytesRead2 Then
            Let DiffCheck = False
            GoTo Cleanup
        End If
        Dim i As Long: For i = LBound(block1) To UBound(block1)
            If block1(i) <> block2(i) Then
                Let DiffCheck = False
                GoTo Cleanup
            End If
        Next i
    Loop
    Let DiffCheck = True
Cleanup:
    Call file1.Close
    Call file2.Close
    Set file1 = Nothing
    Set scripts1 = Nothing
    Set file2 = Nothing
    Set scripts2 = Nothing
End Function

Friend Function GetEndOfFile(ByRef fileName As String) As String
    Dim scripts As Utils_Scripts: Set scripts = New Utils_Scripts
    Dim shell As Object: Set shell = scripts.WShell
    ' If Not IsExistFile(fileName) Then
    '     Let GetEndOfFile = "File Not Found!"
    '     Goto CleanUp
    ' End If
    Const NKF_CONST As String = "D:\share\ASS\Tasks\template\autotest.exe"
    Const GUESS As String = " --guess "
    Dim NKF As String: Let NKF = Me.Where("nkf")
    If NKF = vbNullString Then Let NKF = NKF_CONST
    If Not IsExistFile(NKF) Then
        Let GetEndOfFile = "NKF Not Found!"
        GoTo Cleanup
    End If
    Dim command As String: Let command = """" & NKF & """" & GUESS & """" & fileName & """"
    Dim exec As Object: Set exec = shell.exec(command)
    Dim result As String: Let result = exec.StdOut.ReadAll
    If result = vbNullString Then Let result = "File not found!"
    Let GetEndOfFile = result
Cleanup:
    Set shell = Nothing
    Set scripts = Nothing
End Function

' TODO DANG bi loi encodign file
Friend Function GetEndOfFiles(ByRef fileNames() As String) As Variant
    Const NKF_CONST As String = "D:\share\ASS\Tasks\template\autotest.exe"
    Const GUESS As String = " --guess "
    Const LIMIT As Byte = 10
    Dim NKF As String: Let NKF = NKF_CONST
    Dim results() As String: ReDim results(LBound(fileNames) To UBound(fileNames))
    Dim scripts As Utils_Scripts: Set scripts = New Utils_Scripts
    Dim shell As Object: Set shell = scripts.WShell
    Dim dict As Object: Set dict = scripts.Dictionary
    Dim batch() As String: ReDim batch(1 To LIMIT)
    Dim count As Long, i As Long, j As Long
    For i = LBound(fileNames) To UBound(fileNames)
        Let count = count + 1
        Let batch(count) = """" & fileNames(i) & """"
        If count = LIMIT Or i = UBound(fileNames) Then
            Dim command As String: Let command = """" & NKF & """" & GUESS & Join(batch, " ")
            Dim exec As Object: Set exec = shell.exec(command)
            Dim buffer As String: Let buffer = ""
            Do While Not exec.StdOut.AtEndOfStream
                Let buffer = buffer & exec.StdOut.ReadLine & vbCrLf
            Loop
            Dim lines() As String: Let lines = Split(Trim(buffer), vbCrLf)
            For j = 0 To count - 1
                If j <= UBound(lines) Then
                    Dim parts() As String: Let parts = Split(lines(j), ": ", 2)
                    If UBound(parts) >= 1 Then
                        ' TODO Fix cai nay sai encode
                        Debug.Print parts(0)
                        Let dict(parts(0)) = parts(1)
                    Else
                        Let dict(Replace(batch(j + 1), """", "")) = ""
                    End If
                End If
            Next j
            Let count = 0
            Erase batch
            ReDim batch(1 To LIMIT)
        End If
    Next i
    For i = LBound(fileNames) To UBound(fileNames)
        Debug.Print "File:", fileNames(i), "Exists:", dict.Exists(fileNames(i))
        If dict.Exists(fileNames(i)) Then
            Let results(i) = dict(fileNames(i))
        Else
            Let results(i) = ""
        End If
    Next i
    Let GetEndOfFiles = results
Cleanup:
    Set dict = Nothing
    Set shell = Nothing
    Set scripts = Nothing
End Function

Friend Function ReadLastLines( _
        ByRef fileName As String _
        , ByRef lineCount As Long _
        , Optional charset As String = "utf-8" _
    ) As Variant
    Const BLOCKSIZE As Long = 65536 ' 64KB
    Dim scripts As Utils_Scripts: Set scripts = New Utils_Scripts
    Dim fileStream As Object: Set fileStream = scripts.fileStream
    Let fileStream.Type = STREAM_TYPE.Text
    Let fileStream.Charset = charset ' Unix override to 'x-euc-jp'
    Call fileStream.Open
    Call fileStream.LoadFromFile(fileName)
    Dim fileSize As Long: Let fileSize = fileStream.Size
    Dim position As Long: Let position = IIf(fileSize - BLOCKSIZE > 0, fileSize - BLOCKSIZE, 0)
    Let fileStream.Position = position
    Dim blockText As String: Let blockText = Replace(fileStream.ReadText, vbCrLf, vbLf)
    Call fileStream.Close
    Set fileStream = Nothing
    Dim lines() As String: Let lines = Split(blockText, vbCrLf)
    Dim startLine As Long: Let startLine = IIf(UBound(lines) - lineCount + 1 > 0, UBound(lines) - lineCount + 1, 0)
    Dim result() As String: ReDim result(0 To UBound(lines) - startLine)
    Dim index As Long: Let index = 0
    Dim i As Long: For i = startLine To UBound(lines)
        Let result(index) = lines(i)
        Let index = index + 1
    Next i
    Let ReadLastLines = result
End Function

Friend Function IsSheetExistInExcel( _
    ByRef fullName As String _
    , ByRef sheetName As String _
    , Optional ByRef connection As Object _
) As Boolean
    Dim isTempConnection As Boolean: isTempConnection = False
    If connection Is Nothing Then
        Dim scripts As Utils_Scripts: Set scripts = New Utils_Scripts
        Set connection = scripts.Connection
        Call connection.Open( _
            "Provider=Microsoft.ACE.OLEDB.12.0;" & _
            "Data Source=" & fullName & ";" & _
            "Extended Properties=""Excel 12.0 Xml;HDR=No;IMEX=1"";" _
        )
        Let isTempConnection = True
    End If
    Dim sheetNameLow As String: sheetNameLow = LCase(sheetName & "$")
    Dim Recordset As Object: Set Recordset = connection.OpenSchema(SCHEMA.Tables)
    Do Until Recordset.EOF
        Dim tableName As String: Let tableName = LCase(Recordset.Fields("TABLE_NAME").Value)
        If tableName = sheetNameLow Or tableName = "'" & sheetNameLow & "'" Then
            Let IsSheetExistInExcel = True
            GoTo Cleanup
        End If
        Recordset.MoveNext
    Loop
    Let IsSheetExistInExcel = False
Cleanup:
    If isTempConnection Then
        If connection.State = OBJECT_STATE.StateOpen Then connection.Close
        Set connection = Nothing
    End If
End Function

Friend Function SearchInExcelByPattern( _
        ByRef fullName As String _
        , ByRef sheetName As String _
        , ByRef targetCol As Long _
        , ByVal pattern As String _
        , Optional ByRef isRevert As Boolean = False _
        , Optional ByRef startPattern As String = vbNullString _
        , Optional ByRef stopPattern As String = vbNullString _
    ) As String
    Dim result As String: Let result = "Text Not Found!"
    Dim sSQL As String
    If InStr(sheetName, " ") > 0 Then
        Let sSQL = "SELECT * FROM ['" & sheetName & "$']"
    Else
        Let sSQL = "SELECT * FROM [" & sheetName & "$]"
    End If
    Dim scripts As Utils_Scripts: Set scripts = New Utils_Scripts
    Dim connection As Object: Set connection = scripts.Connection
    Dim Recordset As Object: Set Recordset = scripts.Recordset
    Call connection.Open( _
        "Provider=Microsoft.ACE.OLEDB.12.0;" & _
        "Data Source=" & fullName & ";" & _
        "Extended Properties=""Excel 12.0 Xml;HDR=No;IMEX=1"";" _
    )
    If Not Me.IsSheetExistInExcel(fullName, sheetName, connection) Then
        Let result = "Sheet Not Found!"
        Exit Function
    End If
    Call Recordset.Open( _
        sSQL _
        , connection _
        , CURSOR_TYPE.OpenStatic _
        , LOCK_TYPE.ReadOnly _
    )
    If Recordset.BOF And Recordset.EOF Then GoTo Cleanup
    Dim isStart As Boolean: Let isStart = True
    If startPattern <> vbNullString Then Let isStart = False
    If isRevert Then
        Call Recordset.MoveLast
        Do While Not Recordset.BOF
            With Recordset.Fields(targetCol)
                If stopPattern <> vbNullString And InStr(.Value, stopPattern) > 0 Then Exit Do
                If Not isStart Then
                    If InStr(.Value, startPattern) > 0 Then Let isStart = True
                ElseIf InStr(.Value, pattern) > 0 Then
                    Let result = .Value
                    Exit Do
                End If
            End With
            Call Recordset.MovePrevious
        Loop
    Else
        Call Recordset.MoveFirst
        Do While Not Recordset.EOF
            With Recordset.Fields(targetCol)
                If stopPattern <> vbNullString And InStr(.Value, stopPattern) > 0 Then Exit Do
                If Not isStart Then
                    If InStr(.Value, startPattern) > 0 Then Let isStart = True
                ElseIf InStr(.Value, pattern) > 0 Then
                    Let result = .Value
                    Exit Do
                End If
            End With
            Call Recordset.MoveNext
        Loop
    End If
Cleanup:
    If Not Recordset Is Nothing Then
        If Recordset.State = OBJECT_STATE.StateOpen Then Call Recordset.Close
        Set Recordset = Nothing
    End If
    If Not connection Is Nothing Then
        If connection.State = OBJECT_STATE.StateOpen Then Call connection.Close
        Set connection = Nothing
    End If
    Let SearchInExcelByPattern = result
End Function

' HELPER
Friend Function ColumnNameToIndex(colName As String) As Long
    Dim result As Long
    Let colName = VBA.Strings.UCase$(colName)
    Dim i As Long:  For i = 1 To VBA.Strings.Len(colName)
        Let result = result * 26 + (Asc(Mid(colName, i, 1)) - Asc("A") + 1)
    Next i
    Let ColumnNameToIndex = result - 1
End Function

Friend Function ExtracByRegex( _
    text As String _
    , pattern As String _
) As String
    Dim scripts As Utils_Scripts: Set scripts = New Utils_Scripts
    Dim regex As Object: Set regex = scripts.RegexExpression
    Let regex.Pattern = pattern
    Let regex.Global = False
    Let regex.IgnoreCase = True
    If regex.Test(text) Then
        Dim matches As Object: Set matches = regex.Execute(text)
        Let ExtracByRegex = matches(0).SubMatches(0)
    Else
        Let ExtracByRegex = ""
    End If
End Function

Friend Function SearchInFileWithPattern( _
    ByRef fileName As String _
    , ByRef pattern As String _
    , Optional ByRef lineAmount As Long = 0 _
    , Optional ByRef isRevert As Boolean = False _
    , Optional ByRef charset As String = "utf-8" _
) As Collection
    Const BLOCKSIZE As Long = 65536 ' 64KB
    Dim result As New Collection
    Dim scripts As Utils_Scripts: Set scripts = New Utils_Scripts
    Dim fileStream As Object: Set fileStream = scripts.FileStream
    Let fileStream.Type = STREAM_TYPE.Text
    Let fileStream.Charset = charset
    Call fileStream.Open
    Call fileStream.LoadFromFile(fileName)
    Dim fileSize As Long: Let fileSize = fileStream.Size
    Dim position As Long
    Dim blockText As String
    Dim lines() As String
    Dim i As Long
    If isRevert Then
        Dim chunk As String
        Dim tmpPosition As Long
        Let tmpPosition = fileSize
        Do While tmpPosition > 0
            Let position = IIf(tmpPosition - BLOCKSIZE > 0, tmpPosition - BLOCKSIZE, 0)
            Let fileStream.Position = position
            Let chunk = fileStream.ReadText(IIf(tmpPosition - position, tmpPosition - position, BLOCKSIZE))
            Let chunk = Replace(chunk, vbCrLf, vbLf)
            Let lines = Split(chunk, vbLf)
            For i = UBound(lines) To LBound(lines) Step -1
                If InStr(lines(i), pattern) > 0 Then
                    Call result.Add(lines(i))
                    If lineAmount > 0 And result.Count >= lineAmount Then Exit Do
                End If
            Next i
            Let tmpPosition = position
        Loop
    Else
        Do Until fileStream.EOS
            Let blockText = Replace(fileStream.ReadText(BLOCKSIZE), vbCrLf, vbLf)
            Let lines = Split(blockText, vbLf)
            For i = LBound(lines) To UBound(lines)
                If InStr(lines(i), pattern) > 0 Then
                    Call result.Add(lines(i))
                    If lineAmount > 0 And result.Count >= lineAmount Then Exit Do
                End If
            Next i
        Loop
    End If
    Call fileStream.Close
    Set fileStream = Nothing
    Set SearchInFileWithPattern = result
End Function

