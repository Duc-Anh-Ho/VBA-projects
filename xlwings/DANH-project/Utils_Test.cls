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
Attribute targetWorksheet.VB_VarHelpID = -1
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
    adLockBatchOptimistic = 4
    adLockOptimistic = 3
    adLockPessimistic = 2
    adLockReadOnly = 1
    adLockUnspecified = -1
End Enum
' REF: https://learn.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/cursortypeenum
Private Enum CURSOR_TYPE
    adOpenDynamic = 2
    adOpenForwardOnly = 0
    adOpenKeyset = 1
    adOpenStatic = 3
    adOpenUnspecified = -1
End Enum
' REF: https://www.w3schools.com/asp/prop_conn_state.asp
Private Enum OBJECT_STATE
    adStateClosed = 0
    adStateOpen = 1
    adStateConnecting = 2
    adStateExecuting = 4
    adStateFetching = 8
End Enum
' REF: https://learn.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/schemaenum
Private Enum SCHEMA
    adSchemaAsserts = 0
    adSchemaCatalogs = 1
    adSchemaCharacterSets = 2
    adSchemaCheckConstraints = 5
    adSchemaCollations = 3
    adSchemaColumnPrivileges = 13
    adSchemaColumns = 4
    adSchemaColumnsDomainUsage = 11
    adSchemaConstraintColumnUsage = 6
    adSchemaConstraintTableUsage = 7
    adSchemaCubes = 32
    adSchemaDBInfoKeywords = 30
    adSchemaDBInfoLiterals = 31
    adSchemaDimensions = 33
    adSchemaForeignKeys = 27
    adSchemaHierarchies = 34
    adSchemaIndexes = 12
    adSchemaKeyColumnUsage = 8
    adSchemaLevels = 35
    adSchemaMeasures = 36
    adSchemaMembers = 38
    adSchemaPrimaryKeys = 28
    adSchemaProcedureColumns = 29
    adSchemaProcedureParameters = 26
    adSchemaProcedures = 16
    adSchemaProperties = 37
    adSchemaProviderSpecific = -1
    adSchemaProviderTypes = 22
    adSchemaReferentialConstraints = 9
    adSchemaSchemata = 17
    adSchemaSQLLanguages = 18
    adSchemaStatistics = 19
    adSchemaTableConstraints = 10
    adSchemaTablePrivileges = 14
    adSchemaTables = 20
    adSchemaTranslations = 21
    adSchemaTrustees = 39
    adSchemaUsagePrivileges = 15
    adSchemaViewColumnUsage = 24
    adSchemaViews = 23
    adSchemaViewTableUsage = 25
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
            Call Error.FAIL(ECODE.Err_OperationNotSupported, "Utils_Test.OS")
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
                Call Error.FAIL(ECODE.Err_OperationNotSupported, "Utils_Test.SPLASH")
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
    Dim Scripts As Utils_Scripts: Set Scripts = New Utils_Scripts
    Dim wShell As Object: Set wShell = Scripts.wShell
    Dim command As String
    Dim execObject As Object
    Select Case Me.OS
        Case WINOS
            Let command = WIN_WHERE & execName
            Set execObject = wShell.exec(command)
            If Not execObject.StdOut.AtEndOfStream Then
                Let m_wherePath = execObject.StdOut.ReadLine
            Else
                Call Error.FAIL(ECODE.Err_FileNotFound, "Utils_Test.Where")
            End If
        Case MACOS
            Let command = MAC_WHERE & execName
            Set execObject = wShell.exec(command)
            If Not execObject.StdOut.AtEndOfStream Then
                Let m_wherePath = execObject.StdOut.ReadLine
            Else
                Call Error.FAIL(ECODE.Err_FileNotFound, "Utils_Test.Where")
            End If
        Case Else
            Call Error.FAIL(ECODE.Err_FileNotFound, "Utils_Test.Where")
    End Select
Cleanup:
    Let Where = m_wherePath
    Set execObject = Nothing
    Set wShell = Nothing
    Set Scripts = Nothing
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

Friend Function IsExistFile(ByRef filename As String) As Boolean
    If VBA.Strings.Trim$(filename) = vbNullString Then
        Let IsExistFile = False
    ElseIf Dir(filename, vbNormal) = vbNullString Then
        Let IsExistFile = False
    ElseIf (VBA.fileSystem.GetAttr(filename) And vbDirectory) = 0 Then
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
    Dim filename As String: Let filename = Dir(folderPath & WIN_SPLASH & pattern, vbNormal)
    Dim count As Long
    Do While filename <> ""
        If (VBA.fileSystem.GetAttr(folderPath & WIN_SPLASH & filename) And vbDirectory) <> vbDirectory Then
            Let count = count + 1
        End If
        Let filename = Dir
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

Friend Function GetParentFolder(ByVal filename As String) As String
    Dim posSlash As Long: Let posSlash = InStrRev(filename, MAC_SPLASH)
    Dim posBack As Long: Let posBack = InStrRev(filename, WIN_SPLASH)
    Dim pos As Long: Let pos = IIf(posSlash > posBack, posSlash, posBack)
    Do While pos > 0 And (Mid$(filename, pos, 1) = WIN_SPLASH Or Mid$(filename, pos, 1) = MAC_SPLASH)
        Let pos = pos - 1
    Loop
    If pos > 0 Then
        Let GetParentFolder = VBA.Strings.left$(filename, pos)
    Else
        Let GetParentFolder = vbNullString
    End If
End Function

Friend Function GetShortName(ByRef filename As String) As String
    If Not Me.IsExistFile(filename) Then
        Call Error.FAIL(ECODE.Err_FileNotFound, "Utils_Test.GetShortName")
        Exit Function
    End If
    Dim posSlash As Long: Let posSlash = InStrRev(filename, MAC_SPLASH)
    Dim posBack As Long: Let posBack = InStrRev(filename, WIN_SPLASH)
    Dim pos As Long: Let pos = IIf(posSlash > posBack, posSlash, posBack)
    If pos > 0 Then
        Let GetShortName = VBA.Strings.Mid$(filename, pos + 1)
    Else
        Let GetShortName = filename
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
        Call Error.FAIL(ECODE.Err_FileNotFound, "Utils_Test.RenameFile")
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
    , ByRef targetFile As String _
) As Boolean
    If Not IsExistFile(sourceFile) Then
        Call Error.FAIL(ECODE.Err_FileNotFound, "Utils_Test.CopyFileSmall")
        Exit Function
    End If
    Dim targetFolder As String: Let targetFolder = Me.GetParentFolder(targetFile)
    If Not Me.IsExistFolder(targetFolder) Then Call Me.CreateFolder(targetFolder)
    If Me.IsExistFile(targetFile) Then Call VBA.fileSystem.Kill(targetFile)
    Call VBA.fileSystem.FileCopy(sourceFile, targetFile)
    Let CopyFileSmall = Me.IsExistFile(targetFile)
End Function

Friend Sub CopyFile(ByRef sourceFile As String, ByRef targetFile As String)
    Const BLOCKSIZE As Long = 65536 ' 64KB
    If Not IsExistFile(sourceFile) Then
        Call Error.FAIL(ECODE.Err_FileNotFound, "Utils_Test.CopyFile")
        Exit Sub
    End If
    Dim targetFolder As String: Let targetFolder = Me.GetParentFolder(targetFile)
    If Not Me.IsExistFolder(targetFolder) Then Call Me.CreateFolder(targetFolder)
    If Me.IsExistFile(targetFile) Then Call VBA.fileSystem.Kill(targetFile)
    Dim sourceNum As Long: Let sourceNum = FreeFile
    Open sourceFile For Binary Access Read As #sourceNum
    Dim targetNum As Long: Let targetNum = FreeFile
    Open targetFile For Binary Access Write As #targetNum
    Dim bytesRead As Long
    Dim buffer() As Byte
    Do While Not eof(sourceNum)
        ReDim buffer(1 To BLOCKSIZE)
        Let bytesRead = LOF(sourceNum) - Loc(sourceNum) + 1
        If bytesRead > BLOCKSIZE Then bytesRead = BLOCKSIZE
        Get #sourceNum, , buffer
        Put #targetNum, , buffer
    Loop
    Close #sourceNum
    Close #targetNum
End Sub

Friend Sub CopyFolder(ByRef sourceFolder As String, ByRef targetFolder As String)
    Stop
    If Not Me.IsExistFolder(sourceFolder) Then
        Call Error.FAIL(ECODE.Err_FolderNotFound, "Utils_Test.CopyFolder")
        Exit Sub
    End If
    Dim Scripts As Utils_Scripts: Set Scripts = New Utils_Scripts
    Dim fileSystem As Object: Set fileSystem = Scripts.fileSystem
    If Me.IsExistFolder(targetFolder) Then Call Me.DeleteFolder(targetFolder)
    Call fileSystem.CopyFolder(sourceFolder, targetFolder)
End Sub

Friend Sub DeleteFolder(ByRef folderPath As String)
    If Not Me.IsExistFolder(folderPath) Then Exit Sub
    Dim Scripts As Utils_Scripts: Set Scripts = New Utils_Scripts
    Dim fileSystem As Object: Set fileSystem = Scripts.fileSystem
    Call fileSystem.DeleteFolder(folderPath, True)
End Sub

Friend Sub ExtractExcelByZip(ByVal filename As String, ByVal targetFolder As String)
    If Not Me.IsExistFile(filename) Then
        Call Error.FAIL(ECODE.Err_FileNotFound, "Utils_Test.ExtractExcelByZip")
        Exit Sub
    End If
    Const TIMEOUT As Byte = 30 's
    If Not Me.IsExistFolder(targetFolder) Then Call Me.CreateFolder(targetFolder)
    Dim tempFolder As String: tempFolder = targetFolder & "\_temp"
    If Not Me.IsExistFolder(tempFolder) Then Call Me.CreateFolder(tempFolder)
    Dim tempZip As String: tempZip = tempFolder & "\file.zip"
    Call Me.CopyFile(filename, tempZip)
    Dim Scripts As Utils_Scripts: Set Scripts = New Utils_Scripts
    Dim shell As Object: Set shell = Scripts.AShell
    Dim zipFolder As Object: Set zipFolder = shell.Namespace(tempZip)
    Dim destFolder As Object: Set destFolder = shell.Namespace(tempFolder)
    If Not zipFolder Is Nothing Then
        Dim item As Object: For Each item In zipFolder.Items
            Call destFolder.CopyHere(item)
        Next item
        Dim t As Double: Let t = Timer
        Dim expectedCount As Long: Let expectedCount = zipFolder.Items.count
        Dim actualCount As Long
        Do
            actualCount = destFolder.Items.count
            DoEvents
            If Timer - t > TIMEOUT Then Exit Do
        Loop While actualCount < expectedCount
    End If
    Call Me.CopyFolder(tempFolder, targetFolder)
    Call Me.DeleteFolder(tempFolder)
End Sub

Friend Function diffCheck( _
    ByRef fileName1 As String _
    , ByRef fileName2 As String _
) As Boolean
    Const BLOCKSIZE As Long = 65536 ' 64KB
    Dim Scripts1 As Utils_Scripts: Set Scripts1 = New Utils_Scripts
    Dim Scripts2 As Utils_Scripts: Set Scripts2 = New Utils_Scripts
    Dim file1 As Object: Set file1 = Scripts1.fileStream
    Let file1.Type = 1
    Call file1.Open
    Call file1.LoadFromFile(fileName1)
    Dim file2 As Object: Set file2 = Scripts2.fileStream
    Let file2.Type = 1
    Call file2.Open
    Call file2.LoadFromFile(fileName2)
    Do While Not file1.EOS Or Not file2.EOS
        Dim block1() As Byte: Let block1 = file1.Read(BLOCKSIZE)
        Dim block2() As Byte: Let block2 = file2.Read(BLOCKSIZE)
        Dim bytesRead1 As Long: Let bytesRead1 = UBound(block1) - LBound(block1) + 1
        Dim bytesRead2 As Long: Let bytesRead2 = UBound(block2) - LBound(block2) + 1
        If bytesRead1 <> bytesRead2 Then
            Let diffCheck = False
            GoTo Cleanup
        End If
        Dim i As Long: For i = LBound(block1) To UBound(block1)
            If block1(i) <> block2(i) Then
                Let diffCheck = False
                GoTo Cleanup
            End If
        Next i
    Loop
    Let diffCheck = True
Cleanup:
    Call file1.Close
    Call file2.Close
    Set file1 = Nothing
    Set Scripts1 = Nothing
    Set file2 = Nothing
    Set Scripts2 = Nothing
End Function

Friend Function GetEndOfFile(ByRef filename As String) As String
    Dim Scripts As Utils_Scripts: Set Scripts = New Utils_Scripts
    Dim shell As Object: Set shell = Scripts.wShell
    ' If Not IsExistFile(fileName) Then
    '     Let GetEndOfFile = "File Not Found!"
    '     Goto CleanUp
    ' End If
    Const NKF_CONST As String = "D:\share\ASS\Tasks\template\autotest.exe"
    Const GUESS As String = " --guess "
    Dim NKF As String: Let NKF = vbNullString 'Me.Where("nkf")
    If NKF = vbNullString Then Let NKF = NKF_CONST
    If Not IsExistFile(NKF) Then
        Let GetEndOfFile = "NKF Not Found!"
        GoTo Cleanup
    End If
    Dim command As String: Let command = """" & NKF & """" & GUESS & """" & filename & """"
    Dim exec As Object: Set exec = shell.exec(command)
    Dim result As String: Let result = exec.StdOut.ReadAll
    If result = vbNullString Then Let result = "File not found!"
    Let GetEndOfFile = result
Cleanup:
    Set shell = Nothing
    Set Scripts = Nothing
End Function

' TODO DANG bi loi encodign file
Friend Function GetEndOfFiles(ByRef fileNames() As String) As Variant
    Const NKF_CONST As String = "D:\share\ASS\Tasks\template\autotest.exe"
    Const GUESS As String = " --guess "
    Const LIMIT As Byte = 10
    Dim NKF As String: Let NKF = NKF_CONST
    Dim results() As String: ReDim results(LBound(fileNames) To UBound(fileNames))
    Dim Scripts As Utils_Scripts: Set Scripts = New Utils_Scripts
    Dim shell As Object: Set shell = Scripts.wShell
    Dim dict As Object: Set dict = Scripts.Dictionary
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
    Set Scripts = Nothing
End Function

Friend Function ReadLastLines( _
        ByRef filename As String _
        , ByRef lineCount As Long _
        , Optional charset As String = "utf-8" _
    ) As Variant
    Const BLOCKSIZE As Long = 65536 ' 64KB
    Dim Scripts As Utils_Scripts: Set Scripts = New Utils_Scripts
    Dim fileStream As Object: Set fileStream = Scripts.fileStream
    Let fileStream.Type = 2 ' Text
    Let fileStream.charset = charset ' Unix override to 'x-euc-jp'
    Call fileStream.Open
    Call fileStream.LoadFromFile(filename)
    Dim fileSize As Long: Let fileSize = fileStream.size
    Dim position As Long: Let position = IIf(fileSize - BLOCKSIZE > 0, fileSize - BLOCKSIZE, 0)
    Let fileStream.position = position
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
        Dim Scripts As Utils_Scripts: Set Scripts = New Utils_Scripts
        Set connection = Scripts.connection
        Call connection.Open( _
            "Provider=Microsoft.ACE.OLEDB.12.0;" & _
            "Data Source=" & fullName & ";" & _
            "Extended Properties=""Excel 12.0 Xml;HDR=No;IMEX=1"";" _
        )
        Let isTempConnection = True
    End If
    Dim sheetNameLow As String: sheetNameLow = LCase(sheetName & "$")
    Dim Recordset As Object: Set Recordset = connection.OpenSchema(SCHEMA.adSchemaTables)
    Do Until Recordset.eof
        Dim tableName As String: Let tableName = LCase(Recordset.Fields("TABLE_NAME").value)
        If tableName = sheetNameLow Or tableName = "'" & sheetNameLow & "'" Then
            Let IsSheetExistInExcel = True
            GoTo Cleanup
        End If
        Recordset.MoveNext
    Loop
    Let IsSheetExistInExcel = False
Cleanup:
    If isTempConnection Then
        If connection.State = OBJECT_STATE.adStateOpen Then connection.Close
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
    Dim Scripts As Utils_Scripts: Set Scripts = New Utils_Scripts
    Dim connection As Object: Set connection = Scripts.connection
    Dim Recordset As Object: Set Recordset = Scripts.Recordset
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
        , CURSOR_TYPE.adOpenStatic _
        , LOCK_TYPE.adLockReadOnly _
    )
    If Recordset.BOF And Recordset.eof Then GoTo Cleanup
    Dim isStart As Boolean: Let isStart = True
    If startPattern <> vbNullString Then Let isStart = False
    If isRevert Then
        Call Recordset.MoveLast
        Do While Not Recordset.BOF
            With Recordset.Fields(targetCol)
                If stopPattern <> vbNullString And InStr(.value, stopPattern) > 0 Then Exit Do
                If Not isStart Then
                    If InStr(.value, startPattern) > 0 Then Let isStart = True
                ElseIf InStr(.value, pattern) > 0 Then
                    Let result = .value
                    Exit Do
                End If
            End With
            Call Recordset.MovePrevious
        Loop
    Else
        Call Recordset.MoveFirst
        Do While Not Recordset.eof
            With Recordset.Fields(targetCol)
                If stopPattern <> vbNullString And InStr(.value, stopPattern) > 0 Then Exit Do
                If Not isStart Then
                    If InStr(.value, startPattern) > 0 Then Let isStart = True
                ElseIf InStr(.value, pattern) > 0 Then
                    Let result = .value
                    Exit Do
                End If
            End With
            Call Recordset.MoveNext
        Loop
    End If
Cleanup:
    If Not Recordset Is Nothing Then
        If Recordset.State = OBJECT_STATE.adStateOpen Then Call Recordset.Close
        Set Recordset = Nothing
    End If
    If Not connection Is Nothing Then
        If connection.State = adStateOpen Then Call connection.Close
        Set connection = Nothing
    End If
    Let SearchInExcelByPattern = result
End Function

'HELPER
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
    Dim Scripts As Utils_Scripts: Set Scripts = New Utils_Scripts
    Dim regex As Object: Set regex = Scripts.RegexExpression
    Let regex.pattern = pattern
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
    ByRef filename As String _
    , ByRef pattern As String _
    , Optional ByRef lineAmount As Long = 0 _
    , Optional ByRef isRevert As Boolean = False _
    , Optional ByRef charset As String = "utf-8" _
) As Collection
    Const BLOCKSIZE As Long = 65536 ' 64KB
    Dim result As New Collection
    Dim Scripts As Utils_Scripts: Set Scripts = New Utils_Scripts
    Dim fileStream As Object: Set fileStream = Scripts.fileStream
    Let fileStream.Type = 2
    Let fileStream.charset = charset
    Call fileStream.Open
    Call fileStream.LoadFromFile(filename)
    Dim fileSize As Long: Let fileSize = fileStream.size
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
            Let fileStream.position = position
            Let chunk = fileStream.ReadText(IIf(tmpPosition - position, tmpPosition - position, BLOCKSIZE))
            Let chunk = Replace(chunk, vbCrLf, vbLf)
            Let lines = Split(chunk, vbLf)
            For i = UBound(lines) To LBound(lines) Step -1
                If InStr(lines(i), pattern) > 0 Then
                    Call result.Add(lines(i))
                    If lineAmount > 0 And result.count >= lineAmount Then Exit Do
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
                    If lineAmount > 0 And result.count >= lineAmount Then Exit Do
                End If
            Next i
        Loop
    End If
    Call fileStream.Close
    Set fileStream = Nothing
    Set SearchInFileWithPattern = result
End Function


