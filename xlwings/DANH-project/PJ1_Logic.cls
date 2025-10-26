VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PJ1_Logic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Check README.md for more information
Option Explicit
' --- Dependencies ---
Private Test As Utils_Test
Private Addr As Utils_Address
Private Stringify As Utils_Stringify
Private Scripts As Utils_Scripts
' Private Info As InfoConstants
' Private System As SystemUpdate
' --- Object ---
Private userResponse As VbMsgBoxResult
' --- Constant ---
' --- Constructor ---
Friend Sub Init( _
    Optional ByRef targetWorksheet As Worksheet _
)
    If targetWorksheet Is Nothing Then
        Set targetWorksheet = Excel.Application.ActiveSheet
    End If
    Set Test = New Utils_Test
    Set Addr = New Utils_Address
    Set Stringify = New Utils_Stringify
    Set Scripts = New Utils_Scripts
    Call Addr.Init(targetWorksheet, New PJ1_Address)
End Sub
' --- Destructor ---
Private Sub Class_Terminate()
    Set Test = Nothing
    Set Addr = Nothing
    Set Stringify = Nothing
    Set Scripts = Nothing
End Sub

' --- Fuctions / Subroutines ---
Friend Sub ChecklistExits()
    Dim dict As Object: Set dict = Scripts.Dictionary
    Dim tasks() As String: Let tasks = Addr.GetValueArray("TASKS")
    Dim results As Collection: Set results = New Collection
    Dim checklistNames As Collection: Set checklistNames = New Collection
    Dim utPattern As String: Let utPattern = Addr.GetValue("FILE_CHECKLIST_PATTERN")
    Dim i As Long: For i = LBound(tasks) To UBound(tasks)
        Dim taskCode As String: Let taskCode = left$(tasks(i), 8)
        Dim taskName As String: Let taskName = Right$(tasks(i), Len(tasks(i)) - 9)
        Let dict("CODE") = taskCode
        Let dict("NAME") = taskName
        Dim checkListName As String: Let checkListName = Stringify.ReplacePlaceholder( _
            utPattern _
            , dict _
        )
        Dim checkListPath As String: Let checkListPath = Addr.GetValue("CHECKLIST_PATH")
        Dim checkListFullName As String: Let checkListFullName = checkListPath & checkListName
        Call checklistNames.Add(checkListFullName)
        Call results.Add(IIf( _
            Test.IsExistPath(checkListFullName) _
            , Test.PASS _
            , Test.FAIL _
        ))
    Next i
    Call Addr.SetValueArray("CHECKLIST_FULLNAME", CollectionToArray(checklistNames))
    Call Addr.SetValueArray("RESULT_CHECKLIST_EXITS", CollectionToArray(results))
End Sub

Friend Sub UTExits()
    Dim dict As Object: Set dict = Scripts.Dictionary
    Dim tasks() As String: Let tasks = Addr.GetValueArray("TASKS")
    Dim results As Collection: Set results = New Collection
    Dim utNames As Collection: Set utNames = New Collection
    Dim utPattern As String: Let utPattern = Addr.GetValue("FILE_UT_PATTERN")
    Dim i As Long: For i = LBound(tasks) To UBound(tasks)
        Dim taskCode As String: Let taskCode = left$(tasks(i), 8)
        Dim taskName As String: Let taskName = Right$(tasks(i), Len(tasks(i)) - 9)
        Let dict("CODE") = taskCode
        Let dict("NAME") = taskName
        Dim utName As String: Let utName = Stringify.ReplacePlaceholder( _
            utPattern _
            , dict _
        )
        Dim utPath As String: Let utPath = Addr.GetValue("UT_PATH")
        Dim utFullPath As String: Let utFullPath = utPath & taskCode & "_" & taskName & "\"
        Dim utFullName As String: Let utFullName = utFullPath & utName
        Call utNames.Add(utFullName)
        Call results.Add(IIf( _
            Test.IsExistPath(utFullName) _
            , Test.PASS _
            , Test.FAIL _
        ))
    Next i
    Call Addr.SetValueArray("UT_FULLNAME", CollectionToArray(utNames))
    Call Addr.SetValueArray("RESULT_UT_EXITS", CollectionToArray(results))
End Sub

Friend Sub CountInUT()
    Dim tasks() As String: Let tasks = Addr.GetValueArray("TASKS")
    Dim results As Collection: Set results = New Collection
    Dim i As Long: For i = LBound(tasks) To UBound(tasks)
        Dim taskCode As String: Let taskCode = left$(tasks(i), 8)
        Dim taskName As String: Let taskName = Right$(tasks(i), Len(tasks(i)) - 9)
        Dim utPath As String: Let utPath = Addr.GetValue("UT_PATH")
        Dim utFullPath As String: Let utFullPath = utPath & taskCode & "_" & taskName & "\"
        Call results.Add(CStr(Test.CountInFolder(utFullPath)))
    Next i
    Call Addr.SetValueArray("RESULT_COUNT_IN_UT", CollectionToArray(results))
End Sub

Friend Sub CountInOutput()
    Dim tasks() As String: Let tasks = Addr.GetValueArray("TASKS")
    Dim results As Collection: Set results = New Collection
    Dim i As Long: For i = LBound(tasks) To UBound(tasks)
        Dim taskCode As String: Let taskCode = left$(tasks(i), 8)
        Dim taskName As String: Let taskName = Right$(tasks(i), Len(tasks(i)) - 9)
        Dim utPath As String: Let utPath = Addr.GetValue("UT_PATH")
        Dim outputPath As String: Let outputPath = utPath & taskCode & "_" & taskName & "\OUTPUT\"
        Call results.Add(CStr(Test.CountInFolder(outputPath)))
    Next i
    Call Addr.SetValueArray("RESULT_COUNT_IN_OUTPUT", CollectionToArray(results))
End Sub

Friend Sub SfcmaplgExist()
    Dim dict As Object: Set dict = Scripts.Dictionary
    Dim tasks() As String: Let tasks = Addr.GetValueArray("TASKS")
    Dim results As Collection: Set results = New Collection
    Dim sfcmaplgNames As Collection: Set sfcmaplgNames = New Collection
    Dim i As Long: For i = LBound(tasks) To UBound(tasks)
        Dim taskCode As String: Let taskCode = left$(tasks(i), 8)
        Dim taskName As String: Let taskName = Right$(tasks(i), Len(tasks(i)) - 9)
        Let dict("CODE") = taskCode
        Let dict("NAME") = taskName
        Dim utPath As String: Let utPath = Addr.GetValue("UT_PATH")
        Dim outputPath As String: Let outputPath = utPath & taskCode & "_" & taskName & "\OUTPUT\"
        Dim sfcmaplgFullName As String: Let sfcmaplgFullName = outputPath & "sfcmaplg"
        Call sfcmaplgNames.Add(sfcmaplgFullName)
        Call results.Add(IIf( _
            Test.IsExistPath(sfcmaplgFullName) _
            , Test.PASS _
            , Test.FAIL _
        ))
    Next i
    Call Addr.SetValueArray("SFCMAPLG_FULLNAME", CollectionToArray(sfcmaplgNames))
    Call Addr.SetValueArray("RESULT_SFCMAPLG_EXITS", CollectionToArray(results))
End Sub

Friend Sub SfcmerlgExist()
    Dim dict As Object: Set dict = Scripts.Dictionary
    Dim tasks() As String: Let tasks = Addr.GetValueArray("TASKS")
    Dim results As Collection: Set results = New Collection
    Dim sfcmerlgNames As Collection: Set sfcmerlgNames = New Collection
    Dim i As Long: For i = LBound(tasks) To UBound(tasks)
        Dim taskCode As String: Let taskCode = left$(tasks(i), 8)
        Dim taskName As String: Let taskName = Right$(tasks(i), Len(tasks(i)) - 9)
        Let dict("CODE") = taskCode
        Let dict("NAME") = taskName
        Dim utPath As String: Let utPath = Addr.GetValue("UT_PATH")
        Dim outputPath As String: Let outputPath = utPath & taskCode & "_" & taskName & "\OUTPUT\"
        Dim sfcmerlgFullName As String: Let sfcmerlgFullName = outputPath & "sfcmerlg"
        Call sfcmerlgNames.Add(sfcmerlgFullName)
        Call results.Add(IIf( _
            Test.IsExistPath(sfcmerlgFullName) _
            , Test.PASS _
            , Test.FAIL _
        ))
    Next i
    Call Addr.SetValueArray("SFCMERLG_FULLNAME", CollectionToArray(sfcmerlgNames))
    Call Addr.SetValueArray("RESULT_SFCMERLG_EXITS", CollectionToArray(results))
End Sub

Friend Sub LogExist()
    Dim dict As Object: Set dict = Scripts.Dictionary
    Dim tasks() As String: Let tasks = Addr.GetValueArray("TASKS")
    Dim results As Collection: Set results = New Collection
    Dim logNames As Collection: Set logNames = New Collection
    Dim i As Long: For i = LBound(tasks) To UBound(tasks)
        Dim taskCode As String: Let taskCode = left$(tasks(i), 8)
        Dim taskName As String: Let taskName = Right$(tasks(i), Len(tasks(i)) - 9)
        Let dict("CODE") = taskCode
        Let dict("NAME") = taskName
        Dim utPath As String: Let utPath = Addr.GetValue("UT_PATH")
        Dim outputPath As String: Let outputPath = utPath & taskCode & "_" & taskName & "\OUTPUT\"
        Dim logFullName As String: Let logFullName = outputPath & taskCode & ".log"
        Call logNames.Add(logFullName)
        Call results.Add(IIf( _
            Test.IsExistPath(logFullName) _
            , Test.PASS _
            , Test.FAIL _
        ))
    Next i
    Call Addr.SetValueArray("LOG_FULLNAME", CollectionToArray(logNames))
    Call Addr.SetValueArray("RESULT_LOG_EXITS", CollectionToArray(results))
End Sub

' TODO: fix LATER
' Friend Sub LogEOF()
'     Dim dict As Object: Set dict = Scripts.Dictionary
'     Dim tasks() As String: Let tasks = Addr.GetValueArray("TASKS")
'     Dim names As Collection: Set names = New Collection
'     Dim logNames As Collection: Set logNames = New Collection
'     Dim i As Long: For i = LBound(tasks) To UBound(tasks)
'         Dim taskCode As String: Let taskCode = Left$(tasks(i), 8)
'         Dim taskName As String: Let taskName = Right$(tasks(i), len(tasks(i)) - 9)
'         Let dict("CODE") = taskCode
'         Let dict("NAME") = taskName
'         Dim utPath As String: Let utPath = Addr.GetValue("UT_PATH")
'         Dim outputPath As String: Let outputPath = utPath & taskCode & "_" & taskName & "\OUTPUT\"
'         Dim logFullName As String: Let logFullName = outputPath & taskCode & ".log"
'         Call names.Add(logFullName)
'     Next i
'     Dim results() As String: Let results = Test.GetEndOfFiles(CollectionToArray(names))
'     Call Addr.SetValueArray("RESULT_LOG_EOF", results)
' End Sub

Friend Sub LogEOF()
    Dim dict As Object: Set dict = Scripts.Dictionary
    Dim tasks() As String: Let tasks = Addr.GetValueArray("TASKS")
    Dim eof As Collection: Set eof = New Collection
    Dim i As Long: For i = LBound(tasks) To UBound(tasks)
        Dim taskCode As String: Let taskCode = left$(tasks(i), 8)
        Dim taskName As String: Let taskName = Right$(tasks(i), Len(tasks(i)) - 9)
        Let dict("CODE") = taskCode
        Let dict("NAME") = taskName
        Dim utPath As String: Let utPath = Addr.GetValue("UT_PATH")
        Dim outputPath As String: Let outputPath = utPath & taskCode & "_" & taskName & "\OUTPUT\"
        Dim logFullName As String: Let logFullName = outputPath & taskCode & ".log"
        Call eof.Add(Test.GetEndOfFile(logFullName))
    Next i
    Call Addr.SetValueArray("RESULT_LOG_EOF", CollectionToArray(eof))
End Sub

Friend Sub SfcmaplgEOF()
    Dim dict As Object: Set dict = Scripts.Dictionary
    Dim tasks() As String: Let tasks = Addr.GetValueArray("TASKS")
    Dim eof As Collection: Set eof = New Collection
    Dim i As Long: For i = LBound(tasks) To UBound(tasks)
        Dim taskCode As String: Let taskCode = left$(tasks(i), 8)
        Dim taskName As String: Let taskName = Right$(tasks(i), Len(tasks(i)) - 9)
        Let dict("CODE") = taskCode
        Let dict("NAME") = taskName
        Dim utPath As String: Let utPath = Addr.GetValue("UT_PATH")
        Dim outputPath As String: Let outputPath = utPath & taskCode & "_" & taskName & "\OUTPUT\"
        Dim logFullName As String: Let logFullName = outputPath & "sfcmaplg"
        Call eof.Add(Test.GetEndOfFile(logFullName))
    Next i
    Call Addr.SetValueArray("RESULT_SFCMAPLG_EOF", CollectionToArray(eof))
End Sub

Friend Sub SfcmerlgEOF()
    Dim dict As Object: Set dict = Scripts.Dictionary
    Dim tasks() As String: Let tasks = Addr.GetValueArray("TASKS")
    Dim eof As Collection: Set eof = New Collection
    Dim i As Long: For i = LBound(tasks) To UBound(tasks)
        Dim taskCode As String: Let taskCode = left$(tasks(i), 8)
        Dim taskName As String: Let taskName = Right$(tasks(i), Len(tasks(i)) - 9)
        Let dict("CODE") = taskCode
        Let dict("NAME") = taskName
        Dim utPath As String: Let utPath = Addr.GetValue("UT_PATH")
        Dim outputPath As String: Let outputPath = utPath & taskCode & "_" & taskName & "\OUTPUT\"
        Dim logFullName As String: Let logFullName = outputPath & "sfcmerlg"
        Call eof.Add(Test.GetEndOfFile(logFullName))
    Next i
    Call Addr.SetValueArray("RESULT_SFCMERLG_EOF", CollectionToArray(eof))
End Sub

Friend Sub CheckLogContent()
    Dim dict As Object: Set dict = Scripts.Dictionary
    Dim tasks() As String: Let tasks = Addr.GetValueArray("TASKS")
    Dim lastLines As Collection: Set lastLines = New Collection
    Dim results As Collection: Set results = New Collection
    Dim i As Long: For i = LBound(tasks) To UBound(tasks)
        Dim taskCode As String: Let taskCode = left$(tasks(i), 8)
        Dim taskName As String: Let taskName = Right$(tasks(i), Len(tasks(i)) - 9)
        Let dict("CODE") = taskCode
        Let dict("NAME") = taskName
        Dim utPath As String: Let utPath = Addr.GetValue("UT_PATH")
        Dim outputPath As String: Let outputPath = utPath & taskCode & "_" & taskName & "\OUTPUT\"
        Dim logFullName As String: Let logFullName = outputPath & taskCode & ".log"
        If Test.IsExistFile(logFullName) Then
            Dim lines() As String: Let lines = Test.ReadLastLines(logFullName, 1, "x-euc-jp")
            Dim j As Long: For j = LBound(lines) To UBound(lines)
                Call lastLines.Add(lines(j))
            Next j
            Call results.Add(IIf( _
                lines(UBound(lines)) = Addr.GetValue("LOG_CONTENT_PATTERN") _
                , Test.PASS _
                , Test.FAIL _
            ))
        Else
            Call lastLines.Add("File Not Found!")
            Call results.Add(Test.FAIL)
        End If
    Next i
    Call Addr.SetValueArray("LOG_CONTENT", CollectionToArray(lastLines))
    Call Addr.SetValueArray("RESULT_LOG_CONTENT", CollectionToArray(results))
End Sub

Friend Sub LastUpdate()
    Dim dict As Object: Set dict = Scripts.Dictionary
    Dim tasks() As String: Let tasks = Addr.GetValueArray("TASKS")
    Dim lastLines As Collection: Set lastLines = New Collection
    Dim results As Collection: Set results = New Collection
    Dim i As Long: For i = LBound(tasks) To UBound(tasks)
        Dim taskCode As String: Let taskCode = left$(tasks(i), 8)
        Dim taskName As String: Let taskName = Right$(tasks(i), Len(tasks(i)) - 9)
        Let dict("CODE") = taskCode
        Let dict("NAME") = taskName
        Dim utPath As String: Let utPath = Addr.GetValue("UT_PATH")
        Dim outputPath As String: Let outputPath = utPath & taskCode & "_" & taskName & "\OUTPUT\"
        Dim logFullName As String: Let logFullName = outputPath & taskCode & ".log"
        If Test.IsExistFile(logFullName) Then
            Dim lines() As String: Let lines = Test.ReadLastLines(logFullName, 1, "x-euc-jp")
            Dim j As Long: For j = LBound(lines) To UBound(lines)
                Call lastLines.Add(lines(j))
            Next j
            Call results.Add(IIf( _
                lines(UBound(lines)) = Addr.GetValue("LOG_CONTENT_PATTERN") _
                , Test.PASS _
                , Test.FAIL _
            ))
        Else
            Call lastLines.Add("File Not Found!")
            Call results.Add(Test.FAIL)
        End If
    Next i
    Call Addr.SetValueArray("LOG_CONTENT", CollectionToArray(lastLines))
    Call Addr.SetValueArray("RESULT_LOG_CONTENT", CollectionToArray(results))
End Sub

Friend Sub LastRuntime()
    Dim dict As Object: Set dict = Scripts.Dictionary
    Dim tasks() As String: Let tasks = Addr.GetValueArray("TASKS")
    Dim result As String
    Dim lastlogs As Collection: Set lastlogs = New Collection
    Dim lasttimes As Collection: Set lasttimes = New Collection
    Dim lastPids As Collection: Set lastPids = New Collection
    Dim utNames As Collection: Set utNames = New Collection
    Dim utPattern As String: Let utPattern = Addr.GetValue("FILE_UT_PATTERN")
    Dim i As Long: For i = LBound(tasks) To UBound(tasks)
        Dim taskCode As String: Let taskCode = left$(tasks(i), 8)
        Dim taskName As String: Let taskName = Right$(tasks(i), Len(tasks(i)) - 9)
        Let dict("CODE") = taskCode
        Let dict("NAME") = taskName
        Dim utName As String: Let utName = Stringify.ReplacePlaceholder( _
            utPattern _
            , dict _
        )
        Dim utPath As String: Let utPath = Addr.GetValue("UT_PATH")
        Dim utFullPath As String: Let utFullPath = utPath & taskCode & "_" & taskName & "\"
        Dim utFullName As String: Let utFullName = utFullPath & utName
        Dim utSheetName As String: Let utSheetName = Addr.GetValue("UT_SHEET_NAME")
        Dim startPattern As String: Let startPattern = Addr.GetValue("SFCMERLG_PATTERN")
        Dim stopPattern As String: Let stopPattern = Addr.GetValue("SFCMAPLG_PATTERN")
        Dim targetCol As Long: Let targetCol = Test.ColumnNameToIndex("A")
        Dim isRevert As Boolean: Let isRevert = True
        Dim pidPattern As String: Let pidPattern = "\[(\d+)\]"
        If Not Test.IsExistPath(utFullName) Then
            Call lastlogs.Add("File Not Found!")
            Call lasttimes.Add("File Not Found!")
            Call lastPids.Add("File Not Found!")
        Else
            Let result = Test.SearchInExcelByPattern( _
                utFullName _
                , utSheetName _
                , targetCol _
                , taskCode _
                , isRevert _
                , startPattern _
                , stopPattern _
            )
            Call lastlogs.Add(result)
            Call lasttimes.Add(left(result, 15))
            Call lastPids.Add(Test.ExtracByRegex(result, pidPattern))
        End If
    Next i
    Call Addr.SetValueArray("LAST_RUNLOG", CollectionToArray(lastlogs))
    Call Addr.SetValueArray("LAST_RUNTIME", CollectionToArray(lasttimes))
    Call Addr.SetValueArray("LAST_PIDS", CollectionToArray(lastPids))
End Sub

Friend Sub GetTestcsh()
    Dim dict As Object: Set dict = Scripts.Dictionary
    Dim tasks() As String: Let tasks = Addr.GetValueArray("TASKS")
    Dim utPattern As String: Let utPattern = Addr.GetValue("FILE_UT_PATTERN")
    Dim testcshCmd As Collection: Set testcshCmd = New Collection
    Dim result As String
    Dim i As Long: For i = LBound(tasks) To UBound(tasks)
        Dim taskCode As String: Let taskCode = left$(tasks(i), 8)
        Dim taskName As String: Let taskName = Right$(tasks(i), Len(tasks(i)) - 9)
        Let dict("CODE") = taskCode
        Let dict("NAME") = taskName
        Dim utName As String: Let utName = Stringify.ReplacePlaceholder( _
            utPattern _
            , dict _
        )
        Dim utPath As String: Let utPath = Addr.GetValue("UT_PATH")
        Dim utFullPath As String: Let utFullPath = utPath & taskCode & "_" & taskName & "\"
        Dim utFullName As String: Let utFullName = utFullPath & utName
        Dim utSheetName As String: Let utSheetName = Addr.GetValue("UT_SHEET_NAME")
        Dim startPattern As String: Let startPattern = Addr.GetValue("EXEC_CMD_PATTERN")
        Dim stopPattern As String: Let stopPattern = Addr.GetValue("EXEC_LOG_PATTERN")
        Dim targetCol As Long: Let targetCol = Test.ColumnNameToIndex("A")
        Dim isRevert As Boolean: Let isRevert = False
        ' Dim pidPattern As String: Let pidPattern = "\[(\d+)\]"
        If Not Test.IsExistPath(utFullName) Then
            Call testcshCmd.Add("File Not Found!")
        Else
            Let result = Test.SearchInExcelByPattern( _
                utFullName _
                , utSheetName _
                , targetCol _
                , taskCode _
                , isRevert _
                , startPattern _
                , stopPattern _
            )
            Call testcshCmd.Add(result)
        End If
    Next i
    Call Addr.SetValueArray("TESTCSH_COMAND", CollectionToArray(testcshCmd))
End Sub

Friend Sub GetTestcshExec()
    Dim dict As Object: Set dict = Scripts.Dictionary
    Dim tasks() As String: Let tasks = Addr.GetValueArray("TASKS")
    Dim testcshCmds() As String: Let testcshCmds = Addr.GetValueArray("TESTCSH_COMAND")
    Dim utPattern As String: Let utPattern = Addr.GetValue("FILE_UT_PATTERN")
    Dim testcshCmdExec As Collection: Set testcshCmdExec = New Collection
    Dim result As String
    Dim i As Long: For i = LBound(tasks) To UBound(tasks)
        Dim taskCode As String: Let taskCode = left$(tasks(i), 8)
        Dim taskName As String: Let taskName = Right$(tasks(i), Len(tasks(i)) - 9)
        Let dict("CODE") = taskCode
        Let dict("NAME") = taskName
        Dim utName As String: Let utName = Stringify.ReplacePlaceholder( _
            utPattern _
            , dict _
        )
        Dim utPath As String: Let utPath = Addr.GetValue("UT_PATH")
        Dim utFullPath As String: Let utFullPath = utPath & taskCode & "_" & taskName & "\"
        Dim utFullName As String: Let utFullName = utFullPath & utName
        Dim utSheetName As String: Let utSheetName = Addr.GetValue("UT_SHEET_NAME")
        Dim startPattern As String: Let startPattern = Addr.GetValue("EXEC_LOG_PATTERN")
        Dim stopPattern As String: Let stopPattern = Addr.GetValue("SFCMAPLG_PATTERN")
        Dim targetCol As Long: Let targetCol = Test.ColumnNameToIndex("A")
        Dim isRevert As Boolean: Let isRevert = False
        Dim testcshCmd As String: Let testcshCmd = testcshCmds(i)
        If Not Test.IsExistPath(utFullName) Then
            Call testcshCmdExec.Add("File Not Found!")
        Else
            Let result = Test.SearchInExcelByPattern( _
                utFullName _
                , utSheetName _
                , targetCol _
                , testcshCmd _
                , isRevert _
                , startPattern _
                , stopPattern _
            )
            Call testcshCmdExec.Add(result)
        End If
    Next i
    Call Addr.SetValueArray("TESTCSH_COMAND_EXEC", CollectionToArray(testcshCmdExec))
End Sub

Friend Sub GetResultExec()
    Dim dict As Object: Set dict = Scripts.Dictionary
    Dim tasks() As String: Let tasks = Addr.GetValueArray("TASKS")
    Dim testcshCmds() As String: Let testcshCmds = Addr.GetValueArray("TESTCSH_COMAND")
    Dim utPattern As String: Let utPattern = Addr.GetValue("FILE_UT_PATTERN")
    Dim testcshCmdExec As Collection: Set testcshCmdExec = New Collection
    Dim result As String
    Dim i As Long: For i = LBound(tasks) To UBound(tasks)
        Dim taskCode As String: Let taskCode = left$(tasks(i), 8)
        Dim taskName As String: Let taskName = Right$(tasks(i), Len(tasks(i)) - 9)
        Let dict("CODE") = taskCode
        Let dict("NAME") = taskName
        Dim utName As String: Let utName = Stringify.ReplacePlaceholder( _
            utPattern _
            , dict _
        )
        Dim utPath As String: Let utPath = Addr.GetValue("UT_PATH")
        Dim utFullPath As String: Let utFullPath = utPath & taskCode & "_" & taskName & "\"
        Dim utFullName As String: Let utFullName = utFullPath & utName
        Dim utSheetName As String: Let utSheetName = Addr.GetValue("UT_SHEET_NAME")
        Dim startPattern As String: Let startPattern = Addr.GetValue("EXEC_LOG_PATTERN")
        Dim stopPattern As String: Let stopPattern = Addr.GetValue("SFCMERLG_PATTERN")
        Dim targetCol As Long: Let targetCol = Test.ColumnNameToIndex("A")
        Dim isRevert As Boolean: Let isRevert = False
        Dim testcshCmd As String: Let testcshCmd = testcshCmds(i)
        If Not Test.IsExistPath(utFullName) Then
            Call testcshCmdExec.Add("File Not Found!")
        Else
            Let result = Test.SearchInExcelByPattern( _
                utFullName _
                , utSheetName _
                , targetCol _
                , testcshCmd _
                , isRevert _
                , startPattern _
                , stopPattern _
            )
            Call testcshCmdExec.Add(result)
        End If
    Next i
    Call Addr.SetValueArray("TESTCSH_COMAND_EXEC", CollectionToArray(testcshCmdExec))
End Sub

Friend Sub GetALELog()
    Dim dict As Object: Set dict = Scripts.Dictionary
    Dim tasks() As String: Let tasks = Addr.GetValueArray("TASKS")
    Dim lasttimes() As String: Let lasttimes = Addr.GetValueArray("LAST_RUNTIME")
    Dim utPattern As String: Let utPattern = Addr.GetValue("FILE_UT_PATTERN")
    Dim ALSLogs As Collection: Set ALSLogs = New Collection
    Dim result As String
    Dim i As Long: For i = LBound(tasks) To UBound(tasks)
        Dim taskCode As String: Let taskCode = left$(tasks(i), 8)
        Dim taskName As String: Let taskName = Right$(tasks(i), Len(tasks(i)) - 9)
        Let dict("CODE") = taskCode
        Let dict("NAME") = taskName
        Dim utName As String: Let utName = Stringify.ReplacePlaceholder( _
            utPattern _
            , dict _
        )
        Dim utPath As String: Let utPath = Addr.GetValue("UT_PATH")
        Dim utFullPath As String: Let utFullPath = utPath & taskCode & "_" & taskName & "\"
        Dim utFullName As String: Let utFullName = utFullPath & utName
        Dim utSheetName As String: Let utSheetName = Addr.GetValue("UT_SHEET_NAME")
        Dim targetCol As Long: Let targetCol = Test.ColumnNameToIndex("A")
        Dim isRevert As Boolean: Let isRevert = True
        Dim startPattern As String: Let startPattern = Addr.GetValue("SFCMERLG_PATTERN")
        Dim stopPattern As String: Let stopPattern = Addr.GetValue("SFCMAPLG_PATTERN")
        Dim lastime As String: Let lastime = lasttimes(i)
        If Not Test.IsExistPath(utFullName) Then
            Call ALSLogs.Add("File Not Found!")
        Else
            Let result = Test.SearchInExcelByPattern( _
                utFullName _
                , utSheetName _
                , targetCol _
                , lastime _
                , isRevert _
                , startPattern _
                , stopPattern _
            )
            Call ALSLogs.Add(result)
        End If
    Next i
    Call Addr.SetValueArray("ALS_LOG", CollectionToArray(ALSLogs))
End Sub

Friend Sub CheckSfcmaplgOutput()
    Dim dict As Object: Set dict = Scripts.Dictionary
    Dim tasks() As String: Let tasks = Addr.GetValueArray("TASKS")
    Dim lastPids() As String: Let lastPids = Addr.GetValueArray("LAST_RUNTIME")
    Dim utPattern As String: Let utPattern = Addr.GetValue("FILE_UT_PATTERN")
    Dim results As Collection: Set results = New Collection
    Dim noOutput As String: Let noOutput = Addr.GetValue("NO_OUTPUT_PATTERN")
    Dim i As Long: For i = LBound(tasks) To UBound(tasks)
        Dim taskCode As String: Let taskCode = left$(tasks(i), 8)
        Dim taskName As String: Let taskName = Right$(tasks(i), Len(tasks(i)) - 9)
        Let dict("CODE") = taskCode
        Let dict("NAME") = taskName
        Dim utName As String: Let utName = Stringify.ReplacePlaceholder( _
            utPattern _
            , dict _
        )
        Dim utPath As String: Let utPath = Addr.GetValue("UT_PATH")
        Dim utFullPath As String: Let utFullPath = utPath & taskCode & "_" & taskName & "\"
        Dim utFullName As String: Let utFullName = utFullPath & utName
        Dim utSheetName As String: Let utSheetName = Addr.GetValue("UT_SHEET_NAME")
        Dim targetCol As Long: Let targetCol = Test.ColumnNameToIndex("A")
        Dim startPattern As String: Let startPattern = Addr.GetValue("SFCMERLG_PATTERN")
        Dim stopPattern As String: Let stopPattern = Addr.GetValue("SFCMAPLG_PATTERN")
        Dim outputPath As String: Let outputPath = utPath & taskCode & "_" & taskName & "\OUTPUT\"
        Dim sfcmerlgFullName As String: Let sfcmerlgFullName = outputPath & "sfcmaplg"
        Dim isRevert As Boolean: Let isRevert = True
        Dim lastPid As String: Let lastPid = lastPids(i)
        If Not Test.IsExistPath(utFullName) Then
            Call results.Add("File UT Not Found!")
        ElseIf Not Test.IsExistPath(sfcmerlgFullName) Then
            Call results.Add("File Sfcmaplg Found!")
        Else
            Call results.Add(Test.SearchInExcelByPattern( _
                utFullName _
                , utSheetName _
                , targetCol _
                , noOutput _
                , isRevert _
                , startPattern _
                , stopPattern _
            ))
        End If
    Next i
    Call Addr.SetValueArray("CHECK_SFCMAPLG_NO_OUTPUT", CollectionToArray(results))
End Sub

Friend Sub CheckSfcmerlgOutput()
    Dim dict As Object: Set dict = Scripts.Dictionary
    Dim tasks() As String: Let tasks = Addr.GetValueArray("TASKS")
    Dim lastPids() As String: Let lastPids = Addr.GetValueArray("LAST_RUNTIME")
    Dim utPattern As String: Let utPattern = Addr.GetValue("FILE_UT_PATTERN")
    Dim results As Collection: Set results = New Collection
    Dim noOutput As String: Let noOutput = Addr.GetValue("NO_OUTPUT_PATTERN")
    Dim i As Long: For i = LBound(tasks) To UBound(tasks)
        Dim taskCode As String: Let taskCode = left$(tasks(i), 8)
        Dim taskName As String: Let taskName = Right$(tasks(i), Len(tasks(i)) - 9)
        Let dict("CODE") = taskCode
        Let dict("NAME") = taskName
        Dim utName As String: Let utName = Stringify.ReplacePlaceholder( _
            utPattern _
            , dict _
        )
        Dim utPath As String: Let utPath = Addr.GetValue("UT_PATH")
        Dim utFullPath As String: Let utFullPath = utPath & taskCode & "_" & taskName & "\"
        Dim utFullName As String: Let utFullName = utFullPath & utName
        Dim utSheetName As String: Let utSheetName = Addr.GetValue("UT_SHEET_NAME")
        Dim targetCol As Long: Let targetCol = Test.ColumnNameToIndex("A")
        Dim stopPattern As String: Let stopPattern = Addr.GetValue("SFCMERLG_PATTERN")
        Dim outputPath As String: Let outputPath = utPath & taskCode & "_" & taskName & "\OUTPUT\"
        Dim sfcmerlgFullName As String: Let sfcmerlgFullName = outputPath & "sfcmerlg"
        Dim isRevert As Boolean: Let isRevert = True
        Dim lastPid As String: Let lastPid = lastPids(i)
        If Not Test.IsExistPath(utFullName) Then
            Call results.Add("File UT Not Found!")
        ElseIf Not Test.IsExistPath(sfcmerlgFullName) Then
            Call results.Add("File Sfcmaplg Found!")
        Else
            Call results.Add(Test.SearchInExcelByPattern( _
                utFullName _
                , utSheetName _
                , targetCol _
                , noOutput _
                , isRevert _
                , vbNullString _
                , stopPattern _
            ))
        End If
    Next i
    Call Addr.SetValueArray("CHECK_SFCMERLG_NO_OUTPUT", CollectionToArray(results))
End Sub

Friend Sub CheckSfcmerlgLine()
    Dim dict As Object: Set dict = Scripts.Dictionary
    Dim tasks() As String: Let tasks = Addr.GetValueArray("TASKS")
    Dim lastPids() As String: Let lastPids = Addr.GetValueArray("LAST_PIDS")
    Dim utPattern As String: Let utPattern = Addr.GetValue("FILE_UT_PATTERN")
    Dim results As Collection: Set results = New Collection
    Dim i As Long: For i = LBound(tasks) To UBound(tasks)
        Dim taskCode As String: Let taskCode = left$(tasks(i), 8)
        Dim taskName As String: Let taskName = Right$(tasks(i), Len(tasks(i)) - 9)
        Let dict("CODE") = taskCode
        Let dict("NAME") = taskName
        Dim utName As String: Let utName = Stringify.ReplacePlaceholder( _
            utPattern _
            , dict _
        )
        Dim utPath As String: Let utPath = Addr.GetValue("UT_PATH")
        Dim utFullPath As String: Let utFullPath = utPath & taskCode & "_" & taskName & "\"
        Dim utFullName As String: Let utFullName = utFullPath & utName
        Dim utSheetName As String: Let utSheetName = Addr.GetValue("UT_SHEET_NAME")
        Dim targetCol As Long: Let targetCol = Test.ColumnNameToIndex("A")
        Dim stopPattern As String: Let stopPattern = Addr.GetValue("SFCMERLG_PATTERN")
        Dim outputPath As String: Let outputPath = utPath & taskCode & "_" & taskName & "\OUTPUT\"
        Dim sfcmerlgFullName As String: Let sfcmerlgFullName = outputPath & "sfcmerlg"
        Dim isRevert As Boolean: Let isRevert = True
        Dim lastPid As String: Let lastPid = lastPids(i)
        If Not Test.IsExistPath(utFullName) Then
            Call results.Add("File UT Not Found!")
        ElseIf Not Test.IsExistPath(sfcmerlgFullName) Then
            Call results.Add("File sfcmerlg Found!")
        Else
            Dim logs As Collection
            Dim log As Variant
            Set logs = Test.SearchInFileWithPattern( _
                sfcmerlgFullName _
                , "[" & lastPid & "]" _
                , 50 _
                , isRevert _
                , "x-euc-jp" _
            )
            Dim line As String
            Dim isFound As String: Let isFound = True
            If logs.count = 0 Then
                Call results.Add("No sfcmerlg line!")
            Else
                For Each log In logs
                    Let line = Test.SearchInExcelByPattern( _
                        utFullName _
                        , utSheetName _
                        , targetCol _
                        , log _
                        , isRevert _
                        , vbNullString _
                        , stopPattern _
                    )
                    If log <> line Then
                        Let isFound = False
                        Exit For
                    End If
                Next log
                Call results.Add(isFound)
            End If
        End If
    Next i
    Call Addr.SetValueArray("CHECK_SFCMERLG_LINE", CollectionToArray(results))
End Sub

Friend Sub CheckSfcmaplgLine()
    Dim dict As Object: Set dict = Scripts.Dictionary
    Dim tasks() As String: Let tasks = Addr.GetValueArray("TASKS")
    Dim lastPids() As String: Let lastPids = Addr.GetValueArray("LAST_PIDS")
    Dim utPattern As String: Let utPattern = Addr.GetValue("FILE_UT_PATTERN")
    Dim results As Collection: Set results = New Collection
    Dim i As Long: For i = LBound(tasks) To UBound(tasks)
        Dim taskCode As String: Let taskCode = left$(tasks(i), 8)
        Dim taskName As String: Let taskName = Right$(tasks(i), Len(tasks(i)) - 9)
        Let dict("CODE") = taskCode
        Let dict("NAME") = taskName
        Dim utName As String: Let utName = Stringify.ReplacePlaceholder( _
            utPattern _
            , dict _
        )
        Dim utPath As String: Let utPath = Addr.GetValue("UT_PATH")
        Dim utFullPath As String: Let utFullPath = utPath & taskCode & "_" & taskName & "\"
        Dim utFullName As String: Let utFullName = utFullPath & utName
        Dim utSheetName As String: Let utSheetName = Addr.GetValue("UT_SHEET_NAME")
        Dim targetCol As Long: Let targetCol = Test.ColumnNameToIndex("A")
        Dim startPattern As String: Let startPattern = Addr.GetValue("SFCMERLG_PATTERN")
        Dim stopPattern As String: Let stopPattern = Addr.GetValue("SFCMAPLG_PATTERN")
        Dim outputPath As String: Let outputPath = utPath & taskCode & "_" & taskName & "\OUTPUT\"
        Dim sfcmaplgFullName As String: Let sfcmaplgFullName = outputPath & "sfcmaplg"
        Dim isRevert As Boolean: Let isRevert = True
        Dim lastPid As String: Let lastPid = lastPids(i)
        If Not Test.IsExistPath(utFullName) Then
            Call results.Add("File UT Not Found!")
        ElseIf Not Test.IsExistPath(sfcmaplgFullName) Then
            Call results.Add("File sfcmaplg Found!")
        Else
            Dim logs As Collection
            Dim log As Variant
            Set logs = Test.SearchInFileWithPattern( _
                sfcmaplgFullName _
                , "[" & lastPid & "]" _
                , 10 _
                , isRevert _
                , "x-euc-jp" _
            )
            Dim line As String
            Dim isFound As String: Let isFound = True
            If logs.count = 0 Then
                Call results.Add("No sfcmaplg line!")
            Else
                For Each log In logs
                    Let line = Test.SearchInExcelByPattern( _
                        utFullName _
                        , utSheetName _
                        , targetCol _
                        , log _
                        , isRevert _
                        , vbNullString _
                        , stopPattern _
                    )
                    If log <> line Then
                        Let isFound = False
                        Exit For
                    End If
                Next log
                Call results.Add(isFound)
            End If
        End If
    Next i
    Call Addr.SetValueArray("CHECK_SFCMAPLG_LINE", CollectionToArray(results))
End Sub

Friend Sub CheckSfcmerlgError()
    Dim dict As Object: Set dict = Scripts.Dictionary
    Dim tasks() As String: Let tasks = Addr.GetValueArray("TASKS")
    Dim lastPids() As String: Let lastPids = Addr.GetValueArray("LAST_PIDS")
    Dim utPattern As String: Let utPattern = Addr.GetValue("FILE_UT_PATTERN")
    Dim results As Collection: Set results = New Collection
    Dim i As Long: For i = LBound(tasks) To UBound(tasks)
        Dim taskCode As String: Let taskCode = left$(tasks(i), 8)
        Dim taskName As String: Let taskName = Right$(tasks(i), Len(tasks(i)) - 9)
        Let dict("CODE") = taskCode
        Let dict("NAME") = taskName
        Dim utName As String: Let utName = Stringify.ReplacePlaceholder( _
            utPattern _
            , dict _
        )
        Dim utPath As String: Let utPath = Addr.GetValue("UT_PATH")
        Dim utFullPath As String: Let utFullPath = utPath & taskCode & "_" & taskName & "\"
        Dim utFullName As String: Let utFullName = utFullPath & utName
        Dim utSheetName As String: Let utSheetName = Addr.GetValue("UT_SHEET_NAME")
        Dim targetCol As Long: Let targetCol = Test.ColumnNameToIndex("A")
        Dim stopPattern As String: Let stopPattern = Addr.GetValue("SFCMAPLG_PATTERN")
        Dim outputPath As String: Let outputPath = utPath & taskCode & "_" & taskName & "\OUTPUT\"
        Dim sfcmerlgFullName As String: Let sfcmerlgFullName = outputPath & "sfcmerlg"
        Dim isRevert As Boolean: Let isRevert = True
        Dim lastPid As String: Let lastPid = lastPids(i)
        If Not Test.IsExistPath(utFullName) Then
            Call results.Add("File UT Not Found!")
        ElseIf Not Test.IsExistPath(sfcmerlgFullName) Then
            Call results.Add("File sfcmerlg Found!")
        Else
            Dim logs As Collection
            Dim log As Variant
            Set logs = Test.SearchInFileWithPattern( _
                sfcmerlgFullName _
                , "[" & lastPid & "]" _
                , 50 _
                , isRevert _
                , "x-euc-jp" _
            )
            Dim isErr As String: Let isErr = True
            If logs.count = 0 Then
                Call results.Add("No sfcmerlg line!")
            Else
                For Each log In logs
                    If InStr(log, "E @") > 0 Then
                        Let isErr = False
                        Exit For
                    End If
                Next log
                Call results.Add(isErr)
            End If
        End If
    Next i
    Call Addr.SetValueArray("CHECK_SFCMERLG_ERROR", CollectionToArray(results))
End Sub

Friend Sub CheckElsePattern()
    Dim dict As Object: Set dict = Scripts.Dictionary
    Dim tasks() As String: Let tasks = Addr.GetValueArray("TASKS")
    Dim elsePatterns() As String: Let elsePatterns = Addr.GetValueArray("ELSE_PATTERN")
    Dim utPattern As String: Let utPattern = Addr.GetValue("FILE_UT_PATTERN")
    Dim results As Collection: Set results = New Collection
    Dim i As Long: For i = LBound(tasks) To UBound(tasks)
        Dim taskCode As String: Let taskCode = left$(tasks(i), 8)
        Dim taskName As String: Let taskName = Right$(tasks(i), Len(tasks(i)) - 9)
        Let dict("CODE") = taskCode
        Let dict("NAME") = taskName
        Dim utName As String: Let utName = Stringify.ReplacePlaceholder( _
            utPattern _
            , dict _
        )
        Dim utPath As String: Let utPath = Addr.GetValue("UT_PATH")
        Dim utFullPath As String: Let utFullPath = utPath & taskCode & "_" & taskName & "\"
        Dim utFullName As String: Let utFullName = utFullPath & utName
        Dim utSheetName As String: Let utSheetName = Addr.GetValue("UT_SHEET_NAME")
        Dim targetCol As Long: Let targetCol = Test.ColumnNameToIndex("A")
        Dim isRevert As Boolean: Let isRevert = False
        If Not Test.IsExistPath(utFullName) Then
            Call results.Add("File UT Not Found!")
        Else
            Dim line As String
            Dim result As String: Let result = "OK"
            Dim j As Long: For j = LBound(elsePatterns) To UBound(elsePatterns)
                Let line = Test.SearchInExcelByPattern( _
                    utFullName _
                    , utSheetName _
                    , targetCol _
                    , elsePatterns(j) _
                    , isRevert _
                )
                If line = "Text Not Found!" Then
                    Let result = elsePatterns(j)
                    Exit For
                End If
            Next j
            Call results.Add(result)
        End If
    Next i
    Call Addr.SetValueArray("CHECK_ELSE_PATTERN", CollectionToArray(results))
End Sub

Friend Sub ExportImages()
    Dim dict As Object: Set dict = Scripts.Dictionary
    Dim tasks() As String: Let tasks = Addr.GetValueArray("TASKS")
    Dim lastPids() As String: Let lastPids = Addr.GetValueArray("LAST_PIDS")
    Dim utPattern As String: Let utPattern = Addr.GetValue("FILE_UT_PATTERN")
    Dim results As Collection: Set results = New Collection
    Dim i As Long: For i = LBound(tasks) To UBound(tasks)
        Dim taskCode As String: Let taskCode = left$(tasks(i), 8)
        Dim taskName As String: Let taskName = Right$(tasks(i), Len(tasks(i)) - 9)
        Let dict("CODE") = taskCode
        Let dict("NAME") = taskName
        Dim utName As String: Let utName = Stringify.ReplacePlaceholder( _
            utPattern _
            , dict _
        )
        Dim utPath As String: Let utPath = Addr.GetValue("UT_PATH")
        Dim utFullPath As String: Let utFullPath = utPath & taskCode & "_" & taskName & "\"
        Dim utFullName As String: Let utFullName = utFullPath & utName
        Dim utSheetName As String: Let utSheetName = Addr.GetValue("UT_SHEET_NAME")
        If Not Test.IsExistPath(utFullName) Then
            Call results.Add("File UT Not Found!")
        Else
            Dim resultPath As String
            Dim outputFolder As String: Let outputFolder = "D:\share\ASS\Tasks\template\0Ken\task\IMAGES\" & taskCode
            ' Let resultPath = Test.ExtractExcelByZip(utFullName, outputFolder)
            ' If resultPath = vbNullString Then Let resultPath = "Export failed."
            Call results.Add(resultPath)
        End If
    Next i
    Call Addr.SetValueArray("EXPORT_IMAGES_PATH", CollectionToArray(results))
End Sub


Sub XXXTEST()
    Dim testFile As String: Let testFile = "D:\share\ASS\Tasks\template\0Ken\task\IMAGES\test.xlsx"
    Dim outputFile As String: Let outputFile = "D:\share\ASS\Tasks\template\0Ken\task\IMAGES\test.xlsx" & ".zip"
    Dim testFolder As String: Let testFolder = "D:\share\ASS\Tasks\template\0Ken\task\IMAGES\test.xlsx.zip\xl\media\"
    Dim outputFolder As String: Let outputFolder = "D:\share\ASS\Tasks\template\0Ken\task\IMAGES\output"
    Dim testZip As String: Let testZip = "D:\share\ASS\Tasks\template\0Ken\task\IMAGES\test.xlsx.zip"

    ' Debug.Print Test.CopyFile(testFile, outputFile)
    ' Call Test.CopyFolder(testFolder, outputFolder)
    Call Test.ExtractExcelByZip(testZip, outputFolder)

End Sub

