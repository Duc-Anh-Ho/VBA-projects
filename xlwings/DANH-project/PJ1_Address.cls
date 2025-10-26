VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PJ1_Address"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Implements Interface_Address
' Check README.md for more information
Option Explicit
' --- Dependencies ---
' --- Object ---
Friend Sub Interface_Address_Pair(ByRef Addr As Utils_Address)
    Call Addr.Add("FILE_UT_PATTERN", "B5")
    Call Addr.Add("FILE_CHECKLIST_PATTERN", "B6")
    Call Addr.Add("CHECKLIST_PATH", "B7")
    Call Addr.Add("UT_PATH", "B8")
    Call Addr.Add("LOG_CONTENT_PATTERN", "B9")
    Call Addr.Add("UT_SHEET_NAME", "B10")
    Call Addr.Add("SFCMAPLG_PATTERN", "B11")
    Call Addr.Add("SFCMERLG_PATTERN", "B12")
    Call Addr.Add("EXEC_CMD_PATTERN", "B13")
    Call Addr.Add("EXEC_LOG_PATTERN", "B14")
    Call Addr.Add("RESULT_PATTERN_1", "B15")
    Call Addr.Add("RESULT_PATTERN_2", "B16")
    Call Addr.Add("NO_OUTPUT_PATTERN", "B17")
    Call Addr.Add("TEMPORARY", "B19")
    Call Addr.Add("ELSE_PATTERN", "D4:D47")
    Call Addr.Add("SFCMAPLG_FULLNAME", "M4:M22")
    Call Addr.Add("SFCMERLG_FULLNAME", "N4:N22")
    Call Addr.Add("TASKS", "L4:L22")
    Call Addr.Add("CHECKLIST_FULLNAME", "O4:O22")
    Call Addr.Add("UT_FULLNAME", "P4:P22")
    Call Addr.Add("LOG_FULLNAME", "Q4:Q22")
    Call Addr.Add("RESULT_CHECKLIST_EXITS", "W4:W22")
    Call Addr.Add("RESULT_UT_EXITS", "X4:X22")
    Call Addr.Add("RESULT_COUNT_IN_UT", "Y4:Y22")
    Call Addr.Add("RESULT_COUNT_IN_OUTPUT", "Z4:Z22")
    Call Addr.Add("RESULT_SFCMAPLG_EXITS", "AA4:AA22")
    Call Addr.Add("RESULT_SFCMERLG_EXITS", "AB4:AB22")
    Call Addr.Add("RESULT_LOG_EXITS", "AC4:AC22")
    Call Addr.Add("RESULT_LOG_EOF", "AD4:AD22")
    Call Addr.Add("RESULT_SFCMAPLG_EOF", "AE4:AE22")
    Call Addr.Add("RESULT_SFCMERLG_EOF", "AF4:AF22")
    Call Addr.Add("LOG_CONTENT", "AG4:AG22")
    Call Addr.Add("RESULT_LOG_CONTENT", "AH4:AH22")
    Call Addr.Add("LAST_RUNLOG", "AI4:AI22")
    Call Addr.Add("LAST_RUNTIME", "AJ4:AJ22")
    Call Addr.Add("LAST_PIDS", "AK4:AK22")
    Call Addr.Add("TESTCSH_COMAND", "AL4:AL22")
    Call Addr.Add("TESTCSH_COMAND_EXEC", "AM4:AM22")
    Call Addr.Add("ALS_LOG", "AO4:AO22")
    Call Addr.Add("CHECK_SFCMAPLG_NO_OUTPUT", "AQ4:AQ22")
    Call Addr.Add("CHECK_SFCMERLG_NO_OUTPUT", "AR4:AR22")
    Call Addr.Add("CHECK_SFCMERLG_LINE", "AS4:AS22")
    Call Addr.Add("CHECK_SFCMERLG_ERROR", "AT4:AT22")
    Call Addr.Add("CHECK_ELSE_PATTERN", "AU4:AU22")
    Call Addr.Add("CHECK_SFCMAPLG_LINE", "AV4:AV22")
    Call Addr.Add("EXPORT_IMAGES_PATH", "AW4:AW22")
End Sub


