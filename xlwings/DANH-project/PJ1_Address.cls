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
    Const MAX_LINE As String = "229"
    ' Const MAX_LINE As String = "10"
    Call Addr.add("FILE_UT_PATTERN"                    , "B5")
    Call Addr.add("FILE_CHECKLIST_PATTERN"             , "B6")
    Call Addr.add("CHECKLIST_PATH"                     , "B7")
    Call Addr.add("UT_PATH"                            , "B8")
    Call Addr.add("LOG_CONTENT_PATTERN"                , "B9")
    Call Addr.add("UT_SHEET_NAME"                      , "B10")
    Call Addr.add("SFCMAPLG_PATTERN"                   , "B11")
    Call Addr.add("SFCMERLG_PATTERN"                   , "B12")
    Call Addr.add("EXEC_CMD_PATTERN"                   , "B13")
    Call Addr.add("EXEC_LOG_PATTERN"                   , "B14")
    Call Addr.add("RESULT_PATTERN_1"                   , "B15")
    Call Addr.add("RESULT_PATTERN_2"                   , "B16")
    Call Addr.add("NO_OUTPUT_PATTERN"                  , "B17")
    Call Addr.add("TEMPORARY"                          , "B19")
    Call Addr.add("ELSE_PATTERN"                       , "D4:D47")
    Call Addr.add("SFCMAPLG_FULLNAME"                  , "M4:M" & MAX_LINE)
    Call Addr.add("SFCMERLG_FULLNAME"                  , "N4:N" & MAX_LINE)
    Call Addr.add("TASKS"                              , "L4:L" & MAX_LINE)
    Call Addr.add("CHECKLIST_FULLNAME"                 , "O4:O" & MAX_LINE)
    Call Addr.add("UT_FULLNAME"                        , "P4:P" & MAX_LINE)
    Call Addr.add("LOG_FULLNAME"                       , "Q4:Q" & MAX_LINE)
    Call Addr.add("RESULT_CHECKLIST_EXITS"             , "W4:W" & MAX_LINE)
    Call Addr.add("RESULT_UT_EXITS"                    , "X4:X" & MAX_LINE)
    Call Addr.add("RESULT_COUNT_IN_UT"                 , "Y4:Y" & MAX_LINE)
    Call Addr.add("RESULT_COUNT_IN_OUTPUT"             , "Z4:Z" & MAX_LINE)
    Call Addr.add("RESULT_SFCMAPLG_EXITS"              , "AA4:AA" & MAX_LINE)
    Call Addr.add("RESULT_SFCMERLG_EXITS"              , "AB4:AB" & MAX_LINE)
    Call Addr.add("RESULT_LOG_EXITS"                   , "AC4:AC" & MAX_LINE)
    Call Addr.add("RESULT_LOG_EOF"                     , "AD4:AD" & MAX_LINE)
    Call Addr.add("RESULT_SFCMAPLG_EOF"                , "AE4:AE" & MAX_LINE)
    Call Addr.add("RESULT_SFCMERLG_EOF"                , "AF4:AF" & MAX_LINE)
    Call Addr.add("LOG_CONTENT"                        , "AG4:AG" & MAX_LINE)
    Call Addr.add("RESULT_LOG_CONTENT"                 , "AH4:AH" & MAX_LINE)
    Call Addr.add("LAST_RUNLOG"                        , "AI4:AI" & MAX_LINE)
    Call Addr.add("LAST_RUNTIME"                       , "AJ4:AJ" & MAX_LINE)
    Call Addr.add("LAST_PIDS"                          , "AK4:AK" & MAX_LINE)
    Call Addr.add("TESTCSH_COMAND"                     , "AL4:AL" & MAX_LINE)
    Call Addr.add("TESTCSH_COMAND_EXEC"                , "AM4:AM" & MAX_LINE)
    Call Addr.add("ALS_LOG"                            , "AO4:AO" & MAX_LINE)
    Call Addr.add("CHECK_SFCMAPLG_NO_OUTPUT"           , "AQ4:AQ" & MAX_LINE)
    Call Addr.add("CHECK_SFCMERLG_NO_OUTPUT"           , "AR4:AR" & MAX_LINE)
    Call Addr.add("CHECK_SFCMERLG_LINE"                , "AS4:AS" & MAX_LINE)
    Call Addr.add("CHECK_SFCMERLG_ERROR"               , "AT4:AT" & MAX_LINE)
    Call Addr.add("CHECK_ELSE_PATTERN"                 , "AU4:AU" & MAX_LINE)
    Call Addr.add("CHECK_SFCMAPLG_LINE"                , "AV4:AV" & MAX_LINE)
    Call Addr.add("EXPORT_IMAGES_PATH"                 , "AW4:AW" & MAX_LINE)
    Call Addr.add("CONFIG_LIST"                        , "BA3:BD" & MAX_LINE)
End Sub

