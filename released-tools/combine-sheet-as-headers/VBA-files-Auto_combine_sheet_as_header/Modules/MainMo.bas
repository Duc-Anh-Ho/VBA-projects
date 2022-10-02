Attribute VB_Name = "MainMo"
' AUTO COMBINE SHEET AS HEADER TOOL
' Author: DANH
' Version: 1.0.0
' Update: 2022/09/19
' Check README.md for more information

Option Explicit
Private Const AUTHOR_PROMPT As String = "AUTHOR: "
Private Const AUTHOR_NAME As String = "DANH"
Private System As New PerCls
'Declare databse variables
Private database() As Variant
Private databasePath As Variant
Private databaseWb As Workbook
Private databaseWs As Worksheet
Enum eHeader
        Row = 1 'First Row
        NextRow = Row + 1
        FirstColumn = 1
        NextColumn = FirstColumn + 1 'Skip 1st Column
End Enum
'Declare workplace variables
Private workplaceWb As Workbook
Private workplaceWs As Worksheet
Private wpHeaderRg As Range
Private wpHeader As Variant
Private wpHeaderLastColumn As Long

Private Static Sub initializeWorkplace()
        Set workplaceWb = MainWb
        Set workplaceWs = workplaceWb.ActiveSheet
        Let wpHeaderLastColumn = System.getLastColumn(workplaceWs, eHeader.Row)
        Set wpHeaderRg = workplaceWs.Range( _
                Cells(eHeader.Row, wpHeaderLastColumn).Address, _
                Cells(eHeader.Row, eHeader.FirstColumn).Address)
        Let wpHeader = wpHeaderRg
End Sub

Private Static Sub clearWorkPlace()
        'Clean from row 2 to row max - 1048576
        workplaceWs.Rows( _
                eHeader.NextRow & _
                ":" & _
                workplaceWs.Rows.Count).ClearContents
End Sub

Private Static Sub openDataBase()
        Let databasePath = System.getExcelPath()
        Set databaseWb = System.getExcelFile(databasePath)
End Sub

Private Static Sub closeDatabase()
        If databasePath <> False Then
                databaseWb.Close SaveChanges:=False
                Let databasePath = False
        End If
End Sub

Private Static Sub deleteStoredVariables()
        Debug.Print System.getTimerMilestone()
        Set System = Nothing 'Run Class Terminate
        End ' Kill Process
End Sub

'MAIN
Public Sub mainExecute()
        On Error GoTo ErrorsHandler
        Call System.speedOn
        Call MainMo.initializeWorkplace
        Call MainMo.clearWorkPlace
        Call MainMo.openDataBase
        Call System.restartTimer
        'Resize database arr with num of headers
        ReDim database(eHeader.FirstColumn To wpHeaderLastColumn) ' *** Use array store array NOT 2D array.
        'Local Variables
        Dim dbWsLastColumn As Long
        Dim dbWsLastRow As Long
        Dim dbWsHeaderRg As Range
        Dim dbWsHeader As Variant
        Dim dbColumn, wpColumn As Long ' Temp
        Dim dataRg As Range
        Dim data As Variant
        Dim dbHeaderRepeated As Long
        Dim dataLength As Long ' Temp
        ' Cancel select file
        If TypeName(databasePath) = "Boolean" Then GoTo ExecuteProcedure
        ' Workplace header columns loop
        For wpColumn = eHeader.FirstColumn To wpHeaderLastColumn
                ' Database sheets loop
                For Each databaseWs In databaseWb.Worksheets
                        Let dbHeaderRepeated = 0
                        Let dbWsLastColumn = _
                                System.getLastColumn(databaseWs, eHeader.Row)
                        Set dbWsHeaderRg = databaseWs.Range( _
                                Cells(eHeader.Row, dbWsLastColumn).Address, _
                                Cells(eHeader.Row, eHeader.Row).Address)
                        Let dbWsHeader = dbWsHeaderRg
                        Let dbWsLastRow = System.getLastRow(databaseWs) ' Take empty row if not match max column *
                        ' Create empty array with sheet column length *
                        ReDim data(1 To dbWsLastRow - 1, 1 To 1)   'Default Range dimension is 1
                        ' Skip empty Header
                        If Not IsEmpty(wpHeader(eHeader.Row, wpColumn)) Then
                                ' Database header columns loop
                                For dbColumn = eHeader.FirstColumn To dbWsLastColumn
                                        ' Check Matching Header
                                        If wpHeader(eHeader.Row, wpColumn) = dbWsHeader(eHeader.Row, dbColumn) Then
                                                Let dbHeaderRepeated = dbHeaderRepeated + 1
                                                'TODO: Check case header repeated in database
                                                Set dataRg = databaseWs.Range( _
                                                        Cells(eHeader.NextRow, dbColumn).Address, _
                                                        Cells(dbWsLastRow, dbColumn).Address)
                                                Let data = dataRg
                                        End If
                                Next dbColumn
                        End If
                        ' UPDATE: 1st Column is worksheet's name
                         If wpColumn = 1 Then
                               Dim item As Long
                               For item = LBound(data, 1) To UBound(data, 1) 'Range dimension 1
                                       data(item, 1) = databaseWs.Name
                               Next item
                         End If
                      ' Store Database
                        Let database(wpColumn) = System.mergeTwoArrays(database(wpColumn), data)
                        Let data = Empty
                Next databaseWs
                ' Paste Data
                Let dataLength = System.getArrayLength(database(wpColumn))
                If dataLength <> False Then
                        workplaceWs.Range(Cells(eHeader.NextRow, wpColumn).Address).Resize(dataLength) = database(wpColumn)
                End If
        Next wpColumn
        Call MainMo.closeDatabase
        Call System.speedOff
ExecuteProcedure:
        Call MainMo.deleteStoredVariables
ErrorsHandler:
        Call System.speedOff
        Call tackleErrors
        Call MainMo.closeDatabase
        Resume ExecuteProcedure
End Sub

Private Static Sub tackleErrors()
        Select Case Err.Number
                Case 0
                'pass
                'VBA file have password
                Case 50289
                        MsgBox Err.Description & _
                                vbNewLine & _
                                AUTHOR_PROMPT & _
                                "Can not access this VBA file because of been protected by a password!"
                'Data base format not matching
                Case 1004
                         MsgBox Err.Description & _
                                vbNewLine & _
                                AUTHOR_PROMPT & _
                                "Your database you selected headers much be the first row!"
                'Un-handled Error
               Case Else
                        Call System.errorDisplay
        End Select
        On Error GoTo 0
End Sub
