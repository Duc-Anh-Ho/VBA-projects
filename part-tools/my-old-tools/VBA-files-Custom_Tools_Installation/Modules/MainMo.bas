Attribute VB_Name = "MainMo"
Option Explicit
Private System As New PerCls

Private database() As Variant
Private databasePath As Variant
Private databaseWb As Workbook
Private databaseWs As Worksheet

Private workplaceWb As Workbook
Private workplaceWs As Worksheet
Private wpHeaderRg As Range
Private wpHeader As Variant
Private wpHeaderLastColumn As Long

Enum eHeader
        Row = 1 'First Row
        NextRow = Row + 1
        FirstColumn = 1
        NextColumn = FirstColumn + 1 'Skip 1st Column
End Enum

Private Static Sub initializeWorkplace()
        With System
                Set workplaceWb = MainWb
                Set workplaceWs = workplaceWb.Worksheets(.getMainSheetName)
                Let wpHeaderLastColumn = .getLastColumn(workplaceWs, eHeader.Row)
                Set wpHeaderRg = workplaceWs.Range( _
                        Cells(eHeader.Row, wpHeaderLastColumn).Address, _
                        Cells(eHeader.Row, eHeader.FirstColumn).Address)
                Let wpHeader = wpHeaderRg
        End With
End Sub

' Clear old content
Private Static Sub clearWorkPlace()
        With workplaceWs
                Dim sRow As String: sRow = eHeader.NextRow & ":" & .Rows.count
                .Rows(sRow).ClearContents
        End With
End Sub

' Set up database
Private Static Sub openDataBase()
        With System
                Let databasePath = .getExcelPath()
                Set databaseWb = .getExcelFile(databasePath)
                ReDim database(eHeader.FirstColumn To wpHeaderLastColumn)
        End With
End Sub

Private Static Sub closeDatabase()
        With databaseWb
                If databasePath <> False Then
                        .Close SaveChanges:=False
                        Let databasePath = False
                End If
        End With
End Sub

' Kill Process
Private Static Sub deleteStoredVariables()
        Set System = Nothing 'Run Class Terminate
        End
End Sub

' MAIN
Private Static Sub Main()
'System.clearImmediateWindow 'Testing nho xoa
        System.speedOn
'       On Error GoTo ErrorsHandler
        Call MainMo.initializeWorkplace
        Call MainMo.clearWorkPlace
        Call MainMo.openDataBase
        
        Dim dbWsLastColumn As Long
        Dim dbWsLastRow As Long
        Dim dbWsHeaderRg As Range
        Dim dbWsHeader As Variant
        Dim dbColumn, wpColumn As Long ' Temp
        Dim dataRg As Range
        Dim data As Variant
        Dim dbHeaderRepeated As Long
        Dim dataLength As Long ' Temp
        
        ' CORE
        While databasePath <> False ' Cancel select file
                ' Workplace Column loop
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
                                Let dbWsLastRow = System.getLastRow(databaseWs) ' Take empty row if not match max colum *
                                ' Create empty array with sheet column length *
                                ReDim data(1 To dbWsLastRow - 1, 1 To 1)   'Range dimension 1
                                ' Skip empty Header
                                If Not IsEmpty(wpHeader(eHeader.Row, wpColumn)) Then
                                        ' Database Column Loop
                                        For dbColumn = eHeader.FirstColumn To dbWsLastColumn
                                                ' Check Matching Header
                                                If wpHeader(eHeader.Row, wpColumn) = dbWsHeader(eHeader.Row, dbColumn) Then
                                                        Let dbHeaderRepeated = dbHeaderRepeated + 1
                                                        Set dataRg = databaseWs.Range( _
                                                                Cells(eHeader.NextRow, dbColumn).Address, _
                                                                Cells(dbWsLastRow, dbColumn).Address) ' *
                                                        Let data = dataRg
                                                End If
                                        Next dbColumn
                                End If
                                ' SPECIAL: 1st Column is worksheet's name
                                 If wpColumn = 1 Then
                                       Dim item As Long
                                       For item = LBound(data, 1) To UBound(data, 1) 'Range dimension 1
                                               data(item, 1) = databaseWs.name
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
        Wend
        System.speedOff
ExecuteProcedure:
        Call MainMo.deleteStoredVariables
ErrorsHandler:
        System.speedOff
        System.errorDisplay
        Call MainMo.closeDatabase
        Resume ExecuteProcedure
End Sub


