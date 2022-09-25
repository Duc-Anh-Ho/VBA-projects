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

Sub testOnly()
'         Dim arr1, arr2, arr3 As Variant
'        arr1 = Array("a")
'        arr2 = Array("a", 2)
'        arr3 = Array(arr1, arr2)
	Call MainMo.Main
	Stop
End Sub

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
'	On Error GoTo ErrorsHandler
	Call MainMo.initializeWorkplace
	Call MainMo.clearWorkPlace
	Call MainMo.openDataBase
        
	Dim dbWsLastColumn As Long
	Dim dbWsLastRow As Long
	Dim dbWsHeaderRg As Range
	Dim dbWsHeader As Variant
	Dim dbColumn, wpColumn As Long 'Temp
	Dim dataRg As Range
	Dim data As Variant
	Dim dataLength As Long 'Temp
        
	'CORE
	While databasePath <> False 'Cancel select file
		' Database sheets loop
		For Each databaseWs In databaseWb.Worksheets
			Let dbWsLastColumn = _
				System.getLastColumn(databaseWs, eHeader.Row)
			Set dbWsHeaderRg = databaseWs.Range( _
				Cells(eHeader.Row, dbWsLastColumn).Address, _
				Cells(eHeader.Row, eHeader.Row).Address)
			Let dbWsHeader = dbWsHeaderRg
			Let dbWsLastRow = System.getLastRow(databaseWs) 'Take empty row also
Debug.Print "ws: " & databaseWs.Index _
    & Space(1) & databaseWs.name
				' Workplace Column loop
			For wpColumn = eHeader.FirstColumn To wpHeaderLastColumn
				' Skip empty Header
				If Not IsEmpty(wpHeader(eHeader.Row, wpColumn)) Then
					'Database Column Loop
				For dbColumn = eHeader.FirstColumn To dbWsLastColumn
					' Check Matching Header
					If wpHeader(eHeader.Row, wpColumn) = dbWsHeader(eHeader.Row, dbColumn) Then
'Debug.Print dbWsLastRow
						Set dataRg = databaseWs.Range( _
							Cells(eHeader.NextRow, dbColumn).Address, _
							Cells(dbWsLastRow, dbColumn).Address)
						Let data = dataRg
						' SPECIAL: 1st Column is worksheet's name
						' If wpColumn = 1 Then
						' 	Dim item As Long
						' 	For item = LBound(data, 1) To UBound(data, 1) 'Range dimension 1
						' 		data(item, 1) = databaseWs.name
						' 	Next item
						' End If
'Debug.Print Space(4) & wpHeader(eHeader.Row, wpColumn) _
	& Space(1) & dbColumn _
	& Space(1) & dbWsLastRow
					Else
						' Create empty array with sheet column length
						ReDim data(1 To dbWsLastRow - 1, 1 To 1) 'Range dimension 1
					End If
					' Store Database
					Let database(wpColumn) = System.mergeTwoArrays(database(wpColumn), data)
					Let data = Empty
				Next dbColumn
			End If
				Let dataLength = System.getArrayLength(database(wpColumn))   
				' Paste Data
				If dataLength <> False Then 
					workplaceWs.Range(Cells(eHeader.NextRow, wpColumn).Address).Resize(dataLength) = database(wpColumn)
				End If 
			Next wpColumn
		Next databaseWs
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


