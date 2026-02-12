Attribute VB_Name = "Developer"
Option Explicit
' FOR DEVELOPER ONLY
Public form As KeyboardShortcutForm
Private Const GIT_LOCAL_PATH As String = "S:\VBA-projects\"
Private Const INSTALL_FILE_NAME As String = "Danh-Tools-Installation.xlsb"
Private Const INSTALL_FILE_FULLNAME As String = GIT_LOCAL_PATH & INSTALL_FILE_NAME
Private Const ADDIN_FILE_NAME As String = "Danh-Tools.xlam"
Private info As InfoConstants
Private system As SystemUpdate
Private fileSystem As Object
Private userResponse As VbMsgBoxResult

' WINDOW Detect DPI
Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hdc As LongPtr, ByVal nIndex As Long) As Long
Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hwnd As LongPtr, ByVal hdc As LongPtr) As Long

Const LOGPIXELSX = 88 ' Horizontal DPI
Const LOGPIXELSY = 90 ' Vertical DPI
'REF: https://learn.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/saveoptionsenum
Private Enum SAVE_OPTIONS
    NotExist = 1 'adSaveCreateNotExist
    OverWrite = 2 'adSaveCreateOverWrite
End Enum
' REF: https://learn.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/streamtypeenum
Private Enum STREAM_TYPE
    Binary = 1 ' adTypeBinary
    Text = 2 ' adTypeText
End Enum
Public Sub showDPI()
    Dim hdc As LongPtr: hdc = GetDC(0)
    Dim dpiX As Long, dpiY As Long
    Let dpiX = GetDeviceCaps(hdc, LOGPIXELSX)
    Let dpiY = GetDeviceCaps(hdc, LOGPIXELSY)
    Call ReleaseDC(0, hdc)

    Debug.Print "Screen DPI" & vbCrLf & "DPI X: " & dpiX & vbCrLf & "DPI Y: " & dpiY
End Sub

Public Sub aSaveBackup()
    If ThisWorkbook.name = ADDIN_FILE_NAME Then
        ThisWorkbook.SaveAs _
            filename:=INSTALL_FILE_FULLNAME, _
            FileFormat:=xlExcel12 ' xlExcel12 = xlsb
    End If
    Application.OnTime Now + TimeValue("00:00:03"), "reOpen"
    ThisWorkbook.Close
End Sub

Private Sub reOpen()
    Workbooks.Open (INSTALL_FILE_FULLNAME)
End Sub

'NOTE: JUST FOR TESTING DON'T RELEASE
Public Sub autoSendWifi()
    Dim internetC As InternetConnector
    Dim email As EmailCDO
    Dim attachment As String
    Set internetC = New InternetConnector
    Set email = New EmailCDO
    Let attachment = ThisWorkbook.path & "\wifi.txt"
    Call internetC.saveWifiAsTxt(True)
    Call email.send(attachment)
    Kill (attachment)
End Sub

' TEST
Public Sub aaTestCode()
    ' Set form = New KeyboardShortcutForm
    ' Call form.Show(vbModal)  ' vbModeless or vbModal
    ' Set form = Nothing
    Call Shortcuts.Test
    ' Call KeyboardShortcutForm.Show(vbModal)
'    If ActiveWorkbook.path = "" Then MsgBox "Not saved"
''''''''''''''''''''
'    Dim system As SystemUpdate
'    Dim PWShell As PowerShellController
'
'    Set system = New SystemUpdate
'    Set PWShell = New PowerShellController
    
'    Debug.Print PWShell.executeScript("ipconfig /all")
''''''''''''''''''
'    Dim system As New SystemUpdate
'    Debug.Print "====="
''    Debug.Print system.hasWorkPlace(True)
''    Debug.Print system.hasWorkPlace(True, "asd")
'    Dim fileController As New FilesController
'    fileController.copyFileName
'    Debug.Print system.getClipboard()
'    Debug.Print "====="
    'Microsoft Excel
    'Workbook
    'Worksheet
    'Chart
    'DialogSheet
    'xlWorksheet
    'xlChart
    'xlExcel4MacroSheet
    'xlExcel4IntMacroSheet
End Sub

' 2025-10-18
' Helper methods that using in Imediate Window

Public Function Clip(Optional ByRef text As String = vbNullString)
    Dim Clipboard As Utils_Clipboard
    Set Clipboard = New Utils_Clipboard
    If text <> vbNullString Then Call Clipboard.SaveByCOM(text)
    Let Clip = Clipboard.LoadByCOM
End Function

Public Function Clip2(Optional ByRef text As String = vbNullString)
    Dim Clipboard As Utils_Clipboard
    Set Clipboard = New Utils_Clipboard
    If text <> vbNullString Then Call Clipboard.SaveByAPI(text)
    Let Clip2 = Clipboard.LoadByAPI
End Function

Public Sub CloseVBEProjectWindow()
    Dim VBE As Utils_VBE
    Set VBE = New Utils_VBE
    Call VBE.CloseProjectWindow
End Sub

Public Sub CloseVBEPropertiesWindow()
    Dim VBE As Utils_VBE
    Set VBE = New Utils_VBE
    Call VBE.ClosePropertiesWindow
End Sub

Public Sub CloseVBEImmediateWindow()
    Dim VBE As Utils_VBE
    Set VBE = New Utils_VBE
    Call VBE.CloseImmediateWindow
End Sub

Public Sub CloseAllVBEComponents()
    Dim VBE As Utils_VBE
    Set VBE = New Utils_VBE
    Call VBE.CloseAllComponents
End Sub

Public Sub CloseAllVBEWindows()
    Dim VBE As Utils_VBE
    Set VBE = New Utils_VBE
    Call VBE.CloseAllWindows
End Sub

Public Sub ListComponents(Optional ByRef typeName As String = vbNullString)
    Dim VBE As Utils_VBE
    Set VBE = New Utils_VBE
    Call VBE.ListComponentsByTypeName(typeName)
End Sub

Public Sub OpenComponent(ByRef name As String)
    Dim VBE As Utils_VBE
    Set VBE = New Utils_VBE
    Call VBE.OpenComponentByName(name)
End Sub

Public Sub ToggleVBEToolbars()
    Dim VBE As Utils_VBE
    Set VBE = New Utils_VBE
    Call VBE.ToggleToolbars
End Sub

Public Sub Cls()
    Dim VBE As Utils_VBE
    Set VBE = New Utils_VBE
    Call VBE.ClearImmediateWindow
End Sub

Public Sub Clear()
    Dim VBE As Utils_VBE
    Set VBE = New Utils_VBE
    Call VBE.ClearImmediateWindowUnix
End Sub

Public Sub ClsAll()
    Dim VBE As Utils_VBE
    Set VBE = New Utils_VBE
    Call VBE.CloseAllWindows ' CloseAllVBEWindows
    Call VBE.ClearImmediateWindow ' Cls
End Sub

Public Sub PrintDictionary(ByRef Dict As Object)
    Dim VBE As Utils_VBE
    Set VBE = New Utils_VBE
    Call VBE.PrintParsedDictionary(Dict)
End Sub

Public Sub PrintStringArray(ByRef StrArr As Variant)
    Dim VBE As Utils_VBE
    Set VBE = New Utils_VBE
    Call VBE.PrintParsedStringArray(StrArr)
End Sub

Public Sub PrintCollection(ByRef coll As Collection)
    Dim VBE As Utils_VBE
    Set VBE = New Utils_VBE
    Call VBE.PrintParsedCollection(coll)
End Sub

' TODO Move to arrayController or utils
Public Function CollectionToArray(ByRef coll As Collection) As String()
    Dim arr() As String
    ReDim arr(1 To coll.Count)
    Dim i As Long: For i = 1 To coll.Count
        Let arr(i) = CStr(coll(i))
    Next i
    Let CollectionToArray = arr
End Function

Public Function ArrayToCollection(ByRef arr() As String) As Collection
    Dim coll As Collection: Set coll = New Collection
    Dim i As Long: For i = LBound(arr) To UBound(arr)
        Call coll.Add(arr(i))
    Next i
    Set ArrayToCollection = coll
End Function

Public Function RangeToArray(ByRef rng As Range) As String()
    Dim arr() As String: ReDim arr(1 To rng.Cells.Count)
    Dim i As Long: For i = 1 To rng.Cells.Count
        Let arr(i) = CStr(rng.Cells(i).Value)
    Next i
    Let RangeToArray = arr
End Function

Public Function RangeToCollection(ByRef rng As Range) As Collection
    Dim coll As Collection: Set coll = New Collection
    Dim cell As Range: For Each cell In rng.Cells
        Call coll.Add(CStr(cell.Value))
    Next cell
    Set RangeToCollection = coll
End Function

Public Function ArrayToRange(ByRef arr() As String, ByRef target As Range)
    Dim i As Long: For i = LBound(arr) To UBound(arr)
        Let target.Cells(i, 1).Value = arr(i)
    Next i
End Function

Public Function CollectionToRange(ByRef coll As Collection, ByRef target As Range)
    Dim i As Long: For i = 1 To coll.Count
        Let target.Cells(i, 1).Value = coll(i)
    Next i
End Function

Public Function Array2DToArray(ByRef arr2D As Variant) As String()
    Dim rows As Long: Let rows = UBound(arr2D, 1)
    Dim cols As Long: Let cols = UBound(arr2D, 2)
    Dim arr() As String: ReDim arr(1 To rows * cols)
    Dim i As Long, j As Long, k As Long: For i = 1 To rows
        For j = 1 To cols
            Let k = k + 1
            Let arr(k) = CStr(arr2D(i, j))
        Next j
    Next i
    Let Array2DToArray = arr
End Function

Public Function ArrayToArray2D(ByRef arr() As String, ByRef numCols As Long) As Variant
    Dim total As Long: Let total = UBound(arr) - LBound(arr) + 1
    Dim numRows As Long: Let numRows = Application.WorksheetFunction.RoundUp(total / numCols, 0)
    Dim arr2D() As Variant: ReDim arr2D(1 To numRows, 1 To numCols)
    Dim index As Long: index = LBound(arr)
    Dim rowNo As Long, columnNo As Long: For rowNo = 1 To numRows
        For columnNo = 1 To numCols
            If index <= UBound(arr) Then
                Let arr2D(rowNo, columnNo) = arr(index)
                Let index = index + 1
            Else
                Let arr2D(rowNo, columnNo) = ""
            End If
        Next columnNo
    Next rowNo
    Let ArrayToArray2D = arr2D
End Function

Public Sub RangeToFile( _
    ByRef rng As Range _
    , ByRef filePath As String _
    , Optional ByRef delimiter As String = "," _
    , Optional ByRef includeHeader As Boolean = True _
    , Optional ByRef charset As String = "utf-8" _
)
    Dim startRow As Long: startRow = IIf(includeHeader, 1, 2)
    Dim scripts As Utils_Scripts: Set scripts = New Utils_Scripts
    Dim fileStream As Object: Set fileStream = scripts.FileStream
    Let fileStream.Type = STREAM_TYPE.Text
    Let fileStream.Charset = charset
    Call fileStream.Open
    Dim rowNo As Long: For rowNo = startRow To rng.Rows.Count
        Dim rowRng As Range: Set rowRng = rng.Rows(rowNo)
        Dim arr() As String: ReDim arr(1 To rowRng.Columns.Count)
        Dim i As Long: For i = 1 To rowRng.Columns.Count
            Let arr(i) = CStr(rowRng.Cells(1, i).Text)
        Next i
        Call fileStream.WriteText(Join(arr, delimiter) & vbCrLf)
    Next rowNo
    Call fileStream.SaveToFile(filePath, SAVE_OPTIONS.OverWrite)
    Call fileStream.Close
End Sub

Public Sub FileToRange( _
    ByRef filePath As String _
    , ByRef target As Range _
    , Optional ByRef delimiter As String = "," _
    , Optional ByRef charset As String = "utf-8" _
)
    Dim scripts As Utils_Scripts: Set scripts = New Utils_Scripts
    Dim fileStream As Object: Set fileStream = scripts.FileStream
    Let fileStream.Type = STREAM_TYPE.Text
    Let fileStream.Charset = charset
    Call fileStream.Open
    Call fileStream.LoadFromFile(filePath)
    Dim textContent As String: Let textContent = fileStream.ReadText
    Call fileStream.Close
    Dim lines() As String: Let lines = Split(textContent, vbCrLf)
    Dim rowNo As Long: For rowNo = 0 To UBound(lines)
        If Trim(lines(rowNo)) <> "" Then
            Dim arr() As String: Let arr = Split(lines(rowNo), delimiter)
            Dim colNo As Long: For colNo = LBound(arr) To UBound(arr)
                Let target.Cells(rowNo + 1, colNo + 1).Value = arr(colNo)
            Next colNo
        End If
    Next rowNo
End Sub

