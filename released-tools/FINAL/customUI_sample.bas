' Check README.md for more information

Option Explicit
'Declare Variables
Private system As SystemUpdate
Private fileSystem As Object
Private info As MyInfo
Private userResponse As VbMsgBoxResult

Private NewAddin As AutoAddin
Private sheetC As SheetsController
Private chartC As ChartsController
Private pivotC As PivotTablesController
Private fileC As FilesController
Private formatC As FormatController
Private pictureC As PicturesController
Private rangeC As RangesController

Private toolsTab As TagController
Private sheetGroup As TagController
Private addSheetsButton As TagController
Private listSheetsButton As TagController
Private deleteSheetsButton As TagController
Private showSheetsButton As TagController
Private hideSheetsButton As TagController
Private veryHideSheetsButton As TagController
Private chartGroup As TagController
Private chartHideErrButton As TagController
Private chartShowButton As TagController
Private pivotGroup As TagController
Private refeshPivotButton As TagController
Private vbaFileGroup As TagController
Private importVbaFiles As TagController
Private importAllVbaFiles As TagController
Private exportAllVbaFiles As TagController
Private infoGroup As TagController
Private toolNameLabel As TagController
Private versionLabel As TagController

Private hasCustomUI As Boolean
Private loadedRibbon As IRibbonUI 'TO-DO: ?
'Constructor
Private Sub Auto_Open()
    '
End Sub
Private Sub Auto_Activate()
    '
End Sub
'Destructor
Private Sub Auto_Close()
    '
End Sub
Private Sub Auto_Deactivate()
    '
End Sub
'METHODS
Private Sub createInstances()
    Set toolsTab = New TagController
    Set sheetGroup = New TagController
    Set addSheetsButton = New TagController
    Set listSheetsButton = New TagController
    Set deleteSheetsButton = New TagController
    Set showSheetsButton = New TagController
    Set hideSheetsButton = New TagController
    Set veryHideSheetsButton = New TagController
    Set infoGroup = New TagController
    Set chartGroup = New TagController
    Set chartHideErrButton = New TagController
    Set chartShowButton = New TagController
    Set pivotGroup = New TagController
    Set vbaFileGroup = New TagController
    Set importVbaFiles = New TagController
    Set importAllVbaFiles = New TagController
    Set exportAllVbaFiles = New TagController
    Set refeshPivotButton = New TagController
    Set toolNameLabel = New TagController
    Set versionLabel = New TagController
End Sub

Private Sub setUpId()
    Let toolsTab.letID = "danh-tools"
    Let sheetGroup.letID = "sheets-controller"
    Let addSheetsButton.letID = "add-sheets"
    Let listSheetsButton.letID = "list-sheets"
    Let deleteSheetsButton.letID = "delete-sheets"
    Let showSheetsButton.letID = "show-sheets"
    Let hideSheetsButton.letID = "hide-sheets"
    Let veryHideSheetsButton.letID = "very-hide-sheets"
    Let chartGroup.letID = "charts-controller"
    Let chartHideErrButton.letID = "hide-error-labels"
    Let chartShowButton.letID = "show-labels"
    Let pivotGroup.letID = "pivot-controller"
    Let refeshPivotButton.letID = "refesh-pivot"
    Let vbaFileGroup.letID = "vba-files-controller"
    Let importVbaFiles.letID = "import-vba-files"
    Let importAllVbaFiles.letID = "import-all-vba-files"
    Let exportAllVbaFiles.letID = "export-all-vba-files"
    Let infoGroup.letID = "infomation"
    Let toolNameLabel.letID = "tool-name"
    Let versionLabel.letID = "version"
End Sub

Private Sub setUpDescription()
    Let addSheetsButton.letDescription = ""
    Let listSheetsButton.letDescription = ""
    Let deleteSheetsButton.letDescription = ""
    Let showSheetsButton.letDescription = ""
    Let hideSheetsButton.letDescription = ""
    Let veryHideSheetsButton.letDescription = ""
    Let chartHideErrButton.letDescription = ""
    Let chartShowButton.letDescription = ""
    Let refeshPivotButton.letDescription = ""
    Let importVbaFiles.letDescription = ""
    Let importAllVbaFiles.letDescription = ""
    Let exportAllVbaFiles.letDescription = ""
End Sub

Private Sub setUpEnabled()
    Dim isEnabled As Boolean
    Let isEnabled = hasWorkPlace()
    Let addSheetsButton.letEnabled = isEnabled
    Let listSheetsButton.letEnabled = isEnabled
    Let deleteSheetsButton.letEnabled = isEnabled
    Let showSheetsButton.letEnabled = isEnabled
    Let hideSheetsButton.letEnabled = isEnabled
    Let veryHideSheetsButton.letEnabled = isEnabled
    Let chartHideErrButton.letEnabled = isEnabled
    Let chartShowButton.letEnabled = isEnabled
    Let refeshPivotButton.letEnabled = isEnabled
    Let importVbaFiles.letEnabled = isEnabled
    Let importAllVbaFiles.letEnabled = isEnabled
    Let exportAllVbaFiles.letEnabled = isEnabled
End Sub

Private Sub setUpShowImage()
    Dim isShowed As Boolean
    Let isShowed = True
    Let addSheetsButton.letShowImage = isShowed
    Let listSheetsButton.letShowImage = isShowed
    Let deleteSheetsButton.letShowImage = isShowed
    Let showSheetsButton.letShowImage = isShowed
    Let hideSheetsButton.letShowImage = isShowed
    Let veryHideSheetsButton.letShowImage = isShowed
    Let chartHideErrButton.letShowImage = isShowed
    Let chartShowButton.letShowImage = isShowed
    Let refeshPivotButton.letShowImage = isShowed
    Let importVbaFiles.letShowImage = isShowed
    Let importAllVbaFiles.letShowImage = isShowed
    Let exportAllVbaFiles.letShowImage = isShowed
End Sub

Private Sub setUpKeytip()
    Let addSheetsButton.letKeytip = ""
    Let listSheetsButton.letKeytip = ""
    Let deleteSheetsButton.letKeytip = ""
    Let showSheetsButton.letKeytip = ""
    Let hideSheetsButton.letKeytip = ""
    Let veryHideSheetsButton.letKeytip = ""
    Let chartHideErrButton.letKeytip = ""
    Let chartShowButton.letKeytip = ""
    Let refeshPivotButton.letKeytip = ""
    Let importVbaFiles.letKeytip = ""
    Let importAllVbaFiles.letKeytip = ""
    Let exportAllVbaFiles.letKeytip = ""
End Sub

Private Sub setUpLabel()
    Let toolsTab.letLabel = "DANH Tools"
    Let sheetGroup.letLabel = "Sheet Controller"
    Let addSheetsButton.letLabel = "Add Sheets"
    Let listSheetsButton.letLabel = "List Sheets"
    Let deleteSheetsButton.letLabel = "Delete Sheets"
    Let showSheetsButton.letLabel = "Show Sheets"
    Let hideSheetsButton.letLabel = "Hide Sheets"
    Let veryHideSheetsButton.letLabel = "Very Hide Sheet"
    Let chartGroup.letLabel = "Chart Controller"
    Let chartHideErrButton.letLabel = "Hide Err Labels"
    Let chartShowButton.letLabel = "Show Labels"
    Let refeshPivotButton.letLabel = "SYNC Pivot"
    Let pivotGroup.letLabel = "Pivot Controller"
    Let infoGroup.letLabel = "Information"
    Let vbaFileGroup.letLabel = "VBA Files"
    Let importVbaFiles.letLabel = "Import Files"
    Let importAllVbaFiles.letLabel = "Import All Files"
    Let exportAllVbaFiles.letLabel = "Export All Files"
    Let toolNameLabel.letLabel = "Tool: DANH"
    Let versionLabel.letLabel = "Version: v2.0.0"
End Sub

Private Sub setUpShowLabel()
    Dim isShowed As Boolean
    Let isShowed = True
    Let addSheetsButton.letShowLabel = isShowed
    Let listSheetsButton.letShowLabel = isShowed
    Let deleteSheetsButton.letShowLabel = isShowed
    Let showSheetsButton.letShowLabel = isShowed
    Let hideSheetsButton.letShowLabel = isShowed
    Let veryHideSheetsButton.letShowLabel = isShowed
    Let chartHideErrButton.letShowLabel = isShowed
    Let chartShowButton.letShowLabel = isShowed
    Let refeshPivotButton.letShowLabel = isShowed
    Let importVbaFiles.letShowLabel = isShowed
    Let importAllVbaFiles.letShowLabel = isShowed
    Let exportAllVbaFiles.letShowLabel = isShowed
End Sub

Private Sub setUpScreentip()
    Let addSheetsButton.letScreentip = "Add Sheets"
    Let listSheetsButton.letScreentip = "List Sheets"
    Let deleteSheetsButton.letScreentip = "Delete Sheets"
    Let showSheetsButton.letScreentip = "Show Sheets"
    Let hideSheetsButton.letScreentip = "Hide Sheets"
    Let veryHideSheetsButton.letScreentip = "Very Hide Sheet"
    Let chartHideErrButton.letScreentip = "Hide Err Labels"
    Let chartShowButton.letScreentip = "Show Labels"
    Let importVbaFiles.letScreentip = "Import Files"
    Let importAllVbaFiles.letScreentip = "Import All Files"
    Let exportAllVbaFiles.letScreentip = "Export All Files"
End Sub

Private Sub setUpSupertip()
    Let addSheetsButton.letSupertip = _
        "1.Select creating name Ranges." & vbNewLine & _
        "2.Click button." & vbNewLine & _
        "3.Auto create Sheets as each cell." & vbNewLine & _
        "Note: Will not working when sheets already exist!"
    Let listSheetsButton.letSupertip = _
        "ON: List all sheets name at columns A and B." & vbNewLine & _
        "OFF: if ON, turn off lists."
    Let deleteSheetsButton.letSupertip = _
        "1.Delete all sheets except hided sheets." & vbNewLine & _
        "Note: Becareful, after deleted sheets can be undo!"
    Let showSheetsButton.letSupertip = _
        "Show all sheets included very hide"
    Let hideSheetsButton.letSupertip = _
        "1.Select sheets you don't want to hide." & vbNewLine & _
        "2.Click hide Sheets button." & vbNewLine & _
        "3.Auto hide all un-selected sheets."
    Let veryHideSheetsButton.letSupertip = _
        "1.Select sheets you don't want to very hide." & vbNewLine & _
        "2.Click hide Sheets button." & vbNewLine & _
        "3.Auto very hide all un-selected sheets." & vbNewLine & _
        "Note: You could click show Sheets to show all!"
    Let chartHideErrButton.letSupertip = _
        "1.Select a chart that you want to hide error labels." & vbNewLine & _
        "2.Click Hide Error button." & vbNewLine & _
        "3.Auto very hide all error labels."
    Let chartShowButton.letSupertip = _
        "Show all labels of a chart"
    Let refeshPivotButton.letSupertip = _
        "Auto update pivot tables" & vbNewLine & _
        "Note: Can't use Undo while this ON. (Update later)"
    Let importVbaFiles.letSupertip = _
        "Choose one or many VBA files to import"
    Let importAllVbaFiles.letSupertip = _
        "Auto import all VBA files in default folder"
    Let exportAllVbaFiles.letSupertip = _
        "Auto export all current VBA files in this workbook to default folder"
End Sub

Private Sub setVisible()
    Dim isVisible As Boolean
    Let isVisible = True
    Let toolsTab.letVisible = isVisible
    Let addSheetsButton.letVisible = isVisible
    Let listSheetsButton.letVisible = isVisible
    Let deleteSheetsButton.letVisible = isVisible
    Let showSheetsButton.letVisible = isVisible
    Let hideSheetsButton.letVisible = isVisible
    Let veryHideSheetsButton.letVisible = isVisible
    Let chartHideErrButton.letVisible = isVisible
    Let chartShowButton.letVisible = isVisible
    Let refeshPivotButton.letVisible = isVisible
    Let importVbaFiles.letVisible = isVisible
    Let importAllVbaFiles.letVisible = isVisible
    Let exportAllVbaFiles.letVisible = isVisible
End Sub

Private Function hasWorkPlace() As Boolean
    If Application.ActiveWorkbook Is Nothing Then
        Let hasWorkPlace = False
    ElseIf Application.ActiveSheet Is Nothing Then
        Let hasWorkPlace = False
    Else
        Let hasWorkPlace = True
    End If
End Function
'MAIN
'Callback for customUI.onLoad
Public Sub customUIOnLoad(Optional ByRef ribbon As IRibbonUI)
    Set loadedRibbon = ribbon 'TO-DO: ?
    Call createInstances
    Call setUpId
    Call setUpDescription
    Call setUpEnabled
    Call setUpShowImage
    Call setUpKeytip
    Call setUpLabel
    Call setUpShowLabel
    Call setUpScreentip
    Call setUpSupertip
    Call setVisible
    Let hasCustomUI = True
End Sub
'Callback for add-sheets getEnabled
Public Sub checkEnabled(ByRef control As IRibbonControl, ByRef returnedVal)
    Call setUpEnabled
    Select Case control.ID
        Case addSheetsButton.getID
            Let returnedVal = addSheetsButton.getEnabled
        Case listSheetsButton.getID
            Let returnedVal = listSheetsButton.getEnabled
        Case deleteSheetsButton.getID
            Let returnedVal = deleteSheetsButton.getEnabled
        Case hideSheetsButton.getID
            Let returnedVal = hideSheetsButton.getEnabled
        Case veryHideSheetsButton.getID
            Let returnedVal = veryHideSheetsButton.getEnabled
        Case showSheetsButton.getID
            Let returnedVal = showSheetsButton.getEnabled
        Case chartHideErrButton.getID
            Let returnedVal = chartHideErrButton.getEnabled
        Case chartShowButton.getID
            Let returnedVal = chartShowButton.getEnabled
        Case refeshPivotButton.getID
            Let returnedVal = refeshPivotButton.getEnabled
        Case importVbaFiles.getID
            Let returnedVal = importVbaFiles.getEnabled
        Case importAllVbaFiles.getID
            Let returnedVal = importAllVbaFiles.getEnabled
        Case exportAllVbaFiles.getID
            Let returnedVal = exportAllVbaFiles.getEnabled
    End Select
End Sub

'Callback for add-sheets getDescription
Public Sub createDescription(ByRef control As IRibbonControl, ByRef returnedVal)
    Select Case control.ID
        Case addSheetsButton.getID
            Let returnedVal = addSheetsButton.getDescription
        Case listSheetsButton.getID
            Let returnedVal = listSheetsButton.getDescription
        Case deleteSheetsButton.getID
            Let returnedVal = deleteSheetsButton.getDescription
        Case hideSheetsButton.getID
            Let returnedVal = hideSheetsButton.getDescription
        Case veryHideSheetsButton.getID
            Let returnedVal = veryHideSheetsButton.getDescription
        Case showSheetsButton.getID
            Let returnedVal = showSheetsButton.getDescription
        Case chartHideErrButton.getID
            Let returnedVal = chartHideErrButton.getDescription
        Case chartShowButton.getID
            Let returnedVal = chartShowButton.getDescription
        Case importVbaFiles.getID
            Let returnedVal = importVbaFiles.getDescription
        Case importAllVbaFiles.getID
            Let returnedVal = importAllVbaFiles.getDescription
        Case exportAllVbaFiles.getID
            Let returnedVal = exportAllVbaFiles.getDescription
    End Select
End Sub

'Callback for add-sheets getShowImage
Public Sub showImage(ByRef control As IRibbonControl, ByRef returnedVal)
    Select Case control.ID
        Case addSheetsButton.getID
            Let returnedVal = addSheetsButton.getShowImage
        Case listSheetsButton.getID
            Let returnedVal = listSheetsButton.getShowImage
        Case deleteSheetsButton.getID
            Let returnedVal = deleteSheetsButton.getShowImage
        Case hideSheetsButton.getID
            Let returnedVal = hideSheetsButton.getShowImage
        Case veryHideSheetsButton.getID
            Let returnedVal = veryHideSheetsButton.getShowImage
        Case showSheetsButton.getID
            Let returnedVal = showSheetsButton.getShowImage
        Case chartHideErrButton.getID
            Let returnedVal = chartHideErrButton.getShowImage
        Case chartShowButton.getID
            Let returnedVal = chartShowButton.getShowImage
        Case refeshPivotButton.getID
            Let returnedVal = refeshPivotButton.getShowImage
        Case importVbaFiles.getID
            Let returnedVal = importVbaFiles.getShowImage
        Case importAllVbaFiles.getID
            Let returnedVal = importAllVbaFiles.getShowImage
        Case exportAllVbaFiles.getID
            Let returnedVal = exportAllVbaFiles.getShowImage
    End Select
End Sub

'Callback for add-sheets getKeytip
Public Sub createKeytip(ByRef control As IRibbonControl, ByRef returnedVal)
    Select Case control.ID
        Case addSheetsButton.getID
            Let returnedVal = addSheetsButton.getKeytip
        Case listSheetsButton.getID
            Let returnedVal = listSheetsButton.getKeytip
        Case deleteSheetsButton.getID
            Let returnedVal = deleteSheetsButton.getKeytip
        Case hideSheetsButton.getID
            Let returnedVal = hideSheetsButton.getKeytip
        Case veryHideSheetsButton.getID
            Let returnedVal = veryHideSheetsButton.getKeytip
        Case showSheetsButton.getID
            Let returnedVal = showSheetsButton.getKeytip
        Case chartHideErrButton.getID
            Let returnedVal = chartHideErrButton.getKeytip
        Case chartShowButton.getID
            Let returnedVal = chartShowButton.getKeytip
        Case refeshPivotButton.getID
            Let returnedVal = refeshPivotButton.getKeytip
        Case importVbaFiles.getID
            Let returnedVal = importVbaFiles.getKeytip
        Case importAllVbaFiles.getID
            Let returnedVal = importAllVbaFiles.getKeytip
        Case exportAllVbaFiles.getID
            Let returnedVal = exportAllVbaFiles.getKeytip
    End Select
End Sub

'Callback for add-sheets getLabel
Public Sub createLabel(ByRef control As IRibbonControl, ByRef returnedVal)
    If Not hasCustomUI Then Call customUIOnLoad
    Select Case control.ID
        Case toolsTab.getID
            Let returnedVal = toolsTab.getLabel
        Case sheetGroup.getID
            Let returnedVal = sheetGroup.getLabel
        Case addSheetsButton.getID
            Let returnedVal = addSheetsButton.getLabel
        Case listSheetsButton.getID
            Let returnedVal = listSheetsButton.getLabel
        Case deleteSheetsButton.getID
            Let returnedVal = deleteSheetsButton.getLabel
        Case hideSheetsButton.getID
            Let returnedVal = hideSheetsButton.getLabel
        Case veryHideSheetsButton.getID
            Let returnedVal = veryHideSheetsButton.getLabel
        Case showSheetsButton.getID
            Let returnedVal = showSheetsButton.getLabel
        Case chartGroup.getID
            Let returnedVal = chartGroup.getLabel
        Case chartHideErrButton.getID
            Let returnedVal = chartHideErrButton.getLabel
        Case chartShowButton.getID
            Let returnedVal = chartShowButton.getLabel
        Case pivotGroup.getID
            Let returnedVal = pivotGroup.getLabel
        Case refeshPivotButton.getID
            Let returnedVal = refeshPivotButton.getLabel
        Case vbaFileGroup.getID
            Let returnedVal = vbaFileGroup.getLabel
        Case importVbaFiles.getID
            Let returnedVal = importVbaFiles.getLabel
        Case importAllVbaFiles.getID
            Let returnedVal = importAllVbaFiles.getLabel
        Case exportAllVbaFiles.getID
            Let returnedVal = exportAllVbaFiles.getLabel
        Case infoGroup.getID
            Let returnedVal = infoGroup.getLabel
        Case toolNameLabel.getID
            Let returnedVal = toolNameLabel.getLabel
        Case versionLabel.getID
            Let returnedVal = versionLabel.getLabel
    End Select
End Sub

'Callback for add-sheets getShowLabel
Public Sub showLabel(ByRef control As IRibbonControl, ByRef returnedVal)
    Select Case control.ID
        Case addSheetsButton.getID
            Let returnedVal = addSheetsButton.getShowLabel
        Case listSheetsButton.getID
            Let returnedVal = listSheetsButton.getShowLabel
        Case deleteSheetsButton.getID
            Let returnedVal = deleteSheetsButton.getShowLabel
        Case hideSheetsButton.getID
            Let returnedVal = hideSheetsButton.getShowLabel
        Case veryHideSheetsButton.getID
            Let returnedVal = veryHideSheetsButton.getShowLabel
        Case showSheetsButton.getID
            Let returnedVal = showSheetsButton.getShowLabel
        Case chartHideErrButton.getID
            Let returnedVal = chartHideErrButton.getShowLabel
        Case chartShowButton.getID
            Let returnedVal = chartShowButton.getShowLabel
        Case refeshPivotButton.getID
            Let returnedVal = refeshPivotButton.getShowLabel
        Case importVbaFiles.getID
            Let returnedVal = importVbaFiles.getShowLabel
        Case importAllVbaFiles.getID
            Let returnedVal = importAllVbaFiles.getShowLabel
        Case exportAllVbaFiles.getID
            Let returnedVal = exportAllVbaFiles.getShowLabel
    End Select
End Sub

'Callback for add-sheets getScreentip
Public Sub createScreentip(ByRef control As IRibbonControl, ByRef returnedVal)
    Select Case control.ID
        Case addSheetsButton.getID
            Let returnedVal = addSheetsButton.getScreentip
        Case listSheetsButton.getID
            Let returnedVal = listSheetsButton.getScreentip
        Case deleteSheetsButton.getID
            Let returnedVal = deleteSheetsButton.getScreentip
        Case hideSheetsButton.getID
            Let returnedVal = hideSheetsButton.getScreentip
        Case veryHideSheetsButton.getID
            Let returnedVal = veryHideSheetsButton.getScreentip
        Case showSheetsButton.getID
            Let returnedVal = showSheetsButton.getScreentip
        Case chartHideErrButton.getID
            Let returnedVal = chartHideErrButton.getScreentip
        Case chartShowButton.getID
            Let returnedVal = chartShowButton.getScreentip
        Case refeshPivotButton.getID
            Let returnedVal = refeshPivotButton.getScreentip
        Case importVbaFiles.getID
            Let returnedVal = importVbaFiles.getScreentip
        Case importAllVbaFiles.getID
            Let returnedVal = importAllVbaFiles.getScreentip
        Case exportAllVbaFiles.getID
            Let returnedVal = exportAllVbaFiles.getScreentip
    End Select
End Sub

'Callback for add-sheets getSupertip
Public Sub createSupertip(ByRef control As IRibbonControl, ByRef returnedVal)
    Select Case control.ID
        Case addSheetsButton.getID
            Let returnedVal = addSheetsButton.getSupertip
        Case listSheetsButton.getID
            Let returnedVal = listSheetsButton.getSupertip
        Case deleteSheetsButton.getID
            Let returnedVal = deleteSheetsButton.getSupertip
        Case hideSheetsButton.getID
            Let returnedVal = hideSheetsButton.getSupertip
        Case veryHideSheetsButton.getID
            Let returnedVal = veryHideSheetsButton.getSupertip
        Case showSheetsButton.getID
            Let returnedVal = showSheetsButton.getSupertip
        Case chartHideErrButton.getID
            Let returnedVal = chartHideErrButton.getSupertip
        Case chartShowButton.getID
            Let returnedVal = chartShowButton.getSupertip
        Case refeshPivotButton.getID
            Let returnedVal = refeshPivotButton.getSupertip
        Case importVbaFiles.getID
            Let returnedVal = importVbaFiles.getSupertip
        Case importAllVbaFiles.getID
            Let returnedVal = importAllVbaFiles.getSupertip
        Case exportAllVbaFiles.getID
            Let returnedVal = exportAllVbaFiles.getSupertip
    End Select
End Sub

'Callback for add-sheets getVisible
Public Sub checkVisible(ByRef control As IRibbonControl, ByRef returnedVal)
    If Not hasCustomUI Then Call customUIOnLoad
    Select Case control.ID
        Case toolsTab.getID
            Let returnedVal = toolsTab.getVisible
        Case addSheetsButton.getID
            Let returnedVal = addSheetsButton.getVisible
        Case listSheetsButton.getID
            Let returnedVal = listSheetsButton.getVisible
        Case deleteSheetsButton.getID
            Let returnedVal = deleteSheetsButton.getVisible
        Case hideSheetsButton.getID
            Let returnedVal = hideSheetsButton.getVisible
        Case veryHideSheetsButton.getID
            Let returnedVal = veryHideSheetsButton.getVisible
        Case showSheetsButton.getID
            Let returnedVal = showSheetsButton.getVisible
        Case chartHideErrButton.getID
            Let returnedVal = chartHideErrButton.getVisible
        Case chartShowButton.getID
            Let returnedVal = chartShowButton.getVisible
        Case refeshPivotButton.getID
            Let returnedVal = refeshPivotButton.getVisible
        Case importVbaFiles.getID
            Let returnedVal = importVbaFiles.getVisible
        Case importAllVbaFiles.getID
            Let returnedVal = importAllVbaFiles.getVisible
        Case exportAllVbaFiles.getID
            Let returnedVal = exportAllVbaFiles.getVisible
    End Select
End Sub

'Callback for add-sheets onAction
Public Sub sheetController(ByRef control As IRibbonControl, Optional ByRef pressed As Boolean)
    Set sheetC = New SheetsController
    Select Case control.ID
        Case addSheetsButton.getID
            Call sheetC.add
        Case listSheetsButton.getID
            Call sheetC.list
        Case deleteSheetsButton.getID
            Call sheetC.deleteAll
        Case hideSheetsButton.getID
            Call sheetC.hide( _
                isHide:=True, _
                isVeryHide:=False)
        Case veryHideSheetsButton.getID
            Call sheetC.hide( _
                isHide:=True, _
                isVeryHide:=True)
        Case showSheetsButton.getID
            Call sheetC.hide( _
                isHide:=False, _
                isVeryHide:=False)
    End Select
End Sub

'Callback for hide-error-labels onAction
Public Sub chartController(ByRef control As IRibbonControl, Optional ByRef pressed As Boolean)
    Set chartC = New ChartsController
    Select Case control.ID
        Case chartHideErrButton.getID
            Call chartC.hide( _
                isHide:=True)
        Case chartShowButton.getID
            Call chartC.hide( _
                isHide:=False)
    End Select
    Set chartC = Nothing ' Don't have event so can clear
End Sub

'Callback for refesh-pivot onAction
Public Sub pivotController(ByRef control As IRibbonControl, Optional ByRef pressed As Boolean)
    Select Case control.ID
        Case refeshPivotButton.getID
            If pressed Then
                Set pivotC = New PivotTablesController
            End If
            If Not pressed Then
                Set pivotC = Nothing
            End If
    End Select
End Sub

'Callback for import-vba-files onAction
Public Sub VBAFilesController(ByRef control As IRibbonControl, Optional ByRef pressed As Boolean)
    Set fileC = New FilesController
    Select Case control.ID

    End Select
End Sub 
