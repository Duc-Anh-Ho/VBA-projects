' Check README.md for more information
Option Explicit
'Declare Variables
Private system As SystemUpdate
Private fileSystem As Object
Private info As MyInfo
Private userResponse As VbMsgBoxResult
Private ribbonEvents As customEvents

Private formatC As FormatController 'TO DO
Private sheetCEvent As SheetsController
Private pivotCEvent As PivotTablesController
Private rangeCEvent As RangesController

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
Private importVbaFilesButton As TagController
Private importAllVbaFilesButton As TagController
Private exportAllVbaFilesButton As TagController

Private rangeGroup As TagController
Private hidePageBrakeDropDown As TagController
Private boldFirstLineButton As TagController
Private invertColorButton As TagController
Private highlightButton As TagController
Private highlightBoldButton As TagController
Private highlightSizeNoneButton As TagController
Private highlightSizeOneButton As TagController
Private highlightSizeTwoButton As TagController
Private highlightSizeThreeButton As TagController
Private highlightSizeFourButton As TagController
Private highlightSizeFiveButton As TagController
Private highlightBlurNoneButton As TagController
Private highlightBlurQuarterButton As TagController
Private highlightBlurHalfButton As TagController
Private highlightBlurThreeQuarterButton As TagController
Private highlightBlurFullButton As TagController
Private highlightColorYellowButton As TagController
Private highlightColorCyanButton As TagController
Private highlightColorMagentaButton As TagController
Private highlightColorGreenButton As TagController
Private highlightColorRedButton As TagController
Private highlightColorBlueButton As TagController
Private highlightColorBlackButton As TagController
Private highlightColorWhiteButton As TagController

Private hasPageBeak As Boolean
Private isHighlight As Boolean
Private highlightIsBold As Boolean
Private highlightUpSize As Byte
Private highlightTransparent As Byte
Private highlightColor As Long
Private Const DEFAULT_HIGHLIGHT_UP_SIZE As Byte = 0
Private Const DEFAULT_HIGHLIGHT_TRANSPARENT As Byte = 75
Private Const DEFAULT_HIGHLIGHT_COLOR As Long = vbYellow

Private pictureGroup As TagController
Public arrangeButton As TagController 'Use in ThisWorkbook Module to modify image
Private autoArrangeButton As TagController
Private snipButton As TagController
Private offsetCBBox As TagController
Private rateLockCheckBox As TagController
Public offsetValue As Byte 'Use in ThisWorkbook Module
Public isRateLock As Boolean 'Use in ThisWorkbook Module
Public isArranging As Boolean 'Use in ThisWorkbook Module
Private isAutoArrange As Boolean
Private Const NUM_OFFSET_ITEMS As Byte = 6
Private Const MAX_OFFSET As Byte = 200
Private Const MIN_OFFSET As Byte = 0
Private Const DEFAULT_OFFSET_VALUE  As Byte = 0

Private optionGroup As TagController
Private settingsButton As TagController
Private exportWifiToTxtButton As TagController
Private exportWifiToCsvButton As TagController
Private exportWifiToJsonButton As TagController
Private removeAddinButton As TagController

Private infoGroup As TagController
Private toolNameLabel As TagController
Private versionLabel As TagController

Private hasCustomUI As Boolean
Public loadedRibbon As IRibbonUI 'TO-DO: ?
'Constructor
Private Sub Auto_Open()
    'Add Shortcuts
    Application.OnKey _
        key:="^+{C}", _
        procedure:="Shortcut.copyF"
    Application.OnKey _
        key:="^+{V}", _
        procedure:="Shortcut.pasteF"
End Sub
Private Sub Auto_Activate()
    '
End Sub
'Destructor
Private Sub Auto_Close()
    'Remove Shortcuts
    Application.OnKey _
        key:="^+{C}"
    Application.OnKey _
        key:="^+{V}"
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
    Set chartGroup = New TagController
    Set chartHideErrButton = New TagController
    Set chartShowButton = New TagController
    Set vbaFileGroup = New TagController
    Set importVbaFilesButton = New TagController
    Set importAllVbaFilesButton = New TagController
    Set exportAllVbaFilesButton = New TagController
    Set pivotGroup = New TagController
    Set refeshPivotButton = New TagController
    Set rangeGroup = New TagController
    Set hidePageBrakeDropDown = New TagController
    Set boldFirstLineButton = New TagController
    Set invertColorButton = New TagController
    Set highlightButton = New TagController
    Set highlightBoldButton = New TagController
    Set highlightSizeNoneButton = New TagController
    Set highlightSizeOneButton = New TagController
    Set highlightSizeTwoButton = New TagController
    Set highlightSizeThreeButton = New TagController
    Set highlightSizeFourButton = New TagController
    Set highlightSizeFiveButton = New TagController
    Set highlightBlurNoneButton = New TagController
    Set highlightBlurQuarterButton = New TagController
    Set highlightBlurHalfButton = New TagController
    Set highlightBlurThreeQuarterButton = New TagController
    Set highlightBlurFullButton = New TagController
    Set highlightColorYellowButton = New TagController
    Set highlightColorCyanButton = New TagController
    Set highlightColorMagentaButton = New TagController
    Set highlightColorGreenButton = New TagController
    Set highlightColorRedButton = New TagController
    Set highlightColorBlueButton = New TagController
    Set highlightColorBlackButton = New TagController
    Set highlightColorWhiteButton = New TagController
    Set pictureGroup = New TagController
    Set snipButton = New TagController
    Set arrangeButton = New TagController
    Set autoArrangeButton = New TagController
    Set offsetCBBox = New TagController
    Set rateLockCheckBox = New TagController
    Set optionGroup = New TagController
    Set settingsButton = New TagController
    Set exportWifiToTxtButton = New TagController
    Set exportWifiToCsvButton = New TagController
    Set exportWifiToJsonButton = New TagController
    Set removeAddinButton = New TagController
    Set infoGroup = New TagController
    Set toolNameLabel = New TagController
    Set versionLabel = New TagController
End Sub

Private Sub setUpEnabled()
    Dim isEnabled As Boolean
    Let isEnabled = hasWorkPlace()
    Let addSheetsButton.letEnabled = _
        isEnabled And _
        Not isAutoArrange 'Disable when Highlight and auto Arrange
    Let listSheetsButton.letEnabled = _
        isEnabled And _
        Not isHighlight And _
        Not isAutoArrange
    Let deleteSheetsButton.letEnabled = _
        isEnabled And _
        Not isAutoArrange 'When AutoArrange Disable all
    Let showSheetsButton.letEnabled = _
        isEnabled And _
        Not isAutoArrange
    Let hideSheetsButton.letEnabled = _
        isEnabled And _
        Not isAutoArrange
    Let veryHideSheetsButton.letEnabled = _
        isEnabled And _
        Not isAutoArrange
    Let chartHideErrButton.letEnabled = _
        isEnabled And _
        Not isAutoArrange
    Let chartShowButton.letEnabled = _
        isEnabled And _
        Not isAutoArrange
    Let refeshPivotButton.letEnabled = _
        isEnabled And _
        Not isAutoArrange
    Let importVbaFilesButton.letEnabled = _
        isEnabled And _
        Not isAutoArrange
    Let importAllVbaFilesButton.letEnabled = _
        isEnabled And _
        Not isAutoArrange
    Let exportAllVbaFilesButton.letEnabled = _
        isEnabled And _
        Not isAutoArrange
    Let hidePageBrakeDropDown.letEnabled = _
        isEnabled And _
        Not isAutoArrange
    Let boldFirstLineButton.letEnabled = _
        isEnabled And _
        Not isAutoArrange
    Let invertColorButton.letEnabled = _
        isEnabled And _
        Not isAutoArrange
    Let highlightButton.letEnabled = _
        isEnabled And _
        Not isAutoArrange
    Let arrangeButton.letEnabled = _
        isEnabled And _
        Not isAutoArrange And _
        Not isHighlight
    Let autoArrangeButton.letEnabled = _
        isEnabled And _
        Not isHighlight
    Let snipButton.letEnabled = _
        isEnabled And _
        Not isAutoArrange And _
        Not isHighlight
    Let offsetCBBox.letEnabled = isEnabled
    Let rateLockCheckBox.letEnabled = isEnabled
    Let settingsButton.letEnabled = isEnabled
    Let removeAddinButton.letEnabled = True 'Able to remove without workplace
End Sub
Private Sub setUpImage()
    Let refeshPivotButton.letImage = "ChartRefresh"
    Let arrangeButton.letImage = "SmartArtLargerShape"
    Let autoArrangeButton.letImage = "PicturesCompress"
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
    Let importVbaFilesButton.letShowImage = isShowed
    Let importAllVbaFilesButton.letShowImage = isShowed
    Let exportAllVbaFilesButton.letShowImage = isShowed
    Let hidePageBrakeDropDown.letShowImage = isShowed
    Let boldFirstLineButton.letShowImage = isShowed
    Let invertColorButton.letShowImage = isShowed
    Let highlightButton.letShowImage = isShowed
    Let snipButton.letShowImage = isShowed
    Let arrangeButton.letShowImage = isShowed
    Let autoArrangeButton.letShowImage = isShowed
    Let offsetCBBox.letShowImage = isShowed
    Let settingsButton.letShowImage = isShowed
    Let removeAddinButton.letShowImage = isShowed
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
    Let importVbaFilesButton.letKeytip = ""
    Let importAllVbaFilesButton.letKeytip = ""
    Let exportAllVbaFilesButton.letKeytip = ""
    Let boldFirstLineButton.letKeytip = ""
    Let invertColorButton.letKeytip = ""
    Let highlightButton.letKeytip = ""
    Let snipButton.letKeytip = ""
    Let arrangeButton.letKeytip = ""
    Let autoArrangeButton.letKeytip = ""
    Let settingsButton.letKeytip = ""
    Let removeAddinButton.letKeytip = ""
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
    Let importVbaFilesButton.letShowLabel = isShowed
    Let importAllVbaFilesButton.letShowLabel = isShowed
    Let exportAllVbaFilesButton.letShowLabel = isShowed
    Let hidePageBrakeDropDown.letShowLabel = isShowed
    Let boldFirstLineButton.letShowLabel = isShowed
    Let invertColorButton.letShowLabel = isShowed
    Let highlightButton.letShowLabel = isShowed
    Let snipButton.letShowLabel = isShowed
    Let arrangeButton.letShowLabel = isShowed
    Let autoArrangeButton.letShowLabel = isShowed
    Let offsetCBBox.letShowLabel = isShowed
    Let rateLockCheckBox.letShowLabel = isShowed
    Let settingsButton.letShowLabel = isShowed
    Let removeAddinButton.letShowLabel = isShowed
End Sub

Private Sub setUpVisible()
    Dim isVisible As Boolean
    If Not TypeName(ActiveSheet) = "Chart" Then Let isVisible = True
    If toolsTab Is Nothing Then Call customUIOnLoad
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
    Let importVbaFilesButton.letVisible = isVisible
    Let importAllVbaFilesButton.letVisible = isVisible
    Let exportAllVbaFilesButton.letVisible = isVisible
    Let hidePageBrakeDropDown.letVisible = isVisible
    Let boldFirstLineButton.letVisible = isVisible
    Let invertColorButton.letVisible = isVisible
    Let snipButton.letVisible = isVisible
    Let arrangeButton.letVisible = isVisible
    Let autoArrangeButton.letVisible = isVisible
    Let offsetCBBox.letVisible = isVisible
    Let rateLockCheckBox.letVisible = isVisible
    Let settingsButton.letVisible = isVisible
    Let removeAddinButton.letVisible = isVisible
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
    Let importVbaFilesButton.letID = "import-vba-files"
    Let importAllVbaFilesButton.letID = "import-all-vba-files"
    Let exportAllVbaFilesButton.letID = "export-all-vba-files"
    Let rangeGroup.letID = "ranges-controller"
    Let hidePageBrakeDropDown.letID = "hide-page-break"
    Let boldFirstLineButton.letID = "bold-first-line"
    Let invertColorButton.letID = "invert-color"
    Let highlightButton.letID = "highlight-range"
    Let highlightBoldButton.letID = "highlight-bold"
    Let highlightSizeNoneButton.letID = "highlight-size-none"
    Let highlightSizeOneButton.letID = "highlight-size-one"
    Let highlightSizeTwoButton.letID = "highlight-size-two"
    Let highlightSizeThreeButton.letID = "highlight-size-three"
    Let highlightSizeFourButton.letID = "highlight-size-four"
    Let highlightSizeFiveButton.letID = "highlight-size-five"
    Let highlightBlurNoneButton.letID = "highlight-transparent-none"
    Let highlightBlurQuarterButton.letID = "highlight-transparent-quarter"
    Let highlightBlurHalfButton.letID = "highlight-transparent-half"
    Let highlightBlurThreeQuarterButton.letID = "highlight-transparent-three-quarter"
    Let highlightBlurFullButton.letID = "highlight-transparent-full"
    Let highlightColorYellowButton.letID = "highlight-color-yellow"
    Let highlightColorCyanButton.letID = "highlight-color-cyan"
    Let highlightColorMagentaButton.letID = "highlight-color-magenta"
    Let highlightColorGreenButton.letID = "highlight-color-green"
    Let highlightColorRedButton.letID = "highlight-color-red"
    Let highlightColorBlueButton.letID = "highlight-color-blue"
    Let highlightColorBlackButton.letID = "highlight-color-black"
    Let highlightColorWhiteButton.letID = "highlight-color-white"
    Let settingsButton.letID = "wifi-export-txt"
    Let settingsButton.letID = "settings"
    Let exportWifiToTxtButton.letID = "wifi-export-txt"
    Let exportWifiToCsvButton.letID = "wifi-export-csv"
    Let exportWifiToJsonButton.letID = "wifi-export-json"
    Let removeAddinButton.letID = "remove-addin"
    Let pictureGroup.letID = "pictures-controller"
    Let snipButton.letID = "snipping"
    Let arrangeButton.letID = "arrange"
    Let autoArrangeButton.letID = "auto-arrange"
    Let offsetCBBox.letID = "offset"
    Let rateLockCheckBox.letID = "rate-lock"
    Let optionGroup.letID = "option"
    Let infoGroup.letID = "infomation"
    Let toolNameLabel.letID = "tool-name"
    Let versionLabel.letID = "version"
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
    Let vbaFileGroup.letLabel = "VBA Files"
    Let importVbaFilesButton.letLabel = "Import Files"
    Let importAllVbaFilesButton.letLabel = "Import All Files"
    Let exportAllVbaFilesButton.letLabel = "Export All Files"
    Let hidePageBrakeDropDown.letLabel = "Page Breaks"
    Let boldFirstLineButton.letLabel = "Bold First Line"
    Let invertColorButton.letLabel = "Invert Color"
    Let highlightButton.letLabel = "Highlight Range"
    Let settingsButton.letLabel = "Settings"
    Let removeAddinButton.letLabel = "Remove Addin"
    Let pictureGroup.letLabel = "Picture Controller"
    Let snipButton.letLabel = "Snipping"
    Let arrangeButton.letLabel = "Arrange"
    Let autoArrangeButton.letLabel = "Auto Arrange"
    Let offsetCBBox.letLabel = "Offset"
    Let rateLockCheckBox.letLabel = "Lock The Rate"
    Let optionGroup.letLabel = "Options"
    Let infoGroup.letLabel = "Information"
    Let rangeGroup.letLabel = "Range Controller"
    Let toolNameLabel.letLabel = "Tool: DANH"
    Let versionLabel.letLabel = "Version: v2.0.0"
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
    Let importVbaFilesButton.letScreentip = "Import Files"
    Let importAllVbaFilesButton.letScreentip = "Import All Files"
    Let exportAllVbaFilesButton.letScreentip = "Export All Files"
    Let boldFirstLineButton.letScreentip = "Bold First Line"
    Let invertColorButton.letScreentip = "Invert Color"
    Let highlightButton.letScreentip = "Highlight Range"
    Let settingsButton.letScreentip = "Settings"
    Let removeAddinButton.letScreentip = "Remove Addin"
    Let snipButton.letScreentip = "Snipping"
    Let arrangeButton.letScreentip = "Arrange"
    Let autoArrangeButton.letScreentip = "Auto Arrange"
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
        "Show all labels of a chart."
    Let refeshPivotButton.letSupertip = _
        "Auto update pivot tables." & vbNewLine & _
        "Note: Can't use Undo while this ON. (Update later)"
    Let importVbaFilesButton.letSupertip = _
        "Choose one or many VBA files to import."
    Let importAllVbaFilesButton.letSupertip = _
        "Auto import all VBA files in default folder."
    Let exportAllVbaFilesButton.letSupertip = _
        "Auto export all current VBA files in this workbook to default folder."
    Let boldFirstLineButton.letSupertip = _
        "Auto bold first line of selected cells."
    Let invertColorButton.letSupertip = _
        "Invert color of selected cells."
    Let highlightButton.letSupertip = _
        "1.Highlight column and row of selected range." & vbNewLine & _
        "2.Auto-fit row and column." & vbNewLine & _
        "3.Optional bold, scale up, color and transparent." & vbNewLine & _
        "Note: when OFF will return all format at before ON."
    Let snipButton.letSupertip = _
        "Option 1: Select range and paste the sniped pictures on with offset" & vbNewLine & _
        "Option 2: Select an object and replace or lay the sniped pictures at that object area"
    Let arrangeButton.letSupertip = _
        "1: Select range or object to paste on." & vbNewLine & _
        "2: Select an object to lay on."
    Let autoArrangeButton.letSupertip = _
        "1: After turning on, all shapes will have marker at top left conner." & vbNewLine & _
        "2: Drag shapes'marker into merge range in order to auto arrange." & vbNewLine & _
        "3: Turn off, the marker will automatically disapear." & vbNewLine & _
        "NOTE: Only works on range and merge range!"
    Let settingsButton.letSupertip = _
        "Open settings form."
    Let removeAddinButton.letSupertip = _
        "DANH Tools will be removed permanently."
End Sub

Private Function hasWorkPlace() As Boolean
    If Application.ActiveWorkbook Is Nothing Then
        Let hasWorkPlace = False
    ElseIf Application.ActiveSheet Is Nothing Then
        Let hasWorkPlace = False
    ElseIf TypeName(Application.ActiveSheet) = "Chart" Then
        Let hasWorkPlace = False
    Else
        Let hasWorkPlace = True
        Let hasPageBeak = ActiveSheet.DisplayPageBreaks
    End If
End Function

'Error Handler
Private Sub tackleErrors()
    Select Case Err.Number
        'Default
        Case 0
        'Do nothing
        'Can't load ribbon
        Case 91
            Call customUIOnLoad 'Reset Ribbon
        'Can't enable Addin
        Case 1004
        'Can not import VBA file
        Case 50057
            MsgBox Err.description & _
                vbNewLine & _
                info.getPrompt & _
                "Can not import this file!"
        'FIle import is not VBA file
        Case 50021
            MsgBox Err.description & _
                vbNewLine & _
                info.getPrompt & _
                "This file is not VBA file!"
        'VBA file have password
        Case 50289
            MsgBox _
                Prompt:=Err.description & _
                    vbNewLine & _
                    info.getPrompt & _
                    "Can not access this VBA file because of been protected by a password!", _
                Buttons:=vbCritical + vbOKOnly, _
                Title:=info.getAuthor
        'Un-handled Error
        Case Else
            Call errorDisplay
    End Select
    On Error GoTo 0
End Sub

'Error Handler
Private Sub errorDisplay()
    Dim errorMessage As String
    Let errorMessage = _
        "Error # " & Str(Err.Number) & _
        " was generated by " & Err.Source & _
        vbNewLine & "Error Line: " & Erl & _
        vbNewLine & Err.description
    MsgBox _
        Prompt:=errorMessage, _
        Title:="Error", _
        HelpFile:=Err.HelpFile, _
        Context:=Err.HelpContext
End Sub
'MAIN
'Callback for customUI.onLoad
Public Sub customUIOnLoad(Optional ByRef ribbon As IRibbonUI)
    Set loadedRibbon = ribbon 'TO-DO: ?
    Call createInstances
    Call setUpId
    Call setUpEnabled
    Call setUpImage
    Call setUpShowImage
    Call setUpKeytip
    Call setUpLabel
    Call setUpShowLabel
    Call setUpScreentip
    Call setUpSupertip
    Call setUpVisible
    Let highlightIsBold = False
    Let highlightUpSize = DEFAULT_HIGHLIGHT_UP_SIZE
    Let highlightTransparent = DEFAULT_HIGHLIGHT_TRANSPARENT
    Let highlightColor = DEFAULT_HIGHLIGHT_COLOR
    Let offsetValue = DEFAULT_OFFSET_VALUE
    Let isRateLock = False
    Let isArranging = False
    Let isAutoArrange = False
    If hasWorkPlace Then
        Let hasPageBeak = ActiveSheet.DisplayPageBreaks
    End If
    'Create Custom Event (Ex. Change Sheet,)
    Set ribbonEvents = New customEvents
    Let hasCustomUI = True
End Sub
'Refesh Rebbon
Public Sub refeshCustomRibbon(Optional ByRef rb As IRibbonUI)
    If rb Is Nothing Then Set rb = loadedRibbon
    Call rb.Invalidate
End Sub
'Callback for getImage
Public Sub createImage(ByRef control As IRibbonControl, ByRef returnedVal)
On Error GoTo ErrorHandle
    Select Case control.id
        Case arrangeButton.getID
            Let returnedVal = arrangeButton.getImage
        Case autoArrangeButton.getID
            Let returnedVal = autoArrangeButton.getImage
        Case refeshPivotButton.getID
            Let returnedVal = refeshPivotButton.getImage
    End Select
GoTo ExecuteProcedure
ErrorHandle:
    Call tackleErrors
ExecuteProcedure:
End Sub
'Callback for getEnabled
Public Sub checkEnabled(ByRef control As IRibbonControl, ByRef returnedVal)
On Error GoTo ErrorHandle
    Call setUpEnabled
    Select Case control.id
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
        Case importVbaFilesButton.getID
            Let returnedVal = importVbaFilesButton.getEnabled
        Case importAllVbaFilesButton.getID
            Let returnedVal = importAllVbaFilesButton.getEnabled
        Case exportAllVbaFilesButton.getID
            Let returnedVal = exportAllVbaFilesButton.getEnabled
        Case hidePageBrakeDropDown.getID
            Let returnedVal = hidePageBrakeDropDown.getEnabled
        Case boldFirstLineButton.getID
            Let returnedVal = boldFirstLineButton.getEnabled
        Case invertColorButton.getID
            Let returnedVal = invertColorButton.getEnabled
        Case highlightButton.getID
            Let returnedVal = highlightButton.getEnabled
        Case snipButton.getID
            Let returnedVal = snipButton.getEnabled
        Case arrangeButton.getID
            Let returnedVal = arrangeButton.getEnabled
        Case autoArrangeButton.getID
            Let returnedVal = autoArrangeButton.getEnabled
        Case offsetCBBox.getID
            Let returnedVal = offsetCBBox.getEnabled
        Case rateLockCheckBox.getID
            Let returnedVal = rateLockCheckBox.getEnabled
        Case settingsButton.getID
            Let returnedVal = settingsButton.getEnabled
        Case removeAddinButton.getID
            Let returnedVal = removeAddinButton.getEnabled
    End Select
GoTo ExecuteProcedure
ErrorHandle:
    Call tackleErrors
ExecuteProcedure:
End Sub
'Callback for getShowImage
Public Sub showImage(ByRef control As IRibbonControl, ByRef returnedVal)
On Error GoTo ErrorHandle
    Select Case control.id
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
        Case importVbaFilesButton.getID
            Let returnedVal = importVbaFilesButton.getShowImage
        Case importAllVbaFilesButton.getID
            Let returnedVal = importAllVbaFilesButton.getShowImage
        Case exportAllVbaFilesButton.getID
            Let returnedVal = exportAllVbaFilesButton.getShowImage
        Case hidePageBrakeDropDown.getID
            Let returnedVal = hidePageBrakeDropDown.getShowImage
        Case boldFirstLineButton.getID
            Let returnedVal = boldFirstLineButton.getShowImage
        Case invertColorButton.getID
            Let returnedVal = invertColorButton.getShowImage
        Case highlightButton.getID
            Let returnedVal = highlightButton.getShowImage
        Case snipButton.getID
            Let returnedVal = snipButton.getShowImage
        Case arrangeButton.getID
            Let returnedVal = arrangeButton.getShowImage
        Case autoArrangeButton.getID
            Let returnedVal = autoArrangeButton.getShowImage
        Case offsetCBBox.getID
            Let returnedVal = offsetCBBox.getShowImage
        Case settingsButton.getID
            Let returnedVal = settingsButton.getShowImage
        Case removeAddinButton.getID
            Let returnedVal = removeAddinButton.getShowImage
    End Select
GoTo ExecuteProcedure
ErrorHandle:
    Call tackleErrors
ExecuteProcedure:
End Sub
'Callback for getKeytip
Public Sub createKeytip(ByRef control As IRibbonControl, ByRef returnedVal)
On Error GoTo ErrorHandle
    Select Case control.id
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
        Case importVbaFilesButton.getID
            Let returnedVal = importVbaFilesButton.getKeytip
        Case importAllVbaFilesButton.getID
            Let returnedVal = importAllVbaFilesButton.getKeytip
        Case exportAllVbaFilesButton.getID
            Let returnedVal = exportAllVbaFilesButton.getKeytip
        Case boldFirstLineButton.getID
            Let returnedVal = boldFirstLineButton.getKeytip
        Case invertColorButton.getID
            Let returnedVal = invertColorButton.getKeytip
        Case highlightButton.getID
            Let returnedVal = highlightButton.getKeytip
        Case snipButton.getID
            Let returnedVal = snipButton.getKeytip
        Case arrangeButton.getID
            Let returnedVal = arrangeButton.getKeytip
        Case autoArrangeButton.getID
            Let returnedVal = autoArrangeButton.getKeytip
        Case settingsButton.getID
            Let returnedVal = settingsButton.getKeytip
        Case removeAddinButton.getID
            Let returnedVal = removeAddinButton.getKeytip
    End Select
GoTo ExecuteProcedure
ErrorHandle:
    Call tackleErrors
ExecuteProcedure:
End Sub
'Callback for getLabel
Public Sub createLabel(ByRef control As IRibbonControl, ByRef returnedVal)
On Error GoTo ErrorHandle
    Select Case control.id
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
        Case importVbaFilesButton.getID
            Let returnedVal = importVbaFilesButton.getLabel
        Case importAllVbaFilesButton.getID
            Let returnedVal = importAllVbaFilesButton.getLabel
        Case exportAllVbaFilesButton.getID
            Let returnedVal = exportAllVbaFilesButton.getLabel
        Case rangeGroup.getID
            Let returnedVal = rangeGroup.getLabel
        Case hidePageBrakeDropDown.getID
            Let returnedVal = hidePageBrakeDropDown.getLabel
        Case boldFirstLineButton.getID
            Let returnedVal = boldFirstLineButton.getLabel
        Case invertColorButton.getID
            Let returnedVal = invertColorButton.getLabel
        Case highlightButton.getID
            Let returnedVal = highlightButton.getLabel
        Case settingsButton.getID
            Let returnedVal = settingsButton.getLabel
        Case removeAddinButton.getID
            Let returnedVal = removeAddinButton.getLabel
        Case pictureGroup.getID
            Let returnedVal = pictureGroup.getLabel
        Case snipButton.getID
            Let returnedVal = snipButton.getLabel
        Case arrangeButton.getID
            Let returnedVal = arrangeButton.getLabel
        Case autoArrangeButton.getID
            Let returnedVal = autoArrangeButton.getLabel
        Case offsetCBBox.getID
            Let returnedVal = offsetCBBox.getLabel
        Case rateLockCheckBox.getID
            Let returnedVal = rateLockCheckBox.getLabel
        Case optionGroup.getID
            Let returnedVal = optionGroup.getLabel
        Case infoGroup.getID
            Let returnedVal = infoGroup.getLabel
        Case toolNameLabel.getID
            Let returnedVal = toolNameLabel.getLabel
        Case versionLabel.getID
            Let returnedVal = versionLabel.getLabel
    End Select
GoTo ExecuteProcedure
ErrorHandle:
    Call tackleErrors
ExecuteProcedure:
End Sub
'Callback for getShowLabel
Public Sub showLabel(ByRef control As IRibbonControl, ByRef returnedVal)
On Error GoTo ErrorHandle
    Select Case control.id
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
        Case importVbaFilesButton.getID
            Let returnedVal = importVbaFilesButton.getShowLabel
        Case importAllVbaFilesButton.getID
            Let returnedVal = importAllVbaFilesButton.getShowLabel
        Case exportAllVbaFilesButton.getID
            Let returnedVal = exportAllVbaFilesButton.getShowLabel
        Case hidePageBrakeDropDown.getID
            Let returnedVal = boldFirstLineButton.getShowLabel
        Case boldFirstLineButton.getID
            Let returnedVal = boldFirstLineButton.getShowLabel
        Case invertColorButton.getID
            Let returnedVal = invertColorButton.getShowLabel
        Case highlightButton.getID
            Let returnedVal = highlightButton.getShowLabel
        Case snipButton.getID
            Let returnedVal = snipButton.getShowLabel
        Case arrangeButton.getID
            Let returnedVal = arrangeButton.getShowLabel
        Case autoArrangeButton.getID
            Let returnedVal = autoArrangeButton.getShowLabel
        Case offsetCBBox.getID
            Let returnedVal = offsetCBBox.getShowLabel
        Case rateLockCheckBox.getID
            Let returnedVal = rateLockCheckBox.getShowLabel
        Case settingsButton.getID
            Let returnedVal = settingsButton.getShowLabel
        Case removeAddinButton.getID
            Let returnedVal = removeAddinButton.getShowLabel
    End Select
GoTo ExecuteProcedure
ErrorHandle:
    Call tackleErrors
ExecuteProcedure:
End Sub
'Callback for getScreentip
Public Sub createScreentip(ByRef control As IRibbonControl, ByRef returnedVal)
On Error GoTo ErrorHandle
    Select Case control.id
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
        Case importVbaFilesButton.getID
            Let returnedVal = importVbaFilesButton.getScreentip
        Case importAllVbaFilesButton.getID
            Let returnedVal = importAllVbaFilesButton.getScreentip
        Case exportAllVbaFilesButton.getID
            Let returnedVal = exportAllVbaFilesButton.getScreentip
        Case boldFirstLineButton.getID
            Let returnedVal = boldFirstLineButton.getScreentip
        Case invertColorButton.getID
            Let returnedVal = invertColorButton.getScreentip
        Case highlightButton.getID
            Let returnedVal = highlightButton.getScreentip
        Case snipButton.getID
            Let returnedVal = snipButton.getScreentip
        Case arrangeButton.getID
            Let returnedVal = arrangeButton.getScreentip
        Case autoArrangeButton.getID
            Let returnedVal = autoArrangeButton.getScreentip
        Case settingsButton.getID
            Let returnedVal = settingsButton.getScreentip
        Case removeAddinButton.getID
            Let returnedVal = removeAddinButton.getScreentip
    End Select
GoTo ExecuteProcedure
ErrorHandle:
    Call tackleErrors
ExecuteProcedure:
End Sub
'Callback for getSupertip
Public Sub createSupertip(ByRef control As IRibbonControl, ByRef returnedVal)
On Error GoTo ErrorHandle
    Select Case control.id
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
        Case importVbaFilesButton.getID
            Let returnedVal = importVbaFilesButton.getSupertip
        Case importAllVbaFilesButton.getID
            Let returnedVal = importAllVbaFilesButton.getSupertip
        Case exportAllVbaFilesButton.getID
            Let returnedVal = exportAllVbaFilesButton.getSupertip
        Case boldFirstLineButton.getID
            Let returnedVal = boldFirstLineButton.getSupertip
        Case invertColorButton.getID
            Let returnedVal = invertColorButton.getSupertip
        Case highlightButton.getID
            Let returnedVal = highlightButton.getSupertip
        Case snipButton.getID
            Let returnedVal = snipButton.getSupertip
        Case arrangeButton.getID
            Let returnedVal = arrangeButton.getSupertip
        Case autoArrangeButton.getID
            Let returnedVal = autoArrangeButton.getSupertip
        Case settingsButton.getID
            Let returnedVal = settingsButton.getSupertip
        Case removeAddinButton.getID
            Let returnedVal = removeAddinButton.getSupertip
    End Select
GoTo ExecuteProcedure
ErrorHandle:
    Call tackleErrors
ExecuteProcedure:
End Sub
'Callback for getVisible
Public Sub checkVisible(ByRef control As IRibbonControl, ByRef returnedVal)
On Error GoTo ErrorHandle
    Call setUpVisible
    Select Case control.id
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
        Case importVbaFilesButton.getID
            Let returnedVal = importVbaFilesButton.getVisible
        Case importAllVbaFilesButton.getID
            Let returnedVal = importAllVbaFilesButton.getVisible
        Case exportAllVbaFilesButton.getID
            Let returnedVal = exportAllVbaFilesButton.getVisible
        Case hidePageBrakeDropDown.getID
            Let returnedVal = hidePageBrakeDropDown.getVisible
        Case boldFirstLineButton.getID
            Let returnedVal = boldFirstLineButton.getVisible
        Case invertColorButton.getID
            Let returnedVal = invertColorButton.getVisible
        Case snipButton.getID
            Let returnedVal = snipButton.getVisible
        Case arrangeButton.getID
            Let returnedVal = arrangeButton.getVisible
        Case autoArrangeButton.getID
            Let returnedVal = autoArrangeButton.getVisible
        Case offsetCBBox.getID
            Let returnedVal = offsetCBBox.getVisible
        Case rateLockCheckBox.getID
            Let returnedVal = rateLockCheckBox.getVisible
        Case settingsButton.getID
            Let returnedVal = settingsButton.getVisible
        Case removeAddinButton.getID
            Let returnedVal = removeAddinButton.getVisible
    End Select
GoTo ExecuteProcedure
ErrorHandle:
    Call tackleErrors
ExecuteProcedure:
End Sub
' Check pressed size buttons
Private Function isHighlightUpSize(ByRef value As Byte) As Boolean
    If highlightUpSize = value Then
        Let isHighlightUpSize = True
    Else
        Let isHighlightUpSize = False
    End If
End Function
' Check pressed transparent buttons
Private Function isHighlightTransparent(ByRef value As Byte) As Boolean
    If highlightTransparent = value Then
        Let isHighlightTransparent = True
    Else
        Let isHighlightTransparent = False
    End If
End Function
' Check pressed color buttons
Private Function isHighlightColor(ByRef value As Long) As Boolean
    If highlightColor = value Then
        Let isHighlightColor = True
    Else
        Let isHighlightColor = False
    End If
End Function
'Callback for getPressed
Public Sub checkPressed(control As IRibbonControl, ByRef returnedVal)
On Error GoTo ErrorHandle
    Select Case control.id
        Case highlightBoldButton.getID
            Let returnedVal = highlightIsBold
        Case highlightButton.getID
            Let returnedVal = isHighlight
        Case highlightSizeNoneButton.getID
            Let returnedVal = isHighlightUpSize(0)
        Case highlightSizeOneButton.getID
            Let returnedVal = isHighlightUpSize(1)
        Case highlightSizeTwoButton.getID
            Let returnedVal = isHighlightUpSize(2)
        Case highlightSizeThreeButton.getID
            Let returnedVal = isHighlightUpSize(3)
        Case highlightSizeFourButton.getID
            Let returnedVal = isHighlightUpSize(4)
        Case highlightSizeFiveButton.getID
            Let returnedVal = isHighlightUpSize(5)
        Case highlightBlurNoneButton.getID
            Let returnedVal = isHighlightTransparent(0)
        Case highlightBlurQuarterButton.getID
            Let returnedVal = isHighlightTransparent(25)
        Case highlightBlurHalfButton.getID
            Let returnedVal = isHighlightTransparent(50)
        Case highlightBlurThreeQuarterButton.getID
            Let returnedVal = isHighlightTransparent(75)
        Case highlightBlurFullButton.getID
            Let returnedVal = isHighlightTransparent(100)
        Case highlightColorYellowButton.getID
            Let returnedVal = isHighlightColor(vbYellow)
        Case highlightColorCyanButton.getID
            Let returnedVal = isHighlightColor(vbCyan)
        Case highlightColorMagentaButton.getID
            Let returnedVal = isHighlightColor(vbMagenta)
        Case highlightColorGreenButton.getID
            Let returnedVal = isHighlightColor(vbGreen)
        Case highlightColorRedButton.getID
            Let returnedVal = isHighlightColor(vbRed)
        Case highlightColorBlueButton.getID
            Let returnedVal = isHighlightColor(vbBlue)
        Case highlightColorBlackButton.getID
            Let returnedVal = isHighlightColor(vbBlack)
        Case highlightColorWhiteButton.getID
            Let returnedVal = isHighlightColor(vbWhite)
        Case rateLockCheckBox.getID
            Let returnedVal = isRateLock
        Case arrangeButton.getID
            Let returnedVal = isArranging
        Case autoArrangeButton.getID
            Let returnedVal = isAutoArrange
        Case Else
            Let returnedVal = False
    End Select
GoTo ExecuteProcedure
ErrorHandle:
    Call tackleErrors
ExecuteProcedure:
End Sub
'Callback for Sheet Controller onAction
Public Sub sheetController(ByRef control As IRibbonControl, Optional ByRef pressed As Boolean)
On Error GoTo ErrorHandle
    Dim sheetC As SheetsController
    Set sheetC = New SheetsController
    Select Case control.id
        Case addSheetsButton.getID
            Call sheetC.add
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
    Set sheetC = Nothing
GoTo ExecuteProcedure
ErrorHandle:
    Call tackleErrors
ExecuteProcedure:
End Sub
'Callback for Action With Event
Public Sub sheetControllerEvent(ByRef control As IRibbonControl, Optional ByRef pressed As Boolean)
On Error GoTo ErrorHandle
    Set sheetCEvent = New SheetsController
    Select Case control.id
        Case listSheetsButton.getID
            Call sheetCEvent.list
    End Select
GoTo ExecuteProcedure
ErrorHandle:
    Call tackleErrors
ExecuteProcedure:
End Sub
'Callback for Chart Controller onAction
Public Sub chartController(ByRef control As IRibbonControl, Optional ByRef pressed As Boolean)
On Error GoTo ErrorHandle
    Dim chartC As ChartsController
    Set chartC = New ChartsController
    Select Case control.id
        Case chartHideErrButton.getID
            Call chartC.hide( _
                isHide:=True)
        Case chartShowButton.getID
            Call chartC.hide( _
                isHide:=False)
    End Select
    Set chartC = Nothing ' Clear Cache while don't have event
GoTo ExecuteProcedure
ErrorHandle:
    Call tackleErrors
ExecuteProcedure:
End Sub
'Callback for refesh-pivot onAction
Public Sub pivotControllerEvent(ByRef control As IRibbonControl, Optional ByRef pressed As Boolean)
On Error GoTo ErrorHandle
    Select Case control.id
        Case refeshPivotButton.getID
            If pressed Then
                Set pivotCEvent = New PivotTablesController
                Let refeshPivotButton.letImage = "GroupSyncStatus"
                Let refeshPivotButton.letLabel = "OFF SYNC"
            End If
            If Not pressed Then
                Set pivotCEvent = Nothing
                Let refeshPivotButton.letImage = "ChartRefresh"
                Let refeshPivotButton.letLabel = "SYNC Pivot"
            End If
            Call refeshCustomRibbon(loadedRibbon)
    End Select
GoTo ExecuteProcedure
ErrorHandle:
    Call tackleErrors
ExecuteProcedure:
End Sub
'Callback for VBAFiles Controller onAction
Public Sub VBAFilesController(ByRef control As IRibbonControl, Optional ByRef pressed As Boolean)
On Error GoTo ErrorHandle
    Dim fileC As FilesController
    Set fileC = New FilesController
    Select Case control.id
        Case importVbaFilesButton.getID
            Call fileC.importSelectedVBAfiles
        Case importAllVbaFilesButton.getID
            Call fileC.importAllVbaFiles
        Case exportAllVbaFilesButton.getID
            Call fileC.exportAllVbaFiles
    End Select
    Set fileC = Nothing ' Clear Cache
GoTo ExecuteProcedure
ErrorHandle:
    Call tackleErrors
ExecuteProcedure:
End Sub
'Callback for Range Controller onAction
Public Sub rangeController(ByRef control As IRibbonControl, Optional ByRef pressed As Boolean)
On Error GoTo ErrorHandle
    Dim rangeC As RangesController
    Select Case control.id
        Case boldFirstLineButton.getID
            Set rangeC = New RangesController
            Call rangeC.boldFirstLine
            Set rangeC = Nothing
        Case invertColorButton.getID
            Set rangeC = New RangesController
            Call rangeC.invertColor
            Set rangeC = Nothing
        Case highlightBoldButton.getID
            Let highlightIsBold = pressed
            Call HighlightRange
        Case highlightSizeNoneButton.getID
            If pressed Then Let highlightUpSize = 0
            Call HighlightRange
        Case highlightSizeOneButton.getID
            If pressed Then Let highlightUpSize = 1
            Call HighlightRange
        Case highlightSizeTwoButton.getID
            If pressed Then Let highlightUpSize = 2
            Call HighlightRange
        Case highlightSizeThreeButton.getID
            If pressed Then Let highlightUpSize = 3
            Call HighlightRange
        Case highlightSizeFourButton.getID
            If pressed Then Let highlightUpSize = 4
            Call HighlightRange
        Case highlightSizeFiveButton.getID
            If pressed Then Let highlightUpSize = 5
            Call HighlightRange
        Case highlightBlurNoneButton.getID
            If pressed Then Let highlightTransparent = 0
            Call HighlightRange
        Case highlightBlurQuarterButton.getID
            If pressed Then Let highlightTransparent = 25
            Call HighlightRange
        Case highlightBlurHalfButton.getID
            If pressed Then Let highlightTransparent = 50
            Call HighlightRange
        Case highlightBlurThreeQuarterButton.getID
            If pressed Then Let highlightTransparent = 75
            Call HighlightRange
        Case highlightBlurFullButton.getID
            If pressed Then Let highlightTransparent = 100
            Call HighlightRange
        Case highlightColorYellowButton.getID
            If pressed Then Let highlightColor = vbYellow
            Call HighlightRange
        Case highlightColorCyanButton.getID
            If pressed Then Let highlightColor = vbCyan
            Call HighlightRange
        Case highlightColorMagentaButton.getID
            If pressed Then Let highlightColor = vbMagenta
            Call HighlightRange
        Case highlightColorGreenButton.getID
            If pressed Then Let highlightColor = vbGreen
            Call HighlightRange
        Case highlightColorRedButton.getID
            If pressed Then Let highlightColor = vbRed
            Call HighlightRange
        Case highlightColorBlueButton.getID
            If pressed Then Let highlightColor = vbBlue
            Call HighlightRange
        Case highlightColorBlackButton.getID
            If pressed Then Let highlightColor = vbBlack
            Call HighlightRange
        Case highlightColorWhiteButton.getID
            If pressed Then Let highlightColor = vbWhite
            Call HighlightRange
    End Select
    Call refeshCustomRibbon(loadedRibbon)
GoTo ExecuteProcedure
ErrorHandle:
    Call tackleErrors
ExecuteProcedure:
End Sub

Private Sub HighlightRange()
    If isHighlight Then
        Call rangeCEvent.pasteHighlightFormat ' Refesh screen after clicking button
        Let rangeCEvent.letBold = highlightIsBold
        Let rangeCEvent.letAddSize = highlightUpSize
        Let rangeCEvent.letBlurRate = highlightTransparent
        Let rangeCEvent.letHighlightColor = highlightColor
        Call rangeCEvent.highlight(target:=Selection)
    End If
End Sub

Private Sub ClearHighLight()
    If Not isHighlight Then
        Call rangeCEvent.pasteHighlightFormat
        Set rangeCEvent = Nothing
    End If
End Sub
'Callback for Hide Page Break onAction
Public Sub hidePageBreakChange(control As IRibbonControl, id As String, index As Integer)
On Error GoTo ErrorHandle
    Dim rangeC As RangesController
    Set rangeC = New RangesController
    Select Case control.id
        Case hidePageBrakeDropDown.getID
            Select Case index
                Case 0 'Hide
                    Call rangeC.displayPageBrake( _
                        isDisplay:=False, _
                        isApplyAll:=False)
                Case 1 'Show
                    Call rangeC.displayPageBrake( _
                        isDisplay:=True, _
                        isApplyAll:=False)
                Case 2 'Hide All
                    Call rangeC.displayPageBrake( _
                        isDisplay:=False, _
                        isApplyAll:=True)
                Case 3 'Show All
                    Call rangeC.displayPageBrake( _
                        isDisplay:=True, _
                        isApplyAll:=True)
            End Select
    End Select
    Call refeshCustomRibbon(loadedRibbon)
    Set rangeC = Nothing
GoTo ExecuteProcedure
ErrorHandle:
    Call tackleErrors
ExecuteProcedure:
End Sub
'Callback for Range Controller onAction
Public Sub rangeControllerEvent(ByRef control As IRibbonControl, Optional ByRef pressed As Boolean)
On Error GoTo ErrorHandle
    Select Case control.id
        Case highlightButton.getID
            Let isHighlight = pressed
            If pressed Then
                Set rangeCEvent = New RangesController '*Reduce procedure
                Call rangeCEvent.storeHighlightFormat
                Call HighlightRange
            Else
                Call ClearHighLight
            End If
    End Select
    Call refeshCustomRibbon(loadedRibbon)
GoTo ExecuteProcedure
ErrorHandle:
    Call tackleErrors
ExecuteProcedure:
End Sub
'Callback for settings onAction
Public Sub accessSettings(ByRef control As IRibbonControl, Optional ByRef pressed As Boolean)
On Error GoTo ErrorHandle
    Select Case control.id
        Case settingsButton.getID
            MsgBox _
                Prompt:= _
                    "This button is on process", _
                Buttons:=vbOKOnly + vbExclamation, _
                Title:="DANH TOOL"
    End Select
GoTo ExecuteProcedure
ErrorHandle:
    Call tackleErrors
ExecuteProcedure:
End Sub
'Callback for Internet Controller onAction
Public Sub internetController(ByRef control As IRibbonControl, Optional ByRef pressed As Boolean)
On Error GoTo ErrorHandle
    Dim InternetC As InternetConnect
    Set InternetC = New InternetConnect
    Select Case control.id
        Case exportWifiToTxtButton.getID
            Call InternetC.saveWifiAsTxt
        Case exportWifiToCsvButton.getID
            Call InternetC.saveWifiAsCsv
        Case exportWifiToJsonButton.getID
            Call InternetC.saveWifiAsJson
    End Select
        Set InternetC = Nothing ' Clear Cache
GoTo ExecuteProcedure
ErrorHandle:
    Call tackleErrors
ExecuteProcedure:
End Sub
'Callback for remove-addin onAction
Public Sub addinController(ByRef control As IRibbonControl, Optional ByRef pressed As Boolean)
On Error GoTo ErrorHandle
    Dim newAddin As AutoAddin
    Set newAddin = New AutoAddin
    Select Case control.id
        Case removeAddinButton.getID
            Call newAddin.remove(hasConfirm:=True)
    End Select
    Set newAddin = Nothing ' Clear Cache
GoTo ExecuteProcedure
ErrorHandle:
    Call tackleErrors
ExecuteProcedure:
End Sub
'Callback for remove-addin onAction
Public Sub pictureController(ByRef control As IRibbonControl, Optional ByRef pressed As Boolean)
On Error GoTo ErrorHandle
    Dim picC As PicturesController
    Select Case control.id
        Case snipButton.getID
            Set picC = New PicturesController
            Let picC.letOffset = offsetValue
            Let picC.letLockRatio = isRateLock
            Call picC.snip
            Set picC = Nothing
        Case arrangeButton.getID
            If ActiveSheet.Shapes.Count = 0 Then
                MsgBox _
                    Prompt:= _
                        "This sheet don't exist any object to arrange yet!", _
                    Buttons:=vbOKOnly + vbExclamation, _
                    Title:="DANH TOOL"
                Let isArranging = False
            Else
                Let isArranging = pressed
                Set picC = New PicturesController
                If isArranging Then
                    Let arrangeButton.letImage = "AutomaticResize"
                    Let arrangeButton.letLabel = "Click a shape"
'                    Let picC.letOffset = offsetValue
'                    Let picC.letLockRatio = isRateLock
                    Call picC.assign
                Else
                    Let arrangeButton.letImage = "SmartArtLargerShape"
                    Let arrangeButton.letLabel = "Arrange"
                    Call picC.clearArrange
                End If
                Set picC = Nothing
            End If
            Call refeshCustomRibbon(loadedRibbon)
        Case autoArrangeButton.getID
            If ActiveSheet.Shapes.Count = 0 Then
                MsgBox _
                Prompt:= _
                    "This sheet don't exist any object to AUTO arrange!", _
                Buttons:=vbOKOnly + vbExclamation, _
                Title:="DANH TOOL"
                Let isAutoArrange = False
            Else
                Let isAutoArrange = pressed
                Set picC = New PicturesController
                If isAutoArrange Then
                    Let autoArrangeButton.letImage = "CancelRequest"
                    Let autoArrangeButton.letLabel = "CANCEL"
                    Call picC.autoArrange(True)
                    Let picC.selectObjectMode = True
                    Call ThisWorkbook.Auto_Run_Continuously
                Else
                    Let autoArrangeButton.letImage = "PicturesCompress"
                    Let autoArrangeButton.letLabel = "Auto Arrange"
                    Call picC.autoArrange(False)
                    Let picC.selectObjectMode = False
                End If
                Set picC = Nothing
            End If
            Call refeshCustomRibbon(loadedRibbon)
        Case rateLockCheckBox.getID
            Let isRateLock = pressed
    End Select
GoTo ExecuteProcedure
ErrorHandle:
    Call tackleErrors
ExecuteProcedure:
End Sub
'Callback for offset getItemCount
Public Sub createItemAmount(control As IRibbonControl, ByRef returnedVal)
On Error GoTo ErrorHandle
    Select Case control.id
        Case offsetCBBox.getID
            Let returnedVal = NUM_OFFSET_ITEMS
    End Select
GoTo ExecuteProcedure
ErrorHandle:
    Call tackleErrors
ExecuteProcedure:
End Sub
'Callback for offset getItemID
Public Sub createItemID(control As IRibbonControl, index As Integer, ByRef returnedVal)
On Error GoTo ErrorHandle
    Select Case control.id
        Case offsetCBBox.getID
            Let returnedVal = "offset-item-" & index + 1
    End Select
GoTo ExecuteProcedure
ErrorHandle:
    Call tackleErrors
ExecuteProcedure:
End Sub
'Callback for offset getItemLabel
Public Sub createItemLabel(control As IRibbonControl, index As Integer, ByRef returnedVal)
On Error GoTo ErrorHandle
    Select Case control.id
        Case offsetCBBox.getID
            Let returnedVal = index * 10 'Steps
    End Select
GoTo ExecuteProcedure
ErrorHandle:
    Call tackleErrors
ExecuteProcedure:
End Sub
'Callback for offset getText
Public Sub createText(control As IRibbonControl, ByRef returnedVal)
On Error GoTo ErrorHandle
    Select Case control.id
        Case hidePageBrakeDropDown.getID
            If hasPageBeak Then
                Let returnedVal = 1 'Index = 1 --> "Show"
            Else
                Let returnedVal = 0   'Index = 0 --> "Hide"
            End If
        Case offsetCBBox.getID
            Let returnedVal = offsetValue
    End Select
GoTo ExecuteProcedure
ErrorHandle:
    Call tackleErrors
ExecuteProcedure:
End Sub
'Callback for offset onChange
Public Sub offsetSelect(control As IRibbonControl, text As String)
On Error GoTo ErrorHandle
    Select Case control.id
        Case offsetCBBox.getID
            If Not IsNumeric(text) Then
                MsgBox _
                    Prompt:= _
                        "You have to input a NUMBER between " & _
                        MIN_OFFSET & " and " & MAX_OFFSET & " !", _
                    Buttons:=vbOKOnly + vbCritical, _
                    Title:="DANH TOOL"
                Call refeshCustomRibbon(loadedRibbon)
            ElseIf _
                CLng(text) < MIN_OFFSET Or _
                CLng(text) > MAX_OFFSET Then
                MsgBox _
                    Prompt:= _
                        "Offset value must be between " & _
                        MIN_OFFSET & " and " & MAX_OFFSET & " !", _
                    Buttons:=vbOKOnly + vbCritical, _
                    Title:="DANH TOOL"
                Call refeshCustomRibbon(loadedRibbon)
            Else
                Let offsetValue = CByte(text)
            End If
    End Select
GoTo ExecuteProcedure
ErrorHandle:
    Call tackleErrors
ExecuteProcedure:
End Sub