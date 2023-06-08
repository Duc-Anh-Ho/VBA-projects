Attribute VB_Name = "CustomUi"
' Check README.md for more information
Option Explicit
'Declare Variables
Private Const ALWAYS_SHOW As Boolean = True
Private Const SIZE_LARGE As Boolean = True ' = 1
Private Const SIZE_NORMAL As Boolean = False ' = 0
Private Const DEFAULT_HIGHLIGHT_UP_SIZE As Byte = 0
Private Const DEFAULT_HIGHLIGHT_TRANSPARENT As Byte = 75
Private Const DEFAULT_HIGHLIGHT_COLOR As Long = vbYellow
Private Const NUM_OFFSET_ITEMS As Byte = 6
Private Const MAX_OFFSET As Byte = 200
Private Const MIN_OFFSET As Byte = 0
Private Const DEFAULT_OFFSET_VALUE  As Byte = 0
'Private Const RIBBON_ID As String = "Danh_Tools_Tab_Ribbon_ID_"

Private system As SystemUpdate
Private fileSystem As Object
Private info As InfoConstants
Private userResponse As VbMsgBoxResult
Private ribbonEvents As customEvents
Public loadedRibbon As IRibbonUI
'Public loadedRibbon As Object

Private formatC As FormatController 'TO DO (Shortcut optional/maping)
Private sheetCEvent As SheetsController
Private pivotCEvent As PivotTablesController
Private rangeCEvent As RangesController

Public hasWorksheet As Boolean 'Used in customEvents Class
Public hasWorkChart As Boolean ' nt
Public hasWorkDialog As Boolean 'nt
Public hasSYNCPivot As Boolean 'Used in customEvents Class
Private hasPageBreak As Boolean 'Used in customEvents Class
Public hasHighlight As Boolean 'Used in customEvents Class
Private highlightIsBold As Boolean
Private highlightUpSize As Byte
Private highlightTransparent As Byte
Private highlightColor As Long
Public hasListSheet As Boolean  'Used in customEvents Class
Public hasRenameSheet As Boolean
Public isAutoArrange As Boolean 'Used in customEvents Class
Public isArranging As Boolean 'Used in ThisWorkbook Module And customEvents Class
Public offsetValue As Byte 'Used in ThisWorkbook Module
Public isRateLock As Boolean 'Used in ThisWorkbook Module

Private toolsTab As CustomUITag

Private sheetGroup As CustomUITag
Private addSheetsButton As CustomUITag
Private listSheetsSplit As CustomUITag
Private listSheetsButton As CustomUITag
Private renameSheetsButton As CustomUITag
Private deleteSheetsButton As CustomUITag
Private showSheetsButton As CustomUITag
Private hideSheetsSplit As CustomUITag
Private hideSheetsButton As CustomUITag
Private veryHideSheetsButton As CustomUITag

Private chartGroup As CustomUITag
Private chartHideErrButton As CustomUITag
Private chartShowButton As CustomUITag

Private pivotGroup As CustomUITag
Private refreshPivotButton As CustomUITag

Private vbaFileGroup As CustomUITag
Private importVbaFilesButton As CustomUITag
Private importAllVbaFilesButton As CustomUITag
Private exportAllVbaFilesButton As CustomUITag

Private rangeGroup As CustomUITag
Private hidePageBreakDropDown As CustomUITag
Private boldFirstLineButton As CustomUITag
Private invertColorButton As CustomUITag
Private highlightSplit As CustomUITag
Private highlightButton As CustomUITag
Private highlightBoldButton As CustomUITag
Private highlightSizeNoneButton As CustomUITag
Private highlightSizeOneButton As CustomUITag
Private highlightSizeTwoButton As CustomUITag
Private highlightSizeThreeButton As CustomUITag
Private highlightSizeFourButton As CustomUITag
Private highlightSizeFiveButton As CustomUITag
Private highlightBlurNoneButton As CustomUITag
Private highlightBlurQuarterButton As CustomUITag
Private highlightBlurHalfButton As CustomUITag
Private highlightBlurThreeQuarterButton As CustomUITag
Private highlightBlurFullButton As CustomUITag
Private highlightColorYellowButton As CustomUITag
Private highlightColorCyanButton As CustomUITag
Private highlightColorMagentaButton As CustomUITag
Private highlightColorGreenButton As CustomUITag
Private highlightColorRedButton As CustomUITag
Private highlightColorBlueButton As CustomUITag
Private highlightColorBlackButton As CustomUITag
Private highlightColorWhiteButton As CustomUITag

Private pictureGroup As CustomUITag
Public arrangeButton As CustomUITag 'Use in ThisWorkbook Module to modify image
Private autoArrangeButton As CustomUITag
Private snipButton As CustomUITag
Private offsetCBBox As CustomUITag
Private rateLockCheckBox As CustomUITag

Private optionGroup As CustomUITag
Private settingsSplit As CustomUITag
Private settingsButton As CustomUITag
Private exportWifiToTxtButton As CustomUITag
Private exportWifiToCsvButton As CustomUITag
Private exportWifiToJsonButton As CustomUITag
Private removeAddinButton As CustomUITag
Private refreshAddinButton As CustomUITag

Private infoGroup As CustomUITag
Private toolNameLabel As CustomUITag
Private versionLabel As CustomUITag

'Constructor
Private Sub Auto_Open()
'    Call Shortcuts.install
End Sub

Private Sub Auto_Activate()
    '
End Sub

'Destructor
Private Sub Auto_Close()
'    Call Shortcuts.uninstall
End Sub

Private Sub Auto_Deactivate()
    '
End Sub

'METHODS
Private Sub configTags()
    Dim isOnAutoArrange As Boolean
    Dim isOnListSheet As Boolean
    Dim isOnHighlight As Boolean
    Dim isOnAll As Boolean
    If hasWorksheet Then Let hasPageBreak = ActiveSheet.DisplayPageBreaks
    If hasListSheet Then Let hasRenameSheet = sheetCEvent.hasRenameSheet ' Reset when changed sheets
    Let isOnHighlight = hasWorksheet And Not hasHighlight
    Let isOnAutoArrange = hasWorksheet And Not isAutoArrange
    Let isOnListSheet = hasWorksheet And Not hasListSheet
    Let isOnAll = hasWorksheet _
        And Not isAutoArrange _
        And Not hasHighlight _
        And Not hasListSheet
    Set toolsTab = New CustomUITag
    With toolsTab
        .letEnabled = ALWAYS_SHOW
        .letVisible = ALWAYS_SHOW
        .letShowImage = ALWAYS_SHOW
        .letShowLabel = ALWAYS_SHOW
        .letID = "danh-tools"
        .letLabel = "DANH Tools"
        .letDescription = ""
        .letImage = ""
        .letKeytip = ""
        .letScreentip = ""
        .letSupertip = ""
    End With

    Set sheetGroup = New CustomUITag
    With sheetGroup
        .letEnabled = ALWAYS_SHOW
        .letVisible = ALWAYS_SHOW
        .letShowImage = ALWAYS_SHOW
        .letShowLabel = ALWAYS_SHOW
        .letID = "sheets-controller"
        .letLabel = "Sheet Controller"
        .letDescription = ""
        .letImage = ""
        .letKeytip = ""
        .letScreentip = ""
        .letSupertip = ""
    End With

    Set addSheetsButton = New CustomUITag
    With addSheetsButton
        .letEnabled = isOnAutoArrange
        .letVisible = ALWAYS_SHOW
        .letShowImage = ALWAYS_SHOW
        .letShowLabel = ALWAYS_SHOW
        .letID = "add-sheets"
        .letSize = SIZE_LARGE
        .letLabel = "Add Sheets"
        .letDescription = ""
        .letImage = ""
        .letKeytip = ""
        .letScreentip = "Add Sheets"
        .letSupertip = _
            "1.Select creating name Ranges." & vbNewLine & _
            "2.Click button." & vbNewLine & _
            "3.Auto create Sheets as each cell." & vbNewLine & _
            "Note: Will not working when sheets already exist!"
    End With

    Set listSheetsSplit = New CustomUITag
    With listSheetsSplit
        .letID = "list-sheets-split"
        .letSize = SIZE_LARGE
    End With
    
    Set listSheetsButton = New CustomUITag
    With listSheetsButton
        .letEnabled = isOnHighlight And isOnAutoArrange
        .letVisible = ALWAYS_SHOW
        .letShowImage = ALWAYS_SHOW
        .letShowLabel = ALWAYS_SHOW
        .letID = "list-sheets"
        .letSize = SIZE_LARGE
        .letLabel = "List Sheets"
        .letDescription = ""
        .letImage = ""
        .letKeytip = ""
        .letScreentip = "List Sheets"
        .letSupertip = _
            "ON: List all sheets name at columns A and B." & vbNewLine & _
            "OFF: if ON, turn off lists."
    End With

    Set renameSheetsButton = New CustomUITag
    With renameSheetsButton
        .letID = "rename-sheets"
        .letEnabled = hasListSheet And hasWorksheet
        .letVisible = ALWAYS_SHOW
        .letShowImage = ALWAYS_SHOW
        .letShowLabel = ALWAYS_SHOW
        .letSize = SIZE_LARGE
        .letLabel = "Rename Sheets"
        .letDescription = ""
        .letImage = "TableColumnsInsertRight"
        .letKeytip = ""
        .letScreentip = "Rename Sheets"
        .letSupertip = _
            "1.Press once to show re-name columns C." & vbNewLine & _
            "2.Input new names want to change in columns C." & vbNewLine & _
            "3.Press this button again to apply rename sheets."
    End With

    Set deleteSheetsButton = New CustomUITag
    With deleteSheetsButton
        .letEnabled = isOnAutoArrange
        .letVisible = ALWAYS_SHOW
        .letShowImage = ALWAYS_SHOW
        .letShowLabel = ALWAYS_SHOW
        .letID = "delete-sheets"
        .letSize = SIZE_LARGE
        .letLabel = "Delete Sheets"
        .letDescription = ""
        .letImage = ""
        .letKeytip = ""
        .letScreentip = "Delete Sheets"
        .letSupertip = _
            "1.Delete all sheets except hided sheets." & vbNewLine & _
            "Note: Becareful, after deleted sheets can be undo!"
    End With

    Set showSheetsButton = New CustomUITag
    With showSheetsButton
        .letEnabled = isOnAutoArrange
        .letVisible = ALWAYS_SHOW
        .letShowImage = ALWAYS_SHOW
        .letShowLabel = ALWAYS_SHOW
        .letID = "show-sheets"
        .letSize = SIZE_LARGE
        .letLabel = "Show Sheets"
        .letDescription = ""
        .letImage = ""
        .letKeytip = ""
        .letScreentip = "Show Sheets"
        .letSupertip = _
            "Show all sheets included very hide"
    End With
    
    Set hideSheetsSplit = New CustomUITag
    With hideSheetsSplit
        .letID = "hide-sheets-split"
        .letSize = SIZE_LARGE
    End With

    Set hideSheetsButton = New CustomUITag
    With hideSheetsButton
        .letEnabled = isOnAutoArrange
        .letVisible = ALWAYS_SHOW
        .letShowImage = ALWAYS_SHOW
        .letShowLabel = ALWAYS_SHOW
        .letID = "hide-sheets"
        .letSize = SIZE_LARGE
        .letLabel = "Hide Sheets"
        .letDescription = ""
        .letImage = ""
        .letKeytip = ""
        .letScreentip = "Hide Sheets"
        .letSupertip = _
            "1.Select sheets you don't want to hide." & vbNewLine & _
            "2.Click hide Sheets button." & vbNewLine & _
            "3.Auto hide all un-selected sheets."
    End With

    Set veryHideSheetsButton = New CustomUITag
    With veryHideSheetsButton
        .letEnabled = isOnAutoArrange
        .letVisible = ALWAYS_SHOW
        .letShowImage = ALWAYS_SHOW
        .letShowLabel = ALWAYS_SHOW
        .letID = "very-hide-sheets"
        .letSize = SIZE_LARGE
        .letLabel = "Very Hide Sheet"
        .letDescription = ""
        .letImage = ""
        .letKeytip = ""
        .letScreentip = "Very Hide Sheet"
        .letSupertip = _
            "1.Select sheets you don't want to hide." & vbNewLine & _
            "2.Click hide Sheets button." & vbNewLine & _
            "3.Auto hide all un-selected sheets."
    End With

    Set chartGroup = New CustomUITag
    With chartGroup
        .letEnabled = ALWAYS_SHOW
        .letVisible = ALWAYS_SHOW
        .letShowImage = ALWAYS_SHOW
        .letShowLabel = ALWAYS_SHOW
        .letID = "charts-controller"
        .letLabel = "Chart Controller"
        .letDescription = ""
        .letImage = ""
        .letKeytip = ""
        .letScreentip = "Chart Controller"
        .letSupertip = ""
    End With

    Set chartHideErrButton = New CustomUITag
    With chartHideErrButton
        .letEnabled = isOnAutoArrange
        .letVisible = ALWAYS_SHOW
        .letShowImage = ALWAYS_SHOW
        .letShowLabel = ALWAYS_SHOW
        .letID = "hide-error-labels"
        .letSize = SIZE_LARGE
        .letLabel = "Hide Err Labels"
        .letDescription = ""
        .letImage = ""
        .letKeytip = ""
        .letScreentip = "Hide Error Labels"
        .letSupertip = _
            "1.Select a chart that you want to hide error labels." & vbNewLine & _
            "2.Click Hide Error button." & vbNewLine & _
            "3.Auto very hide all error labels."
    End With

    Set chartShowButton = New CustomUITag
    With chartShowButton
        .letEnabled = isOnAutoArrange
        .letVisible = ALWAYS_SHOW
        .letShowImage = ALWAYS_SHOW
        .letShowLabel = ALWAYS_SHOW
        .letID = "show-labels"
        .letSize = SIZE_LARGE
        .letLabel = "Show Labels"
        .letDescription = ""
        .letImage = "ChartRefresh"
        .letKeytip = ""
        .letScreentip = "Show All Labels"
        .letSupertip = _
            "Show all labels of a chart."
    End With

    Set pivotGroup = New CustomUITag
    With pivotGroup
        .letEnabled = ALWAYS_SHOW
        .letVisible = ALWAYS_SHOW
        .letShowImage = ALWAYS_SHOW
        .letShowLabel = ALWAYS_SHOW
        .letID = "pivot-controller"
        .letLabel = ""
        .letDescription = ""
        .letImage = ""
        .letKeytip = ""
        .letScreentip = ""
        .letSupertip = ""
    End With

    Set refreshPivotButton = New CustomUITag
    With refreshPivotButton
        .letEnabled = isOnAutoArrange
        .letVisible = ALWAYS_SHOW
        .letShowImage = ALWAYS_SHOW
        .letShowLabel = ALWAYS_SHOW
        .letID = "refresh-pivot"
        .letSize = SIZE_LARGE
        .letLabel = "SYNC Pivot"
        .letDescription = ""
        .letImage = "ChartRefresh"
        .letKeytip = ""
        .letScreentip = "SYNC Pivot Table"
        .letSupertip = _
            "Auto update pivot tables." & vbNewLine & _
            "Note: Can't use Undo while this ON. (Update later)"
    End With

    Set vbaFileGroup = New CustomUITag
    With vbaFileGroup
        .letEnabled = ALWAYS_SHOW
        .letVisible = ALWAYS_SHOW
        .letShowImage = ALWAYS_SHOW
        .letShowLabel = ALWAYS_SHOW
        .letID = "vba-files-controller"
        .letLabel = "VBA Files"
        .letDescription = ""
        .letImage = ""
        .letKeytip = ""
        .letScreentip = ""
        .letSupertip = ""
    End With

    Set importVbaFilesButton = New CustomUITag
    With importVbaFilesButton
        .letEnabled = isOnAutoArrange
        .letVisible = ALWAYS_SHOW
        .letShowImage = ALWAYS_SHOW
        .letShowLabel = ALWAYS_SHOW
        .letID = "import-vba-files"
        .letSize = SIZE_LARGE
        .letLabel = "Import Files"
        .letDescription = ""
        .letImage = ""
        .letKeytip = ""
        .letScreentip = "Import VBA Files"
        .letSupertip = _
            "Choose one or many VBA files to import."
    End With

    Set importAllVbaFilesButton = New CustomUITag
    With importAllVbaFilesButton
        .letEnabled = isOnAutoArrange
        .letVisible = ALWAYS_SHOW
        .letShowImage = ALWAYS_SHOW
        .letShowLabel = ALWAYS_SHOW
        .letID = "import-all-vba-files"
        .letSize = SIZE_LARGE
        .letLabel = "Import All Files"
        .letDescription = ""
        .letImage = ""
        .letKeytip = ""
        .letScreentip = "Import All VBA Files"
        .letSupertip = _
            "Auto import all VBA files in default folder."
    End With

    Set exportAllVbaFilesButton = New CustomUITag
    With exportAllVbaFilesButton
        .letEnabled = isOnAutoArrange
        .letVisible = ALWAYS_SHOW
        .letShowImage = ALWAYS_SHOW
        .letShowLabel = ALWAYS_SHOW
        .letID = "export-all-vba-files"
        .letSize = SIZE_LARGE
        .letLabel = "Export All Files"
        .letDescription = ""
        .letImage = ""
        .letKeytip = ""
        .letScreentip = "Export All VBA Files"
        .letSupertip = _
            "Auto export all current VBA files in this workbook to default folder."
    End With

    Set rangeGroup = New CustomUITag
    With rangeGroup
        .letEnabled = ALWAYS_SHOW
        .letVisible = ALWAYS_SHOW
        .letShowImage = ALWAYS_SHOW
        .letShowLabel = ALWAYS_SHOW
        .letID = "ranges-controller"
        .letLabel = "Range Controller"
        .letDescription = ""
        .letImage = ""
        .letKeytip = ""
        .letScreentip = ""
        .letSupertip = ""
    End With

    Set hidePageBreakDropDown = New CustomUITag
    With hidePageBreakDropDown
        .letEnabled = isOnAutoArrange
        .letVisible = ALWAYS_SHOW
        .letShowImage = ALWAYS_SHOW
        .letShowLabel = ALWAYS_SHOW
        .letID = "hide-page-break"
        .letLabel = "Page Breaks"
        .letDescription = ""
        .letImage = ""
        .letKeytip = ""
        .letScreentip = ""
        .letSupertip = ""
    End With

    Set boldFirstLineButton = New CustomUITag
    With boldFirstLineButton
        .letEnabled = isOnAutoArrange
        .letVisible = ALWAYS_SHOW
        .letShowImage = ALWAYS_SHOW
        .letShowLabel = ALWAYS_SHOW
        .letID = "bold-first-line"
        .letSize = SIZE_LARGE
        .letLabel = "Bold First Line"
        .letDescription = ""
        .letImage = ""
        .letKeytip = ""
        .letScreentip = "Bold First Line"
        .letSupertip = _
            "Auto bold first line of selected cells."
    End With

    Set invertColorButton = New CustomUITag
    With invertColorButton
        .letEnabled = isOnAutoArrange
        .letVisible = ALWAYS_SHOW
        .letShowImage = ALWAYS_SHOW
        .letShowLabel = ALWAYS_SHOW
        .letID = "invert-color"
        .letSize = SIZE_LARGE
        .letLabel = "Invert Color"
        .letDescription = ""
        .letImage = ""
        .letKeytip = ""
        .letScreentip = "Invert Color"
        .letSupertip = _
            "Invert color of selected cells."
    End With

    Set highlightSplit = New CustomUITag
    With highlightSplit
        .letID = "highlight-split"
        .letSize = SIZE_LARGE
    End With

    Set highlightButton = New CustomUITag
    With highlightButton
        .letEnabled = isOnAutoArrange And isOnListSheet
        .letVisible = ALWAYS_SHOW
        .letShowImage = ALWAYS_SHOW
        .letShowLabel = ALWAYS_SHOW
        .letID = "highlight-range"
        .letLabel = "Highlight Range"
        .letDescription = ""
        .letImage = ""
        .letKeytip = ""
        .letScreentip = "Highlight Range"
        .letSupertip = _
            "1.Highlight column and row of selected range." & vbNewLine & _
            "2.Auto-fit row and column." & vbNewLine & _
            "3.Optional bold, scale up, color and transparent." & vbNewLine & _
            "Note: when OFF will return all format at before ON."
    End With

    Set highlightBoldButton = New CustomUITag
    With highlightBoldButton
        .letEnabled = ALWAYS_SHOW
        .letVisible = ALWAYS_SHOW
        .letShowImage = ALWAYS_SHOW
        .letShowLabel = ALWAYS_SHOW
        .letID = "highlight-bold"
        .letLabel = ""
        .letDescription = ""
        .letImage = ""
        .letKeytip = ""
        .letScreentip = ""
        .letSupertip = ""
    End With

    Set highlightSizeNoneButton = New CustomUITag
    With highlightSizeNoneButton
        .letEnabled = ALWAYS_SHOW
        .letVisible = ALWAYS_SHOW
        .letShowImage = ALWAYS_SHOW
        .letShowLabel = ALWAYS_SHOW
        .letID = "highlight-size-none"
        .letLabel = ""
        .letDescription = ""
        .letImage = ""
        .letKeytip = ""
        .letScreentip = ""
        .letSupertip = ""
    End With

    Set highlightSizeOneButton = New CustomUITag
    With highlightSizeOneButton
        .letEnabled = ALWAYS_SHOW
        .letVisible = ALWAYS_SHOW
        .letShowImage = ALWAYS_SHOW
        .letShowLabel = ALWAYS_SHOW
        .letID = "highlight-size-one"
        .letLabel = ""
        .letDescription = ""
        .letImage = ""
        .letKeytip = ""
        .letScreentip = ""
        .letSupertip = ""
    End With

    Set highlightSizeTwoButton = New CustomUITag
    With highlightSizeTwoButton
        .letEnabled = ALWAYS_SHOW
        .letVisible = ALWAYS_SHOW
        .letShowImage = ALWAYS_SHOW
        .letShowLabel = ALWAYS_SHOW
        .letID = "highlight-size-two"
        .letLabel = ""
        .letDescription = ""
        .letImage = ""
        .letKeytip = ""
        .letScreentip = ""
        .letSupertip = ""
    End With

    Set highlightSizeThreeButton = New CustomUITag
    With highlightSizeThreeButton
        .letEnabled = ALWAYS_SHOW
        .letVisible = ALWAYS_SHOW
        .letShowImage = ALWAYS_SHOW
        .letShowLabel = ALWAYS_SHOW
        .letID = "highlight-size-three"
        .letLabel = ""
        .letDescription = ""
        .letImage = ""
        .letKeytip = ""
        .letScreentip = ""
        .letSupertip = ""
    End With

    Set highlightSizeFourButton = New CustomUITag
    With highlightSizeFourButton
        .letEnabled = ALWAYS_SHOW
        .letVisible = ALWAYS_SHOW
        .letShowImage = ALWAYS_SHOW
        .letShowLabel = ALWAYS_SHOW
        .letID = "highlight-size-four"
        .letLabel = ""
        .letDescription = ""
        .letImage = ""
        .letKeytip = ""
        .letScreentip = ""
        .letSupertip = ""
    End With

    Set highlightSizeFiveButton = New CustomUITag
    With highlightSizeFiveButton
        .letEnabled = ALWAYS_SHOW
        .letVisible = ALWAYS_SHOW
        .letShowImage = ALWAYS_SHOW
        .letShowLabel = ALWAYS_SHOW
        .letID = "highlight-size-five"
        .letLabel = ""
        .letDescription = ""
        .letImage = ""
        .letKeytip = ""
        .letScreentip = ""
        .letSupertip = ""
    End With

    Set highlightBlurNoneButton = New CustomUITag
    With highlightBlurNoneButton
        .letEnabled = ALWAYS_SHOW
        .letVisible = ALWAYS_SHOW
        .letShowImage = ALWAYS_SHOW
        .letShowLabel = ALWAYS_SHOW
        .letID = "highlight-transparent-none"
        .letLabel = ""
        .letDescription = ""
        .letImage = ""
        .letKeytip = ""
        .letScreentip = ""
        .letSupertip = ""
    End With

    Set highlightBlurQuarterButton = New CustomUITag
    With highlightBlurQuarterButton
        .letEnabled = ALWAYS_SHOW
        .letVisible = ALWAYS_SHOW
        .letShowImage = ALWAYS_SHOW
        .letShowLabel = ALWAYS_SHOW
        .letID = "highlight-transparent-quarter"
        .letLabel = ""
        .letDescription = ""
        .letImage = ""
        .letKeytip = ""
        .letScreentip = ""
        .letSupertip = ""
    End With

    Set highlightBlurHalfButton = New CustomUITag
    With highlightBlurHalfButton
        .letEnabled = ALWAYS_SHOW
        .letVisible = ALWAYS_SHOW
        .letShowImage = ALWAYS_SHOW
        .letShowLabel = ALWAYS_SHOW
        .letID = "highlight-transparent-half"
        .letLabel = ""
        .letDescription = ""
        .letImage = ""
        .letKeytip = ""
        .letScreentip = ""
        .letSupertip = ""
    End With

    Set highlightBlurThreeQuarterButton = New CustomUITag
    With highlightBlurThreeQuarterButton
        .letEnabled = ALWAYS_SHOW
        .letVisible = ALWAYS_SHOW
        .letShowImage = ALWAYS_SHOW
        .letShowLabel = ALWAYS_SHOW
        .letID = "highlight-transparent-three-quarter"
        .letLabel = ""
        .letDescription = ""
        .letImage = ""
        .letKeytip = ""
        .letScreentip = ""
        .letSupertip = ""
    End With

    Set highlightBlurFullButton = New CustomUITag
    With highlightBlurFullButton
        .letEnabled = ALWAYS_SHOW
        .letVisible = ALWAYS_SHOW
        .letShowImage = ALWAYS_SHOW
        .letShowLabel = ALWAYS_SHOW
        .letID = "highlight-transparent-full"
        .letLabel = ""
        .letDescription = ""
        .letImage = ""
        .letKeytip = ""
        .letScreentip = ""
        .letSupertip = ""
    End With

    Set highlightColorYellowButton = New CustomUITag
    With highlightColorYellowButton
        .letEnabled = ALWAYS_SHOW
        .letVisible = ALWAYS_SHOW
        .letShowImage = ALWAYS_SHOW
        .letShowLabel = ALWAYS_SHOW
        .letID = "highlight-color-yellow"
        .letLabel = ""
        .letDescription = ""
        .letImage = ""
        .letKeytip = ""
        .letScreentip = ""
        .letSupertip = ""
    End With

    Set highlightColorCyanButton = New CustomUITag
    With highlightColorCyanButton
        .letEnabled = ALWAYS_SHOW
        .letVisible = ALWAYS_SHOW
        .letShowImage = ALWAYS_SHOW
        .letShowLabel = ALWAYS_SHOW
        .letID = "highlight-color-cyan"
        .letLabel = ""
        .letDescription = ""
        .letImage = ""
        .letKeytip = ""
        .letScreentip = ""
        .letSupertip = ""
    End With

    Set highlightColorMagentaButton = New CustomUITag
    With highlightColorMagentaButton
        .letEnabled = ALWAYS_SHOW
        .letVisible = ALWAYS_SHOW
        .letShowImage = ALWAYS_SHOW
        .letShowLabel = ALWAYS_SHOW
        .letID = "highlight-color-magenta"
        .letLabel = ""
        .letDescription = ""
        .letImage = ""
        .letKeytip = ""
        .letScreentip = ""
        .letSupertip = ""
    End With

    Set highlightColorGreenButton = New CustomUITag
    With highlightColorGreenButton
        .letEnabled = ALWAYS_SHOW
        .letVisible = ALWAYS_SHOW
        .letShowImage = ALWAYS_SHOW
        .letShowLabel = ALWAYS_SHOW
        .letID = "highlight-color-green"
        .letLabel = ""
        .letDescription = ""
        .letImage = ""
        .letKeytip = ""
        .letScreentip = ""
        .letSupertip = ""
    End With

    Set highlightColorRedButton = New CustomUITag
    With highlightColorRedButton
        .letEnabled = ALWAYS_SHOW
        .letVisible = ALWAYS_SHOW
        .letShowImage = ALWAYS_SHOW
        .letShowLabel = ALWAYS_SHOW
        .letID = "highlight-color-red"
        .letLabel = ""
        .letDescription = ""
        .letImage = ""
        .letKeytip = ""
        .letScreentip = ""
        .letSupertip = ""
    End With

    Set highlightColorBlueButton = New CustomUITag
    With highlightColorBlueButton
        .letEnabled = ALWAYS_SHOW
        .letVisible = ALWAYS_SHOW
        .letShowImage = ALWAYS_SHOW
        .letShowLabel = ALWAYS_SHOW
        .letID = "highlight-color-blue"
        .letLabel = ""
        .letDescription = ""
        .letImage = ""
        .letKeytip = ""
        .letScreentip = ""
        .letSupertip = ""
    End With

    Set highlightColorBlackButton = New CustomUITag
    With highlightColorBlackButton
        .letEnabled = ALWAYS_SHOW
        .letVisible = ALWAYS_SHOW
        .letShowImage = ALWAYS_SHOW
        .letShowLabel = ALWAYS_SHOW
        .letID = "highlight-color-black"
        .letLabel = ""
        .letDescription = ""
        .letImage = ""
        .letKeytip = ""
        .letScreentip = ""
        .letSupertip = ""
    End With

    Set highlightColorWhiteButton = New CustomUITag
    With highlightColorWhiteButton
        .letEnabled = ALWAYS_SHOW
        .letVisible = ALWAYS_SHOW
        .letShowImage = ALWAYS_SHOW
        .letShowLabel = ALWAYS_SHOW
        .letID = "highlight-color-white"
        .letLabel = ""
        .letDescription = ""
        .letImage = ""
        .letKeytip = ""
        .letScreentip = ""
        .letSupertip = ""
    End With

    Set pictureGroup = New CustomUITag
    With pictureGroup
        .letEnabled = ALWAYS_SHOW
        .letVisible = ALWAYS_SHOW
        .letShowImage = ALWAYS_SHOW
        .letShowLabel = ALWAYS_SHOW
        .letID = "pictures-controller"
        .letLabel = "Picture Controller"
        .letDescription = ""
        .letImage = ""
        .letKeytip = ""
        .letScreentip = ""
        .letSupertip = ""
    End With

    Set arrangeButton = New CustomUITag
    With arrangeButton
        .letEnabled = isOnAll
        .letVisible = ALWAYS_SHOW
        .letShowImage = ALWAYS_SHOW
        .letShowLabel = ALWAYS_SHOW
        .letID = "arrange"
        .letSize = SIZE_LARGE
        .letLabel = "Arrange"
        .letDescription = ""
        .letImage = "SmartArtLargerShape"
        .letKeytip = ""
        .letScreentip = "Arrange"
        .letSupertip = _
            "1: Select range or object to paste on." & vbNewLine & _
            "2: Select an object to lay on."
    End With

    Set autoArrangeButton = New CustomUITag
    With autoArrangeButton
        .letEnabled = isOnHighlight And isOnListSheet
        .letVisible = ALWAYS_SHOW
        .letShowImage = ALWAYS_SHOW
        .letShowLabel = ALWAYS_SHOW
        .letID = "auto-arrange"
        .letSize = SIZE_LARGE
        .letLabel = "Auto Arrange"
        .letDescription = ""
        .letImage = "PicturesCompress"
        .letKeytip = ""
        .letScreentip = "Auto Arrange"
        .letSupertip = _
            "1: After turning on, all shapes will have marker at top left conner." & vbNewLine & _
            "2: Drag shapes'marker into merge range in order to auto arrange." & vbNewLine & _
            "3: Turn off, the marker will automatically disapear." & vbNewLine & _
            "NOTE: Only works on range and merge range!"
    End With

    Set snipButton = New CustomUITag
    With snipButton
        .letEnabled = isOnAll
        .letVisible = ALWAYS_SHOW
        .letShowImage = ALWAYS_SHOW
        .letShowLabel = ALWAYS_SHOW
        .letID = "snipping"
        .letSize = SIZE_LARGE
        .letLabel = "Snip"
        .letDescription = ""
        .letImage = ""
        .letKeytip = ""
        .letScreentip = "Snip Tool"
        .letSupertip = _
            "1: After turning on, all shapes will have marker at top left conner." & vbNewLine & _
            "2: Drag shapes'marker into merge range in order to auto arrange." & vbNewLine & _
            "3: Turn off, the marker will automatically disapear." & vbNewLine & _
            "NOTE: Only works on range and merge range!"
    End With

    Set offsetCBBox = New CustomUITag
    With offsetCBBox
        .letEnabled = hasWorksheet
        .letVisible = ALWAYS_SHOW
        .letShowImage = ALWAYS_SHOW
        .letShowLabel = ALWAYS_SHOW
        .letID = "offset"
        .letLabel = "Offset"
        .letDescription = ""
        .letImage = ""
        .letKeytip = ""
        .letScreentip = ""
        .letSupertip = ""
    End With

    Set rateLockCheckBox = New CustomUITag
    With rateLockCheckBox
        .letEnabled = hasWorksheet
        .letVisible = ALWAYS_SHOW
        .letShowImage = ALWAYS_SHOW
        .letShowLabel = ALWAYS_SHOW
        .letID = "rate-lock"
        .letLabel = "Lock The Rate"
        .letDescription = ""
        .letImage = ""
        .letKeytip = ""
        .letScreentip = ""
        .letSupertip = ""
    End With

    Set optionGroup = New CustomUITag
    With optionGroup
        .letEnabled = ALWAYS_SHOW
        .letVisible = ALWAYS_SHOW
        .letShowImage = ALWAYS_SHOW
        .letShowLabel = ALWAYS_SHOW
        .letID = "option"
        .letLabel = "Options"
        .letDescription = ""
        .letImage = ""
        .letKeytip = ""
        .letScreentip = ""
        .letSupertip = ""
    End With

    Set settingsSplit = New CustomUITag
    With settingsSplit
        .letID = "settings-split"
        .letSize = SIZE_LARGE
    End With

    Set settingsButton = New CustomUITag
    With settingsButton
        .letEnabled = hasWorksheet
        .letVisible = ALWAYS_SHOW
        .letShowImage = ALWAYS_SHOW
        .letShowLabel = ALWAYS_SHOW
        .letID = "settings"
        .letLabel = "Settings"
        .letDescription = ""
        .letImage = ""
        .letKeytip = ""
        .letScreentip = "Settings"
        .letSupertip = _
            "Open settings form."
    End With

    Set exportWifiToTxtButton = New CustomUITag
    With exportWifiToTxtButton
        .letEnabled = ALWAYS_SHOW
        .letVisible = ALWAYS_SHOW
        .letShowImage = ALWAYS_SHOW
        .letShowLabel = ALWAYS_SHOW
        .letID = "wifi-export-txt"
        .letLabel = ""
        .letDescription = ""
        .letImage = ""
        .letKeytip = ""
        .letScreentip = ""
        .letSupertip = ""
    End With

    Set exportWifiToCsvButton = New CustomUITag
    With exportWifiToCsvButton
        .letEnabled = ALWAYS_SHOW
        .letVisible = ALWAYS_SHOW
        .letShowImage = ALWAYS_SHOW
        .letShowLabel = ALWAYS_SHOW
        .letID = "wifi-export-csv"
        .letLabel = ""
        .letDescription = ""
        .letImage = ""
        .letKeytip = ""
        .letScreentip = ""
        .letSupertip = ""
    End With

    Set exportWifiToJsonButton = New CustomUITag
    With exportWifiToJsonButton
        .letEnabled = ALWAYS_SHOW
        .letVisible = ALWAYS_SHOW
        .letShowImage = ALWAYS_SHOW
        .letShowLabel = ALWAYS_SHOW
        .letID = "wifi-export-json"
        .letLabel = ""
        .letDescription = ""
        .letImage = ""
        .letKeytip = ""
        .letScreentip = ""
        .letSupertip = ""
    End With

    Set removeAddinButton = New CustomUITag
    With removeAddinButton
        .letEnabled = ALWAYS_SHOW
        .letVisible = ALWAYS_SHOW
        .letShowImage = ALWAYS_SHOW
        .letShowLabel = ALWAYS_SHOW
        .letID = "remove-addin"
        .letSize = SIZE_LARGE
        .letLabel = "Remove Addin"
        .letDescription = ""
        .letImage = ""
        .letKeytip = ""
        .letScreentip = "Remove Addin"
        .letSupertip = _
            "DANH Tools will be removed permanently."
    End With

    Set refreshAddinButton = New CustomUITag
    With refreshAddinButton
        .letEnabled = ALWAYS_SHOW
        .letVisible = ALWAYS_SHOW
        .letShowImage = ALWAYS_SHOW
        .letShowLabel = ALWAYS_SHOW
        .letID = "refresh-ribbon"
        .letSize = SIZE_LARGE
        .letLabel = "Refresh Ribbon"
        .letDescription = ""
        .letImage = ""
        .letKeytip = ""
        .letScreentip = "Refresh Ribbon Tab"
        .letSupertip = _
            "When DANH Tools tab is disabled or some unknown errors happened." & vbNewLine & _
            "Just try this button."
    End With
    Set infoGroup = New CustomUITag
    With infoGroup
        .letEnabled = ALWAYS_SHOW
        .letVisible = ALWAYS_SHOW
        .letShowImage = ALWAYS_SHOW
        .letShowLabel = ALWAYS_SHOW
        .letID = "infomation"
        .letLabel = "Information"
        .letDescription = ""
        .letImage = ""
        .letKeytip = ""
        .letScreentip = ""
        .letSupertip = ""
    End With

    Set toolNameLabel = New CustomUITag
    With toolNameLabel
        .letEnabled = ALWAYS_SHOW
        .letVisible = ALWAYS_SHOW
        .letShowImage = ALWAYS_SHOW
        .letShowLabel = ALWAYS_SHOW
        .letID = "tool-name"
        .letLabel = "Tool: DANH"
        .letDescription = ""
        .letImage = ""
        .letKeytip = ""
        .letScreentip = ""
        .letSupertip = ""
    End With

    Set versionLabel = New CustomUITag
    With versionLabel
        .letEnabled = ALWAYS_SHOW
        .letVisible = ALWAYS_SHOW
        .letShowImage = ALWAYS_SHOW
        .letShowLabel = ALWAYS_SHOW
        .letID = "version"
        .letLabel = "Version: " & info.getVersion
        .letDescription = ""
        .letImage = ""
        .letKeytip = ""
        .letScreentip = ""
        .letSupertip = ""
    End With
    
    Call setButtonAsMode
    Call setStatusBarAsMode
    
End Sub
' Override Labels & Images of buttons as modes
Private Sub setButtonAsMode()
    If hasSYNCPivot Then
        Let refreshPivotButton.letImage = "GroupSyncStatus"
        Let refreshPivotButton.letLabel = "OFF SYNC"
    Else
        Let refreshPivotButton.letImage = "ChartRefresh"
        Let refreshPivotButton.letLabel = "SYNC Pivot"
    End If
    If isArranging Then
        Let arrangeButton.letImage = "AutomaticResize"
        Let arrangeButton.letLabel = "Click a shape"
    Else
        Let arrangeButton.letImage = "SmartArtLargerShape"
        Let arrangeButton.letLabel = "Arrange"
    End If
    If isAutoArrange Then
        Let autoArrangeButton.letImage = "CancelRequest"
        Let autoArrangeButton.letLabel = "CANCEL"
    Else
        Let autoArrangeButton.letImage = "PicturesCompress"
        Let autoArrangeButton.letLabel = "Auto Arrange"
    End If
    If hasRenameSheet Then
        Let renameSheetsButton.letImage = "TagMarkComplete"
        Let renameSheetsButton.letLabel = "Confirm"
    Else
        Let renameSheetsButton.letImage = "TableColumnsInsertRight"
        Let renameSheetsButton.letLabel = "Rename Sheets"
    End If
    If Not hasListSheet Then
        Let renameSheetsButton.letImage = "TableColumnsInsertRight"
        Let renameSheetsButton.letLabel = "Rename Sheets"
        Let hasRenameSheet = False ' Reset rename
    End If
End Sub

' Display Application Status Bar as Modes
Private Sub setStatusBarAsMode()
    Dim content As String
    Let content = "Mode: || "
    If hasSYNCPivot Then Let content = content & "AUTO SYNC PIVOT || "
    If hasHighlight Then Let content = content & "AUTO HIGHLIGHT || "
    If isArranging Then Let content = content & "ARRANGING OBJECT || "
    If isAutoArrange Then Let content = content & "AUTO ARRANGE ONJECTS || "
    If hasListSheet Then Let content = content & "LIST SHEETS || "
    If hasRenameSheet Then Let content = content & "RANAME SHEETS || "
    If content <> "Mode: || " Then
        Let system.setStatusBar = content
    Else
        Let system.setStatusBar = False
    End If
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

Private Sub HighlightRange()
    If hasHighlight Then
        Call rangeCEvent.pasteHighlightFormat ' refresh screen after clicking button
        Let rangeCEvent.letBold = highlightIsBold
        Let rangeCEvent.letAddSize = highlightUpSize
        Let rangeCEvent.letBlurRate = highlightTransparent
        Let rangeCEvent.letHighlightColor = highlightColor
        Call rangeCEvent.highlight(Target:=Selection)
    End If
End Sub

Private Sub ClearHighLight()
    If Not hasHighlight And Not rangeCEvent Is Nothing Then
        Call rangeCEvent.pasteHighlightFormat
        Set rangeCEvent = Nothing
    End If
End Sub

'Error Handler
Private Sub tackleErrors()
    Select Case Err.Number
        'Default
        Case 0
        'Do nothing
        'Can't load ribbon
        Case 91
            Call refreshCustomRibbon(loadedRibbon) 'Reset Ribbon
        'Can't enable Addin
        Case 1004
        'Can not import VBA file
        Case 50057
            Let userResponse = MsgBox( _
                Err.description & _
                vbNewLine & _
                info.getPrompt & _
                "Can not import this file!")
        'FIle import is not VBA file
        Case 50021
            Let userResponse = MsgBox( _
                Err.description & _
                vbNewLine & _
                info.getPrompt & _
                "This file is not VBA file!")
        'VBA file have password
        Case 50289
            Let userResponse = MsgBox( _
                Prompt:= _
                    Err.description & _
                    vbNewLine & _
                    info.getPrompt & _
                    "Can not access this VBA file because of been protected by a password!", _
                Buttons:=vbCritical + vbOKOnly, _
                Title:=info.getAuthor)
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
    Let userResponse = MsgBox( _
        Prompt:=errorMessage, _
        Title:="Error", _
        HelpFile:=Err.HelpFile, _
        Context:=Err.HelpContext)
End Sub

Public Sub setDefaultSettings()
    Set ribbonEvents = New customEvents
    Set sheetCEvent = Nothing
    Set pivotCEvent = Nothing
    Set rangeCEvent = Nothing
    Let hasWorksheet = system.hasWorkPlace(False, "xlWorksheet")
    Let hasWorkChart = system.hasWorkPlace(False, "Chart")
    Let hasWorkDialog = system.hasWorkPlace(False, "DialogSheet")
'    system.ws.Unprotect
    Let highlightIsBold = False
    Let highlightUpSize = DEFAULT_HIGHLIGHT_UP_SIZE
    Let highlightTransparent = DEFAULT_HIGHLIGHT_TRANSPARENT
    Let highlightColor = DEFAULT_HIGHLIGHT_COLOR
    Let offsetValue = DEFAULT_OFFSET_VALUE
    Let isRateLock = False
    If hasWorksheet Then Let hasPageBreak = ActiveSheet.DisplayPageBreaks
    'Reset Event Flags
    Let isArranging = False
    Let isAutoArrange = False
    Let hasListSheet = False
    Let hasRenameSheet = False
    Let hasSYNCPivot = False
    Let hasHighlight = False
End Sub

'Refresh Ribbon
Public Sub refreshCustomRibbon(Optional ByRef rb As IRibbonUI)
On Error GoTo ErrorHandle
    Set system = New SystemUpdate
    Set info = New InfoConstants
    If loadedRibbon Is Nothing Then
        Debug.Print ("loadedRibbon Is Nothing") 'For watching debug
        ' Reload Ribbon from Pointer (Preprocessor.GetRibbon is Public)
'        Set loadedRibbon = GetRibbon(Workbooks(info.getAddinName).Names(RIBBON_ID))
        Set loadedRibbon = GetRibbon(Workbooks(info.getAddinName).Names(info.getRibbonID))
        ' Reload setting
        Call setDefaultSettings
    End If
    If rb Is Nothing Then Set rb = loadedRibbon
    Call configTags
    Call rb.Invalidate
GoTo ExecuteProcedure
ErrorHandle:
    Call tackleErrors
ExecuteProcedure:
End Sub

'MAIN
'Callback for CustomUI.onLoad
Public Sub CustomUIOnLoad(Optional ByRef ribbon As IRibbonUI)
On Error GoTo ErrorHandle
    Set system = New SystemUpdate
    Set info = New InfoConstants
    'Store ribbon Object to Public variable
    Set loadedRibbon = ribbon
    'Store pointer to IRibbonUI in a Named Range within add-in file
    Workbooks(info.getAddinName).Names.add _
        Name:=info.getRibbonID, _
        RefersTo:=ObjPtr(ribbon)
    'Create Custom Event (Ex. Change Sheet,)
    Call setDefaultSettings
    Call configTags
GoTo ExecuteProcedure
ErrorHandle:
    Call tackleErrors
ExecuteProcedure:
End Sub

'Callback for getSize
Public Sub getSize(ByRef control As IRibbonControl, ByRef returnedVal)
On Error GoTo ErrorHandle
    If loadedRibbon Is Nothing Then Call refreshCustomRibbon
    Select Case control.id
        Case addSheetsButton.getID
            Let returnedVal = addSheetsButton.getSize
        Case listSheetsSplit.getID
            Let returnedVal = listSheetsSplit.getSize
        Case listSheetsButton.getID
            Let returnedVal = listSheetsButton.getSize
        Case renameSheetsButton.getID
            Let returnedVal = renameSheetsButton.getSize
        Case deleteSheetsButton.getID
            Let returnedVal = deleteSheetsButton.getSize
        Case showSheetsButton.getID
            Let returnedVal = showSheetsButton.getSize
        Case hideSheetsSplit.getID
            Let returnedVal = hideSheetsSplit.getSize
        Case hideSheetsButton.getID
            Let returnedVal = hideSheetsButton.getSize
        Case veryHideSheetsButton.getID
            Let returnedVal = veryHideSheetsButton.getSize
        Case chartHideErrButton.getID
            Let returnedVal = chartHideErrButton.getSize
        Case chartShowButton.getID
            Let returnedVal = chartShowButton.getSize
        Case refreshPivotButton.getID
            Let returnedVal = refreshPivotButton.getSize
        Case importVbaFilesButton.getID
            Let returnedVal = importVbaFilesButton.getSize
        Case importAllVbaFilesButton.getID
            Let returnedVal = importAllVbaFilesButton.getSize
        Case exportAllVbaFilesButton.getID
            Let returnedVal = exportAllVbaFilesButton.getSize
        Case boldFirstLineButton.getID
            Let returnedVal = boldFirstLineButton.getSize
        Case invertColorButton.getID
            Let returnedVal = invertColorButton.getSize
        Case highlightSplit.getID
            Let returnedVal = highlightSplit.getSize
        Case arrangeButton.getID
            Let returnedVal = arrangeButton.getSize
        Case autoArrangeButton.getID
            Let returnedVal = autoArrangeButton.getSize
        Case snipButton.getID
            Let returnedVal = snipButton.getSize
        Case settingsSplit.getID
            Let returnedVal = settingsSplit.getSize
        Case removeAddinButton.getID
            Let returnedVal = removeAddinButton.getSize
        Case refreshAddinButton.getID
            Let returnedVal = refreshAddinButton.getSize
    End Select
GoTo ExecuteProcedure
ErrorHandle:
    Call tackleErrors
ExecuteProcedure:
End Sub

'Callback for getImage
Public Sub getImage(ByRef control As IRibbonControl, ByRef returnedVal)
On Error GoTo ErrorHandle
    If loadedRibbon Is Nothing Then Call refreshCustomRibbon
    Select Case control.id
        Case renameSheetsButton.getID
            Let returnedVal = renameSheetsButton.getImage
        Case arrangeButton.getID
            Let returnedVal = arrangeButton.getImage
        Case autoArrangeButton.getID
            Let returnedVal = autoArrangeButton.getImage
        Case refreshPivotButton.getID
            Let returnedVal = refreshPivotButton.getImage
    End Select
GoTo ExecuteProcedure
ErrorHandle:
    Call tackleErrors
ExecuteProcedure:
End Sub

'Callback for getEnabled
Public Sub getEnabled(ByRef control As IRibbonControl, ByRef returnedVal)
On Error GoTo ErrorHandle
    If loadedRibbon Is Nothing Then Call refreshCustomRibbon
'    Call setUpEnabled
    Select Case control.id
        Case addSheetsButton.getID
            Let returnedVal = addSheetsButton.getEnabled
        Case listSheetsButton.getID
            Let returnedVal = listSheetsButton.getEnabled
        Case renameSheetsButton.getID
            Let returnedVal = renameSheetsButton.getEnabled
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
        Case refreshPivotButton.getID
            Let returnedVal = refreshPivotButton.getEnabled
        Case importVbaFilesButton.getID
            Let returnedVal = importVbaFilesButton.getEnabled
        Case importAllVbaFilesButton.getID
            Let returnedVal = importAllVbaFilesButton.getEnabled
        Case exportAllVbaFilesButton.getID
            Let returnedVal = exportAllVbaFilesButton.getEnabled
        Case hidePageBreakDropDown.getID
            Let returnedVal = hidePageBreakDropDown.getEnabled
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
        Case refreshAddinButton.getID
            Let returnedVal = refreshAddinButton.getEnabled
    End Select
GoTo ExecuteProcedure
ErrorHandle:
    Call tackleErrors
ExecuteProcedure:
End Sub

'Callback for getShowImage
Public Sub getShowImage(ByRef control As IRibbonControl, ByRef returnedVal)
On Error GoTo ErrorHandle
    If loadedRibbon Is Nothing Then Call refreshCustomRibbon
    Select Case control.id
        Case addSheetsButton.getID
            Let returnedVal = addSheetsButton.getShowImage
        Case listSheetsButton.getID
            Let returnedVal = listSheetsButton.getShowImage
        Case renameSheetsButton.getID
            Let returnedVal = renameSheetsButton.getShowImage
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
        Case refreshPivotButton.getID
            Let returnedVal = refreshPivotButton.getShowImage
        Case importVbaFilesButton.getID
            Let returnedVal = importVbaFilesButton.getShowImage
        Case importAllVbaFilesButton.getID
            Let returnedVal = importAllVbaFilesButton.getShowImage
        Case exportAllVbaFilesButton.getID
            Let returnedVal = exportAllVbaFilesButton.getShowImage
        Case hidePageBreakDropDown.getID
            Let returnedVal = hidePageBreakDropDown.getShowImage
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
        Case refreshAddinButton.getID
            Let returnedVal = refreshAddinButton.getShowImage
    End Select
GoTo ExecuteProcedure
ErrorHandle:
    Call tackleErrors
ExecuteProcedure:
End Sub

'Callback for getKeytip
Public Sub getKeytip(ByRef control As IRibbonControl, ByRef returnedVal)
On Error GoTo ErrorHandle
    If loadedRibbon Is Nothing Then Call refreshCustomRibbon
    Select Case control.id
        Case addSheetsButton.getID
            Let returnedVal = addSheetsButton.getKeytip
        Case listSheetsButton.getID
            Let returnedVal = listSheetsButton.getKeytip
        Case renameSheetsButton.getID
            Let returnedVal = renameSheetsButton.getKeytip
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
        Case refreshPivotButton.getID
            Let returnedVal = refreshPivotButton.getKeytip
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
        Case refreshAddinButton.getID
            Let returnedVal = refreshAddinButton.getKeytip
    End Select
GoTo ExecuteProcedure
ErrorHandle:
    Call tackleErrors
ExecuteProcedure:
End Sub

'Callback for getLabel
Public Sub getLabel(ByRef control As IRibbonControl, ByRef returnedVal)
On Error GoTo ErrorHandle
    If loadedRibbon Is Nothing Then Call refreshCustomRibbon
    Select Case control.id
        Case toolsTab.getID
            Let returnedVal = toolsTab.getLabel
        Case sheetGroup.getID
            Let returnedVal = sheetGroup.getLabel
        Case addSheetsButton.getID
            Let returnedVal = addSheetsButton.getLabel
        Case listSheetsButton.getID
            Let returnedVal = listSheetsButton.getLabel
        Case renameSheetsButton.getID
            Let returnedVal = renameSheetsButton.getLabel
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
        Case refreshPivotButton.getID
            Let returnedVal = refreshPivotButton.getLabel
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
        Case hidePageBreakDropDown.getID
            Let returnedVal = hidePageBreakDropDown.getLabel
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
        Case refreshAddinButton.getID
            Let returnedVal = refreshAddinButton.getLabel
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
Public Sub getShowLabel(ByRef control As IRibbonControl, ByRef returnedVal)
On Error GoTo ErrorHandle
    If loadedRibbon Is Nothing Then Call refreshCustomRibbon
    Select Case control.id
        Case addSheetsButton.getID
            Let returnedVal = addSheetsButton.getShowLabel
        Case listSheetsButton.getID
            Let returnedVal = listSheetsButton.getShowLabel
        Case renameSheetsButton.getID
            Let returnedVal = renameSheetsButton.getShowLabel
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
        Case refreshPivotButton.getID
            Let returnedVal = refreshPivotButton.getShowLabel
        Case importVbaFilesButton.getID
            Let returnedVal = importVbaFilesButton.getShowLabel
        Case importAllVbaFilesButton.getID
            Let returnedVal = importAllVbaFilesButton.getShowLabel
        Case exportAllVbaFilesButton.getID
            Let returnedVal = exportAllVbaFilesButton.getShowLabel
        Case hidePageBreakDropDown.getID
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
        Case refreshAddinButton.getID
            Let returnedVal = refreshAddinButton.getShowLabel
    End Select
GoTo ExecuteProcedure
ErrorHandle:
    Call tackleErrors
ExecuteProcedure:
End Sub

'Callback for getScreentip
Public Sub getScreentip(ByRef control As IRibbonControl, ByRef returnedVal)
On Error GoTo ErrorHandle
    If loadedRibbon Is Nothing Then Call refreshCustomRibbon
    Select Case control.id
        Case addSheetsButton.getID
            Let returnedVal = addSheetsButton.getScreentip
        Case listSheetsButton.getID
            Let returnedVal = listSheetsButton.getScreentip
        Case renameSheetsButton.getID
            Let returnedVal = renameSheetsButton.getScreentip
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
        Case refreshPivotButton.getID
            Let returnedVal = refreshPivotButton.getScreentip
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
        Case refreshAddinButton.getID
            Let returnedVal = refreshAddinButton.getScreentip
    End Select
GoTo ExecuteProcedure
ErrorHandle:
    Call tackleErrors
ExecuteProcedure:
End Sub

'Callback for getSupertip
Public Sub getSupertip(ByRef control As IRibbonControl, ByRef returnedVal)
On Error GoTo ErrorHandle
    If loadedRibbon Is Nothing Then Call refreshCustomRibbon
    Select Case control.id
        Case addSheetsButton.getID
            Let returnedVal = addSheetsButton.getSupertip
        Case listSheetsButton.getID
            Let returnedVal = listSheetsButton.getSupertip
        Case renameSheetsButton.getID
            Let returnedVal = renameSheetsButton.getSupertip
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
        Case refreshPivotButton.getID
            Let returnedVal = refreshPivotButton.getSupertip
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
        Case refreshAddinButton.getID
            Let returnedVal = refreshAddinButton.getSupertip
    End Select
GoTo ExecuteProcedure
ErrorHandle:
    Call tackleErrors
ExecuteProcedure:
End Sub

'Callback for getVisible
Public Sub getVisible(ByRef control As IRibbonControl, ByRef returnedVal)
On Error GoTo ErrorHandle
    If loadedRibbon Is Nothing Then Call refreshCustomRibbon
    Select Case control.id
        Case toolsTab.getID
            Let returnedVal = toolsTab.getVisible
        Case addSheetsButton.getID
            Let returnedVal = addSheetsButton.getVisible
        Case listSheetsButton.getID
            Let returnedVal = listSheetsButton.getVisible
        Case renameSheetsButton.getID
            Let returnedVal = renameSheetsButton.getVisible
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
        Case refreshPivotButton.getID
            Let returnedVal = refreshPivotButton.getVisible
        Case importVbaFilesButton.getID
            Let returnedVal = importVbaFilesButton.getVisible
        Case importAllVbaFilesButton.getID
            Let returnedVal = importAllVbaFilesButton.getVisible
        Case exportAllVbaFilesButton.getID
            Let returnedVal = exportAllVbaFilesButton.getVisible
        Case hidePageBreakDropDown.getID
            Let returnedVal = hidePageBreakDropDown.getVisible
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
        Case refreshAddinButton.getID
            Let returnedVal = refreshAddinButton.getVisible
    End Select
GoTo ExecuteProcedure
ErrorHandle:
    Call tackleErrors
ExecuteProcedure:
End Sub

'Callback for getPressed
Public Sub getPressed(control As IRibbonControl, ByRef returnedVal)
On Error GoTo ErrorHandle
    If loadedRibbon Is Nothing Then Call refreshCustomRibbon
    Select Case control.id
        Case highlightBoldButton.getID
            Let returnedVal = highlightIsBold
        Case highlightButton.getID
            Let returnedVal = hasHighlight
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
    If loadedRibbon Is Nothing Then Call refreshCustomRibbon
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

'Callback for Sheet Controller onAction With Event
Public Sub sheetControllerEvent(ByRef control As IRibbonControl, Optional ByRef pressed As Boolean)
On Error GoTo ErrorHandle
    If loadedRibbon Is Nothing Then Call refreshCustomRibbon
    Set sheetCEvent = New SheetsController
    Select Case control.id
        Case listSheetsButton.getID
            Call sheetCEvent.list
            Let hasListSheet = sheetCEvent.hasListSheet
        Case renameSheetsButton.getID
            Call sheetCEvent.rename
            Let hasRenameSheet = sheetCEvent.hasRenameSheet
    End Select
    If Not hasListSheet Then Set sheetCEvent = Nothing 'Clear Event
    Call refreshCustomRibbon(loadedRibbon)
GoTo ExecuteProcedure
ErrorHandle:
    Call tackleErrors
ExecuteProcedure:
End Sub

'Callback for Chart Controller onAction
Public Sub chartController(ByRef control As IRibbonControl, Optional ByRef pressed As Boolean)
On Error GoTo ErrorHandle
    If loadedRibbon Is Nothing Then Call refreshCustomRibbon
    Dim chartC As ChartsController
    Set chartC = New ChartsController
    Select Case control.id
        Case chartHideErrButton.getID
            Call chartC.hide(isHide:=True)
        Case chartShowButton.getID
            Call chartC.hide(isHide:=False)
    End Select
    Set chartC = Nothing ' Clear Cache while don't have event
GoTo ExecuteProcedure
ErrorHandle:
    Call tackleErrors
ExecuteProcedure:
End Sub

'Callback for refresh-pivot onAction
Public Sub pivotControllerEvent(ByRef control As IRibbonControl, Optional ByRef pressed As Boolean)
On Error GoTo ErrorHandle
    If loadedRibbon Is Nothing Then Call refreshCustomRibbon
    Select Case control.id
        Case refreshPivotButton.getID
            Let hasSYNCPivot = pressed
            If hasSYNCPivot Then
                Set pivotCEvent = New PivotTablesController
            Else
                Set pivotCEvent = Nothing
            End If
            Call refreshCustomRibbon(loadedRibbon)
    End Select
GoTo ExecuteProcedure
ErrorHandle:
    Call tackleErrors
ExecuteProcedure:
End Sub

'Callback for VBAFiles Controller onAction
Public Sub VBAFilesController(ByRef control As IRibbonControl, Optional ByRef pressed As Boolean)
On Error GoTo ErrorHandle
    If loadedRibbon Is Nothing Then Call refreshCustomRibbon
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
    If loadedRibbon Is Nothing Then Call refreshCustomRibbon
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
    Call refreshCustomRibbon(loadedRibbon)
GoTo ExecuteProcedure
ErrorHandle:
    Call tackleErrors
ExecuteProcedure:
End Sub

'Callback for Hide Page Break onAction
Public Sub hidePageBreakChange(control As IRibbonControl, id As String, index As Integer)
On Error GoTo ErrorHandle
    If loadedRibbon Is Nothing Then Call refreshCustomRibbon
    Dim rangeC As RangesController
    Set rangeC = New RangesController
    Select Case control.id
        Case hidePageBreakDropDown.getID
            Select Case index
                Case 0 'Hide
                    Call rangeC.displayPageBreak( _
                        isDisplay:=False, _
                        isApplyAll:=False)
                Case 1 'Show
                    Call rangeC.displayPageBreak( _
                    isDisplay:=True, _
                    isApplyAll:=False)
                Case 2 'Hide All
                    Call rangeC.displayPageBreak( _
                        isDisplay:=False, _
                        isApplyAll:=True)
                Case 3 'Show All
                    Call rangeC.displayPageBreak( _
                        isDisplay:=True, _
                        isApplyAll:=True)
            End Select
    End Select
    Call refreshCustomRibbon(loadedRibbon)
    Set rangeC = Nothing
GoTo ExecuteProcedure
ErrorHandle:
    Call tackleErrors
ExecuteProcedure:
End Sub

'Callback for Range Controller onAction
Public Sub rangeControllerEvent(ByRef control As IRibbonControl, Optional ByRef pressed As Boolean)
On Error GoTo ErrorHandle
    If loadedRibbon Is Nothing Then Call refreshCustomRibbon
    Select Case control.id
        Case highlightButton.getID
            Let hasHighlight = pressed
            If pressed Then
                Set rangeCEvent = New RangesController '*Reduce procedure
                Call rangeCEvent.storeHighlightFormat
                Call HighlightRange
            Else
                Call ClearHighLight
                Set rangeCEvent = Nothing 'Clear event
            End If
    End Select
    Call refreshCustomRibbon(loadedRibbon)
GoTo ExecuteProcedure
ErrorHandle:
    Call tackleErrors
ExecuteProcedure:
End Sub

'Callback for settings onAction
Public Sub accessSettings(ByRef control As IRibbonControl, Optional ByRef pressed As Boolean)
On Error GoTo ErrorHandle
    If loadedRibbon Is Nothing Then Call refreshCustomRibbon
    Select Case control.id
        Case settingsButton.getID
            Let userResponse = MsgBox( _
                Prompt:="This button is on process", _
                Buttons:=vbOKOnly + vbExclamation, _
                Title:="DANH TOOLS")
End Select
GoTo ExecuteProcedure
ErrorHandle:
    Call tackleErrors
ExecuteProcedure:
End Sub

'Callback for Internet Controller onAction
Public Sub internetController(ByRef control As IRibbonControl, Optional ByRef pressed As Boolean)
On Error GoTo ErrorHandle
    If loadedRibbon Is Nothing Then Call refreshCustomRibbon
    Dim internetC As InternetConnector
    Set internetC = New InternetConnector
    Select Case control.id
        Case exportWifiToTxtButton.getID
            Call internetC.saveWifiAsTxt
        Case exportWifiToCsvButton.getID
            Call internetC.saveWifiAsCsv
        Case exportWifiToJsonButton.getID
            Call internetC.saveWifiAsJson
    End Select
        Set internetC = Nothing ' Clear Cache
GoTo ExecuteProcedure
ErrorHandle:
    Call tackleErrors
ExecuteProcedure:
End Sub

'Callback for remove-addin onAction
Public Sub addinController(ByRef control As IRibbonControl, Optional ByRef pressed As Boolean)
On Error GoTo ErrorHandle
    If loadedRibbon Is Nothing Then Call refreshCustomRibbon
    Dim newAddin As AutoAddin
    Set newAddin = New AutoAddin
    Select Case control.id
        Case removeAddinButton.getID
            Call newAddin.remove(hasConfirm:=True)
        Case refreshAddinButton.getID
            Call refreshCustomRibbon(loadedRibbon)
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
    If loadedRibbon Is Nothing Then Call refreshCustomRibbon
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
                Let userResponse = MsgBox( _
                    Prompt:="This sheet don't exist any object to arrange yet!", _
                    Buttons:=vbOKOnly + vbExclamation, _
                    Title:="DANH TOOL")
            Let isArranging = False
            Else
                Let isArranging = pressed
                Set picC = New PicturesController
                If isArranging Then
'                    Let picC.letOffset = offsetValue
'                    Let picC.letLockRatio = isRateLock
                    Call picC.assign
                Else
                    Call picC.clearArrange
                End If
                Set picC = Nothing
            End If
            Call refreshCustomRibbon(loadedRibbon)
        Case autoArrangeButton.getID
            If ActiveSheet.Shapes.Count = 0 Then
                Let userResponse = MsgBox( _
                    Prompt:="This sheet don't exist any object to AUTO arrange!", _
                    Buttons:=vbOKOnly + vbExclamation, _
                    Title:="DANH TOOL")
            Let isAutoArrange = False
            Else
                Let isAutoArrange = pressed
                Set picC = New PicturesController
                If isAutoArrange Then
                    Call ThisWorkbook.Auto_Run_Continuously
                Else
                    Call ThisWorkbook.Stop_Run_Continously
                End If
                Call picC.autoArrange(isAutoArrange)
                Let picC.selectObjectMode = isAutoArrange
                Set picC = Nothing
            End If
            Call refreshCustomRibbon(loadedRibbon)
        Case rateLockCheckBox.getID
            Let isRateLock = pressed
    End Select
GoTo ExecuteProcedure
ErrorHandle:
    Call tackleErrors
ExecuteProcedure:
End Sub

'Callback for getItemCount
Public Sub getItemCount(control As IRibbonControl, ByRef returnedVal)
On Error GoTo ErrorHandle
    If loadedRibbon Is Nothing Then Call refreshCustomRibbon
    Select Case control.id
        Case offsetCBBox.getID
            Let returnedVal = NUM_OFFSET_ITEMS
    End Select
GoTo ExecuteProcedure
ErrorHandle:
    Call tackleErrors
ExecuteProcedure:
End Sub

'Callback for getItemID
Public Sub getItemID(control As IRibbonControl, index As Integer, ByRef returnedVal)
On Error GoTo ErrorHandle
    If loadedRibbon Is Nothing Then Call refreshCustomRibbon
    Select Case control.id
        Case offsetCBBox.getID
            Let returnedVal = "offset-item-" & index + 1
    End Select
GoTo ExecuteProcedure
ErrorHandle:
    Call tackleErrors
ExecuteProcedure:
End Sub

'Callback for getItemLabel
Public Sub getItemLabel(control As IRibbonControl, index As Integer, ByRef returnedVal)
On Error GoTo ErrorHandle
    If loadedRibbon Is Nothing Then Call refreshCustomRibbon
    Select Case control.id
        Case offsetCBBox.getID
            Let returnedVal = index * 10 'Steps
    End Select
GoTo ExecuteProcedure
ErrorHandle:
    Call tackleErrors
ExecuteProcedure:
End Sub

'Callback for getText
Public Sub getText(control As IRibbonControl, ByRef returnedVal)
On Error GoTo ErrorHandle
    If loadedRibbon Is Nothing Then Call refreshCustomRibbon
    Select Case control.id
        Case offsetCBBox.getID
            Let returnedVal = offsetValue
    End Select
GoTo ExecuteProcedure
ErrorHandle:
    Call tackleErrors
ExecuteProcedure:
End Sub

'Callback for getSelectedItemIndex
Public Sub getSelectedItemIndex(control As IRibbonControl, ByRef returnedVal)
On Error GoTo ErrorHandle
    If loadedRibbon Is Nothing Then Call refreshCustomRibbon
    Select Case control.id
        Case hidePageBreakDropDown.getID
            ' Check WS again for case don't remove all Ws
            Let hasWorksheet = system.hasWorkPlace(False, "xlWorksheet")
            If hasWorksheet Then Let hasPageBreak = ActiveSheet.DisplayPageBreaks
            If hasPageBreak Then
                Let returnedVal = 1 'Index = 1 --> "Show"
            Else
                Let returnedVal = 0   'Index = 0 --> "Hide"
            End If
    End Select
GoTo ExecuteProcedure
ErrorHandle:
    Call tackleErrors
ExecuteProcedure:
End Sub

'Callback for offset onChange
Public Sub offsetSelect(control As IRibbonControl, text As String)
On Error GoTo ErrorHandle
    If loadedRibbon Is Nothing Then Call refreshCustomRibbon
    Select Case control.id
        Case offsetCBBox.getID
            If Not IsNumeric(text) Then
                Let userResponse = MsgBox( _
                    Prompt:= _
                        "You have to input a NUMBER between " & _
                        MIN_OFFSET & " and " & MAX_OFFSET & " !", _
                    Buttons:=vbOKOnly + vbCritical, _
                    Title:="DANH TOOL")
                Call refreshCustomRibbon(loadedRibbon)
            ElseIf _
                CLng(text) < MIN_OFFSET Or _
                CLng(text) > MAX_OFFSET Then
                Let userResponse = MsgBox( _
                    Prompt:= _
                        "Offset value must be between " & _
                        MIN_OFFSET & " and " & MAX_OFFSET & " !", _
                    Buttons:=vbOKOnly + vbCritical, _
                    Title:="DANH TOOL")
                Call refreshCustomRibbon(loadedRibbon)
            Else
                Let offsetValue = CByte(text)
            End If
    End Select
GoTo ExecuteProcedure
ErrorHandle:
    Call tackleErrors
ExecuteProcedure:
End Sub

