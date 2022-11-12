VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ChartController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Check README.md for more information
Option Explicit
'Declare Variables
Private info As Object
Private system As Object
Private fileSystem As Object
Private userResponse As VbMsgBoxResult
Private app As Application
Private Wb As Workbook
Private ws As Worksheet
Private chartTarget As Chart
Private excelErrors() As Variant
'Initialize Variables
Private Function hasVariables() As Boolean
On Error GoTo ErrorHandle
    Set info = New MyInfo
    Set system = New SystemUpdate
    Set fileSystem = system.createFileSystem
    If Not system.hasWorkplace Then
        hasVariables = False
        GoTo ExecuteProcedure
    End If
    Let excelErrors = Array("#DIV?0!", "#N/A", "#NAME?", "#NULL!", "#NUM!", "#REF!", "#VALUE!") 'Common error excel
    If ActiveChart Is Nothing Then
        MsgBox _
            Prompt:=info.getPrompt & _
                "Please select a chart first!", _
            Buttons:=vbOKOnly + vbExclamation, _
            Title:=info.getAuthor
            hasVariables = False
        GoTo ExecuteProcedure
    Else
        Set chartTarget = ActiveChart
    End If
    hasVariables = True
    GoTo ExecuteProcedure
ErrorHandle:
    Call system.tackleErrors
ExecuteProcedure:
End Function

'METHODS

Private Sub hideErrorChartLabels(Optional ByVal isHide As Boolean = True)
    Dim seriesItem As Object
    Dim labelItem As Object
    Dim item As Variant
     'Loop through trend series
    For Each seriesItem In chartTarget.FullSeriesCollection
        seriesItem.HasDataLabels = True
        'Loop through each point of DataLabels (not through DataLabels)
        For Each labelItem In seriesItem.Points
            If isHide Then
                'Loop through errors code
                For Each item In excelErrors
                    If labelItem.datalabel.Text = item Then
                        labelItem.datalabel.Text = "" ' Cheating null String
                        'NOTE: Can't use ShowValue or Delete cause error -2147024809
                    End If
                Next item
            ElseIf Not isHide Then
                If labelItem.datalabel.Text = "" Then
                        labelItem.datalabel.delete
                        labelItem.datalabel.ShowValue = True
                End If
            End If
        Next labelItem
    Next seriesItem
End Sub

'MAIN
Public Sub hide(Optional ByVal isHide As Boolean = True)
On Error GoTo ErrorHandle
    If Not hasVariables Then GoTo ExecuteProcedure
    Call hideErrorChartLabels(isHide)
'    Call hideErrorChartLabels(isHide)
GoTo ExecuteProcedure
ErrorHandle:
    Call system.tackleErrors
ExecuteProcedure:
End Sub


