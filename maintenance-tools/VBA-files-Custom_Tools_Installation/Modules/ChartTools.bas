Attribute VB_Name = "ChartTools"
Option Explicit
' Copyright - BuiDanhVN indie software
' Ver.2020.0.1.1
Public Sub hide_show_err_chart_labels()
    'Initialize variables
    Dim ws As Worksheet
    Dim ch As Chart
    Dim errorcodes
    Dim item, series, datalabel As Byte
    
    Set ws = ActiveWorkbook.ActiveSheet
    Set ch = ActiveChart
    'Set ch = ws.ChartObjects("Chart 1").Chart 'test on a particular chart
    errorcodes = Array("#NAME?", "#VALUE!", "#DIV?0!", "#NULL!", "#REF!", "#N/A") 'comon error excel
    On Error GoTo ErrHandler
    With ch
            'looping through trend series
            For series = 1 To .FullSeriesCollection.count
                    'looping through each label
                    For datalabel = 1 To .FullSeriesCollection(series).DataLabels.count
                        'looping through error excel
                        For item = LBound(errorcodes) To UBound(errorcodes)
                            'checking if match excel err
                            If errorcodes(item) = .FullSeriesCollection(series).DataLabels(datalabel).Text Then
                                .FullSeriesCollection(series).DataLabels(datalabel).ShowValue = False
                            End If
                        Next item
                    Next datalabel
        Next series
    End With
    
ErrHandler:
    'reversing if already hidden error label
    If Err.Number = -2147024809 Then
        With ch
            For series = 1 To .FullSeriesCollection.count
                .FullSeriesCollection(series).ApplyDataLabels
'                .FullSeriesCollection(series).DataLabels.ShowValue = True 'xay ra loi neu chay dong nay chu k phai dong tren
            Next series
            End
        End With
     ' Not selecting a chart
    ElseIf Err.Number = 91 Then
        MsgBox ("Please select ONE chart!")
    ' Other cases
    Else
        'pass
    End If
End Sub
