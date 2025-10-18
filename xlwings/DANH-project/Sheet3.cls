Option Explicit

Private info As InfoConstants
Private system As SystemUpdate
Private fileSystem As Object
Private Const PING_SHEET_NAME As String = "ping"
Private Const DATA_FIRST_ROW As Byte = 3
Private Const SWITCH_ON As String = "ON"
Private Const SWITCH_ON_COLOR As Long = 49152
Private Const SWITCH_OFF As String = "OFF"
Private Const SWITCH_OFF_COLOR As Long = 255
Private Enum InternetStt
    Connect = 1
    Disconnect = 0
    Loading = -1
End Enum
Private Enum Column
    No = 1
    name = 2
    Address = 3
    status = 4
End Enum

'PING SHEET
Private Sub pingCheck()
    Dim system As SystemUpdate
    Dim internetC As InternetConnector
    Dim pingSh  As Worksheet
    Dim pingShName As String
    Dim Row As Long
    Dim link As String
    Set system = New SystemUpdate
    Set internetC = New InternetConnector
    Let pingShName = PING_SHEET_NAME
    Set pingSh = Application.ThisWorkbook.Worksheets(pingShName)
    Do Until Not Switch_Button
        For Row = DATA_FIRST_ROW To system.getLastRow(pingSh, Column.Address)
            Let link = pingSh.Cells(Row, Column.Address).value
            If internetC.isConnect(link) Then
                Let pingSh.Cells(Row, Column.status).value = InternetStt.Connect
            Else
                Let pingSh.Cells(Row, Column.status).value = InternetStt.Disconnect
            End If
        Next Row
    Loop
    For Row = DATA_FIRST_ROW To system.getLastRow(pingSh, Column.Address)
        Let pingSh.Cells(Row, Column.status).value = InternetStt.Loading
    Next Row
End Sub

Private Sub Switch_Button_Click()
    If Switch_Button Then
        Let Switch_Button.caption = SWITCH_ON
        Let Switch_Button.foreColor = SWITCH_ON_COLOR
        Call pingCheck
    Else
        Let Switch_Button.caption = SWITCH_OFF
        Let Switch_Button.foreColor = SWITCH_OFF_COLOR
    End If
End Sub
