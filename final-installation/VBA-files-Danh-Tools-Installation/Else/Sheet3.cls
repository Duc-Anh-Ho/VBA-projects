VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
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
    Status = 4
End Enum

'PING SHEET
Private Sub pingCheck()
    Dim system As SystemUpdate
    Dim internetC As InternetConnector
    Dim pingSh  As Worksheet
    Dim pingShName As String
    Dim row As Long
    Dim link As String
    Set system = New SystemUpdate
    Set internetC = New InternetConnector
    Let pingShName = PING_SHEET_NAME
    Set pingSh = Application.ThisWorkbook.Worksheets(pingShName)
    Do Until Not Switch_Button
        For row = DATA_FIRST_ROW To system.getLastRow(pingSh, Column.Address)
            Let link = pingSh.Cells(row, Column.Address).value
            If internetC.isConnect(link) Then
                Let pingSh.Cells(row, Column.Status).value = InternetStt.Connect
            Else
                Let pingSh.Cells(row, Column.Status).value = InternetStt.Disconnect
            End If
        Next row
    Loop
    For row = DATA_FIRST_ROW To system.getLastRow(pingSh, Column.Address)
        Let pingSh.Cells(row, Column.Status).value = InternetStt.Loading
    Next row
End Sub

Private Sub Switch_Button_Click()
    If Switch_Button Then
        Let Switch_Button.Caption = SWITCH_ON
        Let Switch_Button.ForeColor = SWITCH_ON_COLOR
        Call pingCheck
    Else
        Let Switch_Button.Caption = SWITCH_OFF
        Let Switch_Button.ForeColor = SWITCH_OFF_COLOR
    End If
End Sub
