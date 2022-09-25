Attribute VB_Name = "mAddInAutoEnvent"
' Make auto run event in add-in
Option Explicit

Public AddinAutoEvent As cAddinAutoEvent 'set as new auto run add-in class

Sub Auto_open() 'Event when open add-in to connect with class Event\
    Set AddinAutoEvent = New cAddinAutoEvent
End Sub

