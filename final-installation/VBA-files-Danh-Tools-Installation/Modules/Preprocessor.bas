Attribute VB_Name = "Preprocessor"
Option Explicit

'Preprocessor Library Declaration
#If Mac Then ' <--Mac
    MsgBox "MacOS chay khong duoc, cai win di"
#Else '<-- Window
    #If VBA7 Then '<-- VBA7 - Office Ver > 2007
        'Library kernel32
        Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal Milliseconds As LongPtr)
        Public Declare PtrSafe Function GetTickCount Lib "kernel32" () As LongPtr
        Public Declare PtrSafe Sub CopyMemory Lib "kernel32.dll" Alias _
           "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, _
            ByVal Length As Long)
        'Subs/Functions
        ' Get Ribbon From Pointer Memory
        Public Function GetRibbon(ByVal lRibbonName As Name) As IRibbonUI
            Dim objRibbon As Object
            Dim lRibbonPointer As LongPtr
            Let lRibbonPointer = CLngPtr(Replace(lRibbonName.RefersTo, "=", ""))
            CopyMemory _
                Destination:=objRibbon, _
                Source:=lRibbonPointer, _
                Length:=LenB(lRibbonPointer)
            Set GetRibbon = objRibbon
            Set objRibbon = Nothing
        End Function
        
        
        #If Win64 Then '<-- Win 64 Bit
        #Else '<-- Win 32 Bit || Win 16 Bit
        #End If
        
        
    #Else '<-- VBA6 - Office Ver <= 2007
        'Library kernel32
        Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal Milliseconds As Long)
        Public Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long
        Public Declare PtrSafe Sub CopyMemory Lib "kernel32.dll" Alias _
           "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, _
            ByVal Length As Long)
        Public Function GetRibbon(ByVal lRibbonPointer As Long) As IRibbonUI
            Dim objRibbon As Object
            CopyMemory _
                Destination:=objRibbon, _
                Source:=lRibbonPointer, _
                Length:=LenB(lRibbonPointer)
            Set GetRibbon = objRibbon
            Set objRibbon = Nothing
        End Function
    #End If
#End If

' Get Ribbon From Pointer Memory
'#If VBA7 Then
'Public Function GetRibbon(ByVal lRibbonPointer As LongPtr) As Object
'#Else
'Public Function GetRibbon(ByVal lRibbonPointer As Long) As Object
'#End If
'        Dim objRibbon As Object
'        CopyMemory _
'            Destination:=objRibbon, _
'            Source:=lRibbonPointer, _
'            Length:=LenB(lRibbonPointer)
'        Set GetRibbon = objRibbon
'        Set objRibbon = Nothing
'End Function
