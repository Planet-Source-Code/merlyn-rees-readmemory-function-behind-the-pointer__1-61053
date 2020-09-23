Attribute VB_Name = "modReadMemory"
Option Explicit

'modReadMemory – Add to your projects to instantly have a fully usable,
'complete, ReadMemory function!

Enum ReadMemory_VariableType
    xByte = 1
    xInteger = 2
    xLong = 3
    xSingle = 4
    xDouble = 5
    xDate = 6
End Enum

'We need all these temporary variables for copying to.  They're stored
'outside of the ReadMemory function to stop the overhead of 30 or so bytes
'of variables being allocated every call.
Dim bTemp As Byte, iTemp As Integer, lTemp As Long, sTemp As Single, dTemp As Double, dtTemp As Date
'Because I want a perfectly optimized function here, I'm going to have some
''lookup' variables to hold the pointers of the above variables.
Dim l_bTemp As Long, l_iTemp As Long, l_lTemp As Long, l_sTemp As Long, l_dTemp As Long, l_dtTemp As Long

'The CopyMemory (RtlMoveMemory) call.

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)

Sub InitRM()
    'Initialise these variables...
    bTemp = 0
    iTemp = 0
    lTemp = 0
    sTemp = 0#
    dTemp = 0#
    dtTemp = 0
    'Lookup their Pointers and put them in variables, ready for later.
    l_bTemp = VarPtr(bTemp)
    l_iTemp = VarPtr(iTemp)
    l_lTemp = VarPtr(lTemp)
    l_sTemp = VarPtr(sTemp)
    l_dTemp = VarPtr(dTemp)
    l_dtTemp = VarPtr(dtTemp)
End Sub

Function ReadMemory(ByVal Source As Long, LenV As ReadMemory_VariableType)
    Select Case LenV
        Case xByte
            'Pass the Pointer as a number which'll be interpreted by
            'Kernel32 as a Pointer :)
            CopyMemory l_bTemp, Source, 1
            'Return the new contents of bTemp
            ReadMemory = bTemp
        Case xInteger
            CopyMemory l_iTemp, Source, 2
            ReadMemory = iTemp
        Case xLong
            CopyMemory l_lTemp, Source, 4
            ReadMemory = lTemp
        Case xSingle
            CopyMemory l_sTemp, Source, 4
            ReadMemory = sTemp
        Case xDouble
            CopyMemory l_dTemp, Source, 8
            ReadMemory = dTemp
        Case xDate
            CopyMemory l_dtTemp, Source, 8
            ReadMemory = dtTemp
    End Select
End Function


