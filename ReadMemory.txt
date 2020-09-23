NOTE: If you can't open any of the OpenOffice or Word format documents, try the HTML one as a second priority.  This is if your computer's dead, or running off DOS or something...


ReadMemory
The theory is simple enough, copying memory to a variable or two.  In fact, you may be tempted to use just straight CopyMemory (a.k.a. RtlMoveMemory, (which doesn't actually move memory), thanks Hardcore VB 2nd Edition), but then we start to have Decisions.
Should we pass pointers (as ByVal Long variables) or pass what we're actually using as ByRef?  The ByVal Long variables allow you to even do difficult stuff with C and API functions pulled from rusty DLLs, but using ByRef means that VB can do some work for you, but you lose the total control you have over the system.
Maybe some possible implementations might clear the picture -

Pointers
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)
ByRef  (would require a different version of CopyMemory for every data type)
Private Declare Sub CopyMemoryL Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Long, ByRef Source As Long, ByVal Length As Long)
The vague implementation used officially:
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

With the official, vague version, you can either put in any old variable (ByRef) or type ByVal as an operator when using the function (CopyMemory ByVal SomePlace, ...) so as to provide maximum flexibility.
But it's not Basic.  It's ugly, and you could break it just by putting in the wrong value for Length.  Chances are that if you're using CopyMemory, you're probably just after something some Windows API function has given you, but which VB's Declare syntax isn't vague/forgiving/lenient enough to actually let you define.
For example, the function that got me started with trying to make a ReadMemory:
Private Declare Function FreeImage_GetPalette Lib "FreeImage.dll" Alias "_FreeImage_GetPalette@4" (ByVal dib As Long) As Long
The 'Long' it returns is actually a pointer, to a location in memory that'll contain as many RGBQUADs as there are in the image's palette, back-to-back, in the familiar C-style way.  Now, assuming there was a SetPalette function (there isn't) I would call it like this:
Option Explicit
Private Declare Sub FreeImage_SetPalette Lib "FreeImage.dll" Alias "_FreeImage_SetPalette@37" (ByVal dib As Long, ByRef pal As RGBQUAD, ByVal palLength As Long)
Dim Palette(0 To 15) As RGBQuad
'..
FreeImage_SetPalette dibHandle1, Palette(0), 16
'..
..And so, logically, you would be able to call FreeImage_GetPalette like this...
Option Explicit
Private Declare Function FreeImage_GetPalette Lib "FreeImage.dll" Alias "_FreeImage_GetPalette@4" (ByVal dib As Long) As RGBQUAD()
Dim Palette(0 To 15) As RGBQuad
'..
Palette()=FreeImage_GetPalette(dibHandle1)
'..

...But when you run past the line or perform a full compile, you get a 'Compiler error: Can't assign to array' message.
So how about Palette(0)=FreeImage_GetPalette(dibHandle1)?
''Compiler error: Type mismatch'', which makes about as much sense as anything does.

By now, I'm pretty sure that I need my ReadMemory.  So we'll quickly clean up the declaration to give us a pointer...
Option Explicit
Private Declare Function FreeImage_GetPalette Lib "FreeImage.dll" Alias "_FreeImage_GetPalette@4" (ByVal dib As Long) As Long
Dim Palette(0 To 15) As RGBQuad
Dim hPalette As Long
'..
hPalette = FreeImage_GetPalette(dibHandle1)
'..
...And we'll think about our ReadMemory function. :)
I'm not even going to speculate how VB stores UDTs in memory, so our ReadMemory function will just return single variables.  This'll add greatly to the simplicity of the function and make it easier to use, but also give us the control we need.
By now, you've probably got some sort of idea on how we can implement several different ReadMemory's, but it wouldn't be .... Basic to have a ReadMemoryB for Bytes, ReadMemoryI for Integers, ReadMemoryL for Longs, etc.
So how about an Enum?
Enum ReadMemory_VariableType
	xByte = 1
	xInteger = 2
	xLong = 4
	xSingle = 4
	xDouble = 8
	xDate = 8
End Enum
Boolean isn't implemented, as Win32 functions tend to return BOOLs, which is just another word for DWORD, or the Long data type in VB6.  This is 4 bytes long, whereas VB's Boolean data type is 2 bytes long, even though the only values it can take (reasonably) are True and False.
Okay, I've put it off long enough – my first version of my implementation of a ReadMemory function in Visual Basic...

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
	bTemp=0
	iTemp=0
	lTemp=0
	sTemp=0#
	dTemp=0#
	dtTemp=0
	'Lookup their Pointers and put them in variables, ready for later.
	l_bTemp=VarPtr(bTemp)
	l_iTemp=VarPtr(iTemp)
	l_lTemp=VarPtr(lTemp)
	l_sTemp=VarPtr(sTemp)
	l_dTemp=VarPtr(dTemp)
	l_dtTemp=VarPtr(dtTemp)
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

That's the best implementation I've come up with so far, and I feel it's pretty effective too.  Ease of use?
...Pretty easy...