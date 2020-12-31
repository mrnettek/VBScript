' Drag and drop the file over the Script

On Error Resume Next
Dim ofso 		: Set ofso		= CreateObject("Scripting.FileSystemObject")


' --------------------- Constants / working with Textfiles
Const ForReading 	= 1
Const ForWriting 	= 2
Const ForAppending 	= 8


If ofso.FileExists(wscript.arguments(0)) then

	Dim strFilePath : strFilePath = wscript.arguments(0)
	Dim oFile : Set oFile = ofso.OpenTextFile(strFilePath, ForReading)
	Dim sLine : sLine = ""

	Do Until oFile.AtEndOfStream
		sLine = sLine & Trim(oFile.Readline) & vbcrlf
	Loop
	oFile.Close

	sLine = Left(sLine,Len(sLine)-2)

	Set oFile = ofso.OpenTextFile(wscript.arguments(0),ForWriting)
	oFile.Write sLine
	oFile.Close

End if
