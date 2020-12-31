OpTion Explicit

Dim arrText, iCount
arrText = ReadFileToArray("C:\test.log")

If IsArray(arrText) then

	For iCount = 0 to UBound(arrText)

		MsgBox arrText(iCount)
	Next

End if


' ---------------------------------------
Private Function ReadFileToArray(strFile)

Dim ofso : Set ofso = Createobject("Scripting.FileSystemObject")
Const ForReading	= 1
Const ForWriting 	= 2
Const ForAppending 	= 8

Dim strNextLine, arrstrList
Dim arrLines()
Dim iCount : iCount = 0

	If ofso.FileExists(strFile) then

		Dim oFile : Set oFile = ofso.OpenTextFile(strFile, ForReading)

		Do Until oFile.AtEndOfStream

			Redim Preserve arrLines(iCount)
			arrLines(iCount) = oFile.ReadLine
			iCount = iCount + 1

		Loop

		oFile.Close

		ReadFileToArray = arrLines

	Else
		ReadFileToArray = 0
	End if

End Function
