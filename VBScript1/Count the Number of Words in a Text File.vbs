Const ForReading = 1

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.OpenTextFile("c:\scripts\test.txt", ForReading)

strText = objFile.ReadAll
objFile.Close

arrWords = Split(strText, " ")
Wscript.Echo Ubound(arrWords) + 1
  


