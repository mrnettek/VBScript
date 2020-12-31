Const ForReading = 1
Const ForWriting = 2

Set objFSo = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.OpenTextFile("c:\scripts\test.txt", ForReading)

strContents = objFile.ReadAll
objFile.Close

arrLines = Split(strContents, vbCrLf)

Set objFile = objFSO.OpenTextFile("c:\scripts\test.txt", ForWriting)

For i = 0 to UBound(arrLines) - 1
    objFile.WriteLine arrLines(i)
Next

objFile.Close
  


