Const ForReading = 1
Const ForWriting = 2

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objTextFile = objFSO.OpenTextFile("c:\Scripts\Test.txt", ForReading)

strText = objTextFile.ReadAll
objTextFile.Close

arrLines = Split(strText, vbCrLf)

Set objFile = objFSO.OpenTextFile("C:\Scripts\Test.txt", ForWriting)

For i = 1 to (Ubound(arrLines) - 1)
    objFile.WriteLine arrLines(i)
Next

objFile.Close
  


