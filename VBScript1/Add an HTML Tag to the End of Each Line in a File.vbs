Const ForReading = 1
Const ForWriting = 2

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.OpenTextFile("C:\Scripts\Test.txt", ForReading)

strContents = objFile.ReadAll
objFile.Close
strContents = Replace(strContents, vbCrlf, "<br>" & vbCrLf)

Set objFile = objFSO.OpenTextFile("C:\Scripts\Test.txt", ForWriting)

objFile.Write strContents
objFile.Close
  


