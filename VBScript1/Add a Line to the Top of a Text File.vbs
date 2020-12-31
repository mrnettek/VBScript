Const ForReading = 1
Const ForWriting = 2

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.OpenTextFile("C:\Scripts\Test.txt", ForReading)

strContents = objFile.ReadAll
objFile.Close

strFirstLine = "This is the new first line in the text file."
strNewContents = strFirstLine & vbCrLf & strContents

Set objFile = objFSO.OpenTextFile("C:\Scripts\Test.txt", ForWriting)
objFile.WriteLine strNewContents

objFile.Close
  


