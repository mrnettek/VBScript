Const ForReading = 1
Const ForWriting = 2

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.OpenTextFile("C:\Scripts\Test.txt", ForReading)

Do Until objFile.AtEndOfStream
    strLine1 = objFile.ReadLine
    strLine2 = objFile.ReadLine
    strNewLine = strLine1 & strLine2
    strNewContents = strNewContents & strNewLine & vbCrLf
Loop

objFile.Close

Set objFile = objFSO.OpenTextFile("C:\Scripts\Test.txt", ForWriting)
objFile.Write strNewContents
objFile.Close
  


