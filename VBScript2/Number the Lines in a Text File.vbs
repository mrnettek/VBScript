Const ForReading = 1
Const ForWriting = 2

i = 1

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.OpenTextFile("C:\Scripts\Test.txt", ForReading)

Do Until objFile.AtEndOfStream
    strLine = objFile.ReadLine
    strContents = strContents & i & "   " & strLine & vbCrLf
    i = i + 1
Loop

objFile.Close

Set objFile = objFSO.OpenTextFile("C:\Scripts\Test.txt", ForWriting)
objFile.Write strContents
objFile.Close
  


