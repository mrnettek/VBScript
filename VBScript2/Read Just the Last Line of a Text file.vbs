Const ForReading = 1

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.OpenTextFile("C:\Scripts\Test.txt", ForReading)

Do Until objFile.AtEndOfStream
    strLine = objFile.ReadLine
Loop

objFile.Close

Wscript.Echo strLine
  


