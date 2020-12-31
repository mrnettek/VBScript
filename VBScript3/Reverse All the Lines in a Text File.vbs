Const ForReading = 1
Const ForWriting = 2

Dim arrLines()
i = 0

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.OpenTextFile("C:\Scripts\Test.txt", ForReading)

Do Until objFile.AtEndOfStream
    Redim Preserve arrLines(i)
    arrLines(i) = objFile.ReadLine
    i = i + 1
Loop

objFile.Close

Set objFile = objFSO.OpenTextFile("C:\Scripts\Test.txt", ForWriting)

For i = Ubound(arrLines) to LBound(arrLines) Step -1
    objFile.WriteLine arrLines(i)
Next

objFile.Close
  


