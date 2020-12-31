Const ForReading = 1
Const ForWriting = 2

Dim arrFileLines()
i = 0

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.OpenTextFile("C:\Scripts\Test.txt", ForReading)

Do Until objFile.AtEndOfStream
     Redim Preserve arrFileLines(i)
     arrFileLines(i) = objFile.ReadLine
     i = i + 1
Loop

objFile.Close

intLastLine = Ubound(arrFileLines)
intNumber = arrFileLines(intlastLine)
intNumber = intNumber + 1

Set objFile = objFSO.OpenTextFile("c:\scripts\test.txt", ForWriting)

For i = 0 to Ubound(arrFileLines) - 1
    objFile.WriteLine arrFileLines(i)
Next

objFile.WriteLine intNumber

objFile.Close
  


