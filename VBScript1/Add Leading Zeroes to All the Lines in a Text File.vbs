Const ForReading = 1
Const ForWriting = 2

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.OpenTextFile("C:\Scripts\Test.txt", ForReading)

Do Until objFile.AtEndOfStream
    strText = objFile.ReadLine
    intLength = Len(strText)
    intZeroes = 9 - intLength
    For i = 1 to intZeroes
        strText = "0" & strText
    Next
    strContents = strContents & strText & vbCrlf
Loop

objFile.Close

Set objFile = objFSO.OpenTextFile("C:\Scripts\Test.txt", ForWriting)
objFile.WriteLine strContents

objFile.Close
  


