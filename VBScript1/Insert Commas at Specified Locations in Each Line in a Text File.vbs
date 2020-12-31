Const ForReading = 1
Const ForWriting = 2

arrCommas = Array(2,7,11,17,18,20)

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.OpenTextFile("C:\Scripts\Test.txt", ForReading)

Do Until objFile.AtEndOfStream
    strLine = objFile.ReadLine
    intLength = Len(strLine)
    For Each strComma in arrCommas
        strLine = Left(strLine, strComma - 1) + "," + Mid(strLine, strComma, intLength)
    Next
    strText = strText & strLine & vbCrLf
Loop

objFile.Close

Set objFile = objFSO.OpenTextFile("C:\Scripts\Test.txt", ForWriting)
objFile.Write strText
objFile.Close
  


