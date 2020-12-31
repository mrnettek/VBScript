Const ForReading = 1
Const ForWriting = 2

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.OpenTextFile("C:\Scripts\Test.txt", ForReading)

Do Until objFile.AtEndOfStream
    strLine = objFile.ReadLine
    intRight = Len(strLine) - 10
    strRight = Right(strLine, intRight)
    strLeft = Left(strLine, 10)
    strInsert = "This is inserted text"
    strText = strLeft & strInsert & strRight
    strContents = strContents & strText & vbCrLf
Loop

objFile.Close

Set objFile = objFSO.OpenTextFile("C:\Scripts\Test.txt", ForWriting)
objFile.WriteLine strContents

objFile.Close
  


