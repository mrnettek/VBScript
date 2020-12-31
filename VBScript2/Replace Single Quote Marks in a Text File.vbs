Const ForReading = 1
Const ForWriting = 2

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.OpenTextFile("C:\Scripts\Test.txt", ForReading)

strText = objFile.ReadAll
objFile.Close

strOldText = Chr(39)
strNewText = Chr(34)

strNewText = Replace(strText, strOldText, strNewText)

Set objFile = objFSO.OpenTextFile("C:\Scripts\Test.txt", ForWriting)
objFile.WriteLine strNewText
objFile.Close
  


