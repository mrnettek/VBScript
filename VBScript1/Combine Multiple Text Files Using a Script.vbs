Const ForReading = 1

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objOutputFile = objFSO.CreateTextFile("output.txt")

Set objTextFile = objFSO.OpenTextFile("c:\logs\file1.log", ForReading)

strText = objTextFile.ReadAll
objTextFile.Close
objOutputFile.WriteLine strText

Set objTextFile = objFSO.OpenTextFile("c:\logs\file2.log ", ForReading)

strText = objTextFile.ReadAll
objTextFile.Close
objOutputFile.WriteLine strText

objOutputFile.Close
  


