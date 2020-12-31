Const ForReading = 1

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objTextFile = objFSO.OpenTextFile _
    ("C:\Scripts\Test.txt", ForReading)

objTextFile.ReadAll
Wscript.Echo "Number of lines: " & objTextFile.Line
  


