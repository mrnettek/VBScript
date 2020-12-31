Const ForWriting = 2

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.OpenTextFile("C:\Scripts\Test.txt", ForWriting)

objFile.Write ""
objFile.Close
  


