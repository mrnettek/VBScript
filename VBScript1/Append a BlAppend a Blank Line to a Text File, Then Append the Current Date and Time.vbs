Const ForAppending = 8

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.OpenTextFile("C:\Scripts\Test.txt", ForAppending)

objFile.WriteLine
objFile.WriteLine
objFile.Write Now

objFile.Close
  


