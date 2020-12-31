Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.GetFile("C:\Scripts\Test.txt")
Wscript.Echo "File name: " & objFSO.GetFileName(objFile)
  


