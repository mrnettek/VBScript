Const ForReading = 1

Set objFSO = CreateObject("Scripting.FileSystemObject")

Set objFile = objFSO.OpenTextFile("C:\Scripts\Log.txt", ForReading)
strCharacters = objFile.Read(10)
objFile.Close

strNewName = "C:\Scripts\" & strCharacters & ".txt"

objFSO.MoveFile "C:\Scripts\Log.txt", strNewName
  


