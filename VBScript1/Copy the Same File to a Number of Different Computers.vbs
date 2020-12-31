Const ForReading = 1
Const OverwriteExisting = TRUE

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.OpenTextFile("C:\Scripts\Computers.txt")

Do Until objFile.AtEndOfStream
    strComputer = objFile.ReadLine
    strRemoteFile = "\\" & strComputer & "\C$\Scripts\Test.doc"
    objFSO.CopyFile "C:\Scripts\Test.doc", strRemoteFile, OverwriteExisting
Loop
  


