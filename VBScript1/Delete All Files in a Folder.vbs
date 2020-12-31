' Description: Demonstration script that deletes all the .txt files in a folder. Script must be run on the local computer.


Const DeleteReadOnly = TRUE

Set objFSO = CreateObject("Scripting.FileSystemObject")
objFSO.DeleteFile("C:\FSO\*.txt"), DeleteReadOnly

