' Description: Demonstration script that creates a new, empty text file. Script must be run on the local computer.


Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.CreateTextFile("C:\FSO\ScriptLog.txt")

