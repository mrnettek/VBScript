' Description: Demonstration script that uses the FileSystemObject to move a file from one location to another. Script must be run on the local computer.


Set objFSO = CreateObject("Scripting.FileSystemObject")
objFSO.MoveFile "C:\FSO\ScriptLog.log" , "D:\Archive"

