' Description: Demonstration script that uses the FileSystemObject to move all the .txt files in a folder to a new location. Script must be run on the local computer.


Set objFSO = CreateObject("Scripting.FileSystemObject")
objFSO.MoveFile "C:\FSO\*.txt" , "D:\Archive\"

