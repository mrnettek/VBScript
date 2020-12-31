' Description: Demonstration script that uses the FileSystemObject to copy all the .txt files in a folder to a new location.


Const OverwriteExisting = TRUE

Set objFSO = CreateObject("Scripting.FileSystemObject")
objFSO.CopyFile "C:\FSO\*.txt" , "D:\Archive\" , OverwriteExisting

