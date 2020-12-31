' Description: Demonstration script that uses the FileSystemObject to move a folder from one location to another. Script must be run on the local computer.


Set objFSO = CreateObject("Scripting.FileSystemObject")
objFSO.MoveFolder "C:\Scripts" , "M:\helpdesk\management"

