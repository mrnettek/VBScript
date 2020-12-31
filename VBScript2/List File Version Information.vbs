' Description: Demonstration script that uses the FileSystemObject to retrieve the file version for a .dll file. Script must be run on the local computer.


Set objFSO = CreateObject("Scripting.FileSystemObject")
Wscript.Echo objFSO.GetFileVersion("c:\windows\system32\scrrun.dll")

