' Description: Demonstration script that uses the FileSystemObject to return pathname information for a file, including  name,  extension, complete  path, etc. Script must be run on the local computer.


Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.GetFile("ScriptLog.txt")

Wscript.Echo "Absolute path: " & objFSO.GetAbsolutePathName(objFile)
Wscript.Echo "Parent folder: " & objFSO.GetParentFolderName(objFile) 
Wscript.Echo "File name: " & objFSO.GetFileName(objFile)
Wscript.Echo "Base name: " & objFSO.GetBaseName(objFile)
Wscript.Echo "Extension name: " & objFSO.GetExtensionName(objFile)

