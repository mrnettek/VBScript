' Description: Demonstration script that uses the FileSystemObject to enumerate the properties of a folder. Script must be run on the local computer.


Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFolder = objFSO.GetFolder("C:\Scripts")

Wscript.Echo "Date created: " & objFolder.DateCreated
Wscript.Echo "Date last accessed: " & objFolder.DateLastAccessed
Wscript.Echo "Date last modified: " & objFolder.DateLastModified
Wscript.Echo "Drive: " & objFolder.Drive
Wscript.Echo "Is root folder: " & objFolder.IsRootFolder
Wscript.Echo "Name: " & objFolder.Name
Wscript.Echo "Parent folder: " & objFolder.ParentFolder
Wscript.Echo "Path: " & objFolder.Path
Wscript.Echo "Short name: " & objFolder.ShortName
Wscript.Echo "Short path: " & objFolder.ShortPath
Wscript.Echo "Size: " & objFolder.Size
Wscript.Echo "Type: " & objFolder.Type

