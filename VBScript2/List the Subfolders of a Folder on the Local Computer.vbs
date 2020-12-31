' Description: Demonstration script that uses the FileSystemObject to return the folder name and size for all the subfolders in a folder. Script must be run on the local computer.


Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFolder = objFSO.GetFolder("C:\FSO")
Set colSubfolders = objFolder.Subfolders

For Each objSubfolder in colSubfolders
    Wscript.Echo objSubfolder.Name, objSubfolder.Size
Next

