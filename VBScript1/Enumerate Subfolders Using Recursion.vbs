' Description: Demonstration script that uses the FileSystemObject to recursively list all the subfolders (and their subfolders) within a folder. Script must be run on the local computer.


Set FSO = CreateObject("Scripting.FileSystemObject")
ShowSubfolders FSO.GetFolder("C:\Scripts")

Sub ShowSubFolders(Folder)
    For Each Subfolder in Folder.SubFolders
        Wscript.Echo Subfolder.Path
        ShowSubFolders Subfolder
    Next
End Sub

