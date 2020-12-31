On Error Resume Next

Set FSO = CreateObject("Scripting.FileSystemObject")
ShowSubFolders FSO.GetFolder("C:\")

Sub ShowSubFolders(Folder)
    For Each Subfolder in Folder.SubFolders
        If Subfolder.Size = 0 Then
            Wscript.Echo Subfolder.Path
        End If
        ShowSubFolders Subfolder
    Next
End Sub
  


