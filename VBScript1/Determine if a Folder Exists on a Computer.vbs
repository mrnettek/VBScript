Set objFSO = CreateObject("Scripting.FileSystemObject")
If objFSO.FolderExists("C:\Scripts") Then
    Wscript.Echo "The folder exists."
Else
    Wscript.Echo "The folder does not exist."
End If
  


