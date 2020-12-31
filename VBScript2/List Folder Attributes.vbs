' Description: Demonstration script that uses the FileSystemObject to enumerate the attributes of a folder. Script must be run on the local computer.


Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFolder = objFSO.GetFolder("C:\FSO")

If objFolder.Attributes AND 2 Then
    Wscript.Echo "Hidden folder."
End If    
If objFolder.Attributes AND 4 Then
    Wscript.Echo "System folder."
End If    
If objFolder.Attributes AND 16 Then
    Wscript.Echo "Folder."
End If  
If objFolder.Attributes AND 32 Then
    Wscript.Echo "Archive bit set."
End If
If objFolder.Attributes AND 2048 Then
    Wscript.Echo "Compressed folder."
End If

