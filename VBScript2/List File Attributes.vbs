' Description: Demonstration script that uses the FileSystemObject to enumerate the attributes of a file. Script must be run on the local computer.


Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.GetFile("C:\FSO\ScriptLog.txt")

If objFile.Attributes AND 0 Then
    Wscript.Echo "No attributes set."
End If    
If objFile.Attributes AND 1 Then
    Wscript.Echo "Read-only."
End If    
If objFile.Attributes AND 2 Then
    Wscript.Echo "Hidden file."
End If    
If objFile.Attributes AND 4 Then
    Wscript.Echo "System file."
End If    
If objFile.Attributes AND 32 Then
    Wscript.Echo "Archive bit set."
End If    
If objFile.Attributes AND 64 Then
    Wscript.Echo "Link or shortcut."
End If    
If objFile.Attributes AND 2048 Then
    Wscript.Echo "Compressed file."
End If

