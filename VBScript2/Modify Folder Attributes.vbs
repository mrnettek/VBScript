' Description: Demonstration script that uses the FileSystemObject to check if a folder is hidden and, if it is not, hides it. Script must be run on the local computer.


Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFolder = objFSO.GetFolder("C:\FSO")

If objFolder.Attributes = objFolder.Attributes AND 2 Then
    objFolder.Attributes = objFolder.Attributes XOR 2 
End If

