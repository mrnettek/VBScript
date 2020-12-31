' Description: Demonstration script that checks to see if a file is read-only and, if it is not, marks it as read-only. Script must be run on the local computer.


Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.GetFile("C:\FSO\TestScript.vbs")

If objFile.Attributes = objFile.Attributes AND 1 Then
    objFile.Attributes = objFile.Attributes XOR 1 
End If

