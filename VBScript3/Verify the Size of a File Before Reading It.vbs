' Description: Demonstration script that uses the FileSystemObject to ensure that a text file is not empty before attempting to read it. Script must be run on the local computer.


Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.GetFile("C:\Windows\Netlogon.log")

If objFile.Size > 0 Then
    Set objReadFile = objFSO.OpenTextFile("C:\Windows\Netlogon.log", 1)
    strContents = objReadFile.ReadAll
    Wscript.Echo strContents
    objReadFile.Close
Else
    Wscript.Echo "The file is empty."
End If

