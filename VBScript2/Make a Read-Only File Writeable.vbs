Const READ_ONLY = 1

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.GetFile("C:\Scripts\Computers.txt")

If objFile.Attributes AND READ_ONLY Then
    objFile.Attributes = objFile.Attributes XOR READ_ONLY
End If
  


