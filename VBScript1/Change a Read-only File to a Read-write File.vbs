Const ReadOnly = 1

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.GetFile("C:\Scripts\Test.vbs")

If objFile.Attributes AND ReadOnly Then
    objFile.Attributes = objFile.Attributes XOR ReadOnly
End If
  


