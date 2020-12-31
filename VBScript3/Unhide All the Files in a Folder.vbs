strComputer = "."

Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")

Set FileList = objWMIService.ExecQuery _
    ("ASSOCIATORS OF {Win32_Directory.Name='c:\Scripts'} Where " _
        & "ResultClass = CIM_DataFile")

Set objFSO = CreateObject("Scripting.FileSystemObject")

For Each objFile In FileList
    Set objFile = objFSO.GetFile(objFile.Name)

    If objFile.Attributes AND 2 Then
        objFile.Attributes = objFile.Attributes XOR 2 
    End If
Next
  


