strComputer = "."

Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")

Set colFileList = objWMIService.ExecQuery _
    ("ASSOCIATORS OF {Win32_Directory.Name='T:\Act'} Where " _
        & "ResultClass = CIM_DataFile")

For Each objFile In colFileList
    If InStr(objFile.FileName, "current") Then
        objFile.Delete
    End If
Next
  


