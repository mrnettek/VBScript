strComputer = "."

Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")

Set colFileList = objWMIService.ExecQuery _
    ("ASSOCIATORS OF {Win32_Directory.Name='C:\Windows\System32'} Where " _
        & "ResultClass = CIM_DataFile")

For Each objFile In colFileList
    If objFile.Extension = "dll" Then
        Wscript.Echo objFile.Name & " -- " & objFile.Version
    End If
Next
  


