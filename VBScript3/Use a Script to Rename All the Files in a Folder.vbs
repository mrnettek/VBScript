strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")

Set colFileList = objWMIService.ExecQuery _
    ("ASSOCIATORS OF {Win32_Directory.Name='C:\Logs'} Where " _
        & "ResultClass = CIM_DataFile")

For Each objFile In colFileList
    strNewName = objFile.Drive & objFile.Path & "pl-" & _
        objFile.FileName & "." & objFile.Extension
    errResult = objFile.Rename(strNewName)
Next
  


