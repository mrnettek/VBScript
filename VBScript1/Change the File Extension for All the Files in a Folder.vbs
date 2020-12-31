strComputer = "."

Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")

Set colFiles = objWMIService.ExecQuery _
    ("ASSOCIATORS OF {Win32_Directory.Name='C:\Test'} Where " _
        & "ResultClass = CIM_DataFile")

For Each objFile In colFiles
    strExtension = objFile.Extension 
    strExtension = Replace(strExtension, "old", "")
    strNewName = objFile.Drive & objFile.Path & objFile.FileName & "." & strExtension
    errResult = objFile.Rename(strNewName)
Next
  


