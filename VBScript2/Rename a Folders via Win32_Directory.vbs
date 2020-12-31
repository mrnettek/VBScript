strComputer = "."

Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")

Set colFolders = objWMIService.ExecQuery _
    ("Select * From Win32_Directory Where Name = 'C:\\January'")

For Each objFolder in colFolders
    strNewName = objFolder.Name & "_2006"
    objFolder.Rename strNewName
Next
  


