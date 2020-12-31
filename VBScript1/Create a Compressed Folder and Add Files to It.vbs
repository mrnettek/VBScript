strComputer = "."

Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")

Set colFolders = objWMIService.ExecQuery _
    ("Select * From Win32_Directory Where Name = 'C:\\Scripts'")

For Each objFolder in colFolders
    errResults = objFolder.Compress
Next
  


