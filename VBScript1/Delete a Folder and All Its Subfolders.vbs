strComputer = "."
Set objWMIService = GetObject("winmgmts: \\" & strComputer & "\root\cimv2")

Set colFolders = objWMIService.ExecQuery _
    ("Select * from Win32_Directory where Name = 'c:\\Scripts'")

For Each objFolder in colFolders
    errResults = objFolder.Delete
Next
  


