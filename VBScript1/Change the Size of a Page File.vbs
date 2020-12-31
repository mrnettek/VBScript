strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")

Set colPageFiles = objWMIService.ExecQuery _
    ("Select * from Win32_PageFileSetting")

For Each objPageFile in colPageFiles
    objPageFile.InitialSize = 300
    objPageFile.MaximumSize = 600
    objPageFile.Put_
Next
  


