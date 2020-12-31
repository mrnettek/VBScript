' Description: Sets the initial size of a page file to 300 megabytes, and the maximum size to 600 megabytes.


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colPageFiles = objWMIService.ExecQuery _
    ("Select * from Win32_PageFileSetting")

For Each objPageFile in colPageFiles
    objPageFile.InitialSize = 300
    objPageFile.MaximumSize = 600
    objPageFile.Put_
Next

