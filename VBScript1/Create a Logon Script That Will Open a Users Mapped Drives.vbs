On Error Resume Next

strComputer = "."

Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery("Select * From Win32_MappedLogicalDisk")

Set objShell = CreateObject("Wscript.Shell")

For Each objItem in colItems
    strDrive = objItem.DeviceID
    objShell.Run(strDrive)
Next
  


