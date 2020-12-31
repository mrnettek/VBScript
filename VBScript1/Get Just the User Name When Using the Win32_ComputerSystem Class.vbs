strComputer = "."

Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * From Win32_ComputerSystem")

For Each objItem in colItems
    Wscript.Echo objItem.UserName
Next
  


