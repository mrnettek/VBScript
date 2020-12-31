strComputer = "."

Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * from Win32_PNPEntity Where ConfigManagerErrorCode <> 0")

For Each objItem in colItems
    Wscript.Echo "Name: " & objItem.Name
Next
  


