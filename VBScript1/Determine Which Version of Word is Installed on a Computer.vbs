strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")

Set colApps = objWMIService.ExecQuery _
    ("Select * from Win32_Product Where Caption Like '%Microsoft Office%'")
For Each objApp in colApps
    Wscript.Echo objApp.Caption, objApp.Version
Next
  


