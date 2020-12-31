On Error Resume Next

strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * from Win32_NetworkAdapterConfiguration Where MACAddress = '99:99:99:AA:99:A9'")

For Each objItem in colItems
    For Each strIPAddress in objItem.IPAddress
        Wscript.Echo "IP Address: " & strIPAddress
    Next
Next
  


