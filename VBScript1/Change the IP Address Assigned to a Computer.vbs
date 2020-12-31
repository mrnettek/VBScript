strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")

Set colNetAdapters = objWMIService.ExecQuery _
    ("Select * from Win32_NetworkAdapterConfiguration Where IPEnabled=True")

strIPAddress = Array("192.168.1.0")
strSubnetMask = Array("255.255.255.0")

For Each objNetAdapter in colNetAdapters
    errEnable = objNetAdapter.EnableStatic(strIPAddress, strSubnetMask)
Next
  


