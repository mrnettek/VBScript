' Description: Enables the DHCP client service on a computer. The computer will then use DHCP to obtain an IP address rather than use a static IP address.


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colNetAdapters = objWMIService.ExecQuery _
    ("Select * from Win32_NetworkAdapterConfiguration where IPEnabled=TRUE")
 
For Each objNetAdapter In colNetAdapters
    errEnable = objNetAdapter.EnableDHCP()
Next

