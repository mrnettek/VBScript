' Description: Enables NetBIOS for a network adapter. To enable NetBIOS via DHCP, pass the value 0 to the SetTCPIPNetBIOS method (instead of the value 1). Pass the value 2 to disable NetBIOS.


On Error Resume Next

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colNetCards = objWMIService.ExecQuery _
    ("Select * From Win32_NetworkAdapterConfiguration Where IPEnabled = True")

For Each objNetCard in colNetCards
    objNetCard.SetTCPIPNetBIOS(1)
Next

