' Description: Configures a TCP/IP-bound network adapter to use two DNS servers: 192.168.1.100 and 192.168.1.200. Note that even if a computer only uses one DNS server, the IP address of that server must still be passed to the SetDNSServerSearchOrder method as an array (in that case, an array with only one element).


On Error Resume Next

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colNetCards = objWMIService.ExecQuery _
    ("Select * From Win32_NetworkAdapterConfiguration Where IPEnabled = True")

For Each objNetCard in colNetCards
    arrDNSServers = Array("192.168.1.100", "192.168.1.200")
    objNetCard.SetDNSServerSearchOrder(arrDNSServers)
Next

