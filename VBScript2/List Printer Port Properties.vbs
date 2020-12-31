' Description: Lists properties of all the TCP/IP printer ports installed on a computer.


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colPorts =  objWMIService.ExecQuery _
    ("Select * from Win32_TCPIPPrinterPort")

For Each objPort in colPorts
    Wscript.Echo "Description: " & objPort.Description
    Wscript.Echo "Host Address: " & objPort.HostAddress
    Wscript.Echo "Name: " & objPort.Name
    Wscript.Echo "Port Number: " & objPort.PortNumber
    Wscript.Echo "Protocol: " & objPort.Protocol
    Wscript.Echo "SNMP Community: " & objPort.SNMPCommunity
    Wscript.Echo "SNMP Dev Index: " & objPort.SnMPDevIndex
    Wscript.Echo "SNMP Enabled: " & objPort.SNMPEnabled
Next

