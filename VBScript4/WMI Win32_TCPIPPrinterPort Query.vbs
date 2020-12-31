On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_TCPIPPrinterPort",,48)
For Each objItem in colItems
    Wscript.Echo "ByteCount: " & objItem.ByteCount
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "CreationClassName: " & objItem.CreationClassName
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "HostAddress: " & objItem.HostAddress
    Wscript.Echo "InstallDate: " & objItem.InstallDate
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "PortNumber: " & objItem.PortNumber
    Wscript.Echo "Protocol: " & objItem.Protocol
    Wscript.Echo "Queue: " & objItem.Queue
    Wscript.Echo "SNMPCommunity: " & objItem.SNMPCommunity
    Wscript.Echo "SNMPDevIndex: " & objItem.SNMPDevIndex
    Wscript.Echo "SNMPEnabled: " & objItem.SNMPEnabled
    Wscript.Echo "Status: " & objItem.Status
    Wscript.Echo "SystemCreationClassName: " & objItem.SystemCreationClassName
    Wscript.Echo "SystemName: " & objItem.SystemName
    Wscript.Echo "Type: " & objItem.Type
Next

