On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_NetworkProtocol",,48)
For Each objItem in colItems
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "ConnectionlessService: " & objItem.ConnectionlessService
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "GuaranteesDelivery: " & objItem.GuaranteesDelivery
    Wscript.Echo "GuaranteesSequencing: " & objItem.GuaranteesSequencing
    Wscript.Echo "InstallDate: " & objItem.InstallDate
    Wscript.Echo "MaximumAddressSize: " & objItem.MaximumAddressSize
    Wscript.Echo "MaximumMessageSize: " & objItem.MaximumMessageSize
    Wscript.Echo "MessageOriented: " & objItem.MessageOriented
    Wscript.Echo "MinimumAddressSize: " & objItem.MinimumAddressSize
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "PseudoStreamOriented: " & objItem.PseudoStreamOriented
    Wscript.Echo "Status: " & objItem.Status
    Wscript.Echo "SupportsBroadcasting: " & objItem.SupportsBroadcasting
    Wscript.Echo "SupportsConnectData: " & objItem.SupportsConnectData
    Wscript.Echo "SupportsDisconnectData: " & objItem.SupportsDisconnectData
    Wscript.Echo "SupportsEncryption: " & objItem.SupportsEncryption
    Wscript.Echo "SupportsExpeditedData: " & objItem.SupportsExpeditedData
    Wscript.Echo "SupportsFragmentation: " & objItem.SupportsFragmentation
    Wscript.Echo "SupportsGracefulClosing: " & objItem.SupportsGracefulClosing
    Wscript.Echo "SupportsGuaranteedBandwidth: " & objItem.SupportsGuaranteedBandwidth
    Wscript.Echo "SupportsMulticasting: " & objItem.SupportsMulticasting
    Wscript.Echo "SupportsQualityofService: " & objItem.SupportsQualityofService
Next

