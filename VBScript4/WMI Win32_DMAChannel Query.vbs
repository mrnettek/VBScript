On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_DMAChannel",,48)
For Each objItem in colItems
    Wscript.Echo "AddressSize: " & objItem.AddressSize
    Wscript.Echo "Availability: " & objItem.Availability
    Wscript.Echo "BurstMode: " & objItem.BurstMode
    Wscript.Echo "ByteMode: " & objItem.ByteMode
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "ChannelTiming: " & objItem.ChannelTiming
    Wscript.Echo "CreationClassName: " & objItem.CreationClassName
    Wscript.Echo "CSCreationClassName: " & objItem.CSCreationClassName
    Wscript.Echo "CSName: " & objItem.CSName
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "DMAChannel: " & objItem.DMAChannel
    Wscript.Echo "InstallDate: " & objItem.InstallDate
    Wscript.Echo "MaxTransferSize: " & objItem.MaxTransferSize
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "Port: " & objItem.Port
    Wscript.Echo "Status: " & objItem.Status
    Wscript.Echo "TransferWidths: " & objItem.TransferWidths
    Wscript.Echo "TypeCTiming: " & objItem.TypeCTiming
    Wscript.Echo "WordMode: " & objItem.WordMode
Next

