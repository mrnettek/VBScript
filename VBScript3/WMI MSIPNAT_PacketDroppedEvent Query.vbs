On Error Resume Next

Const wbemFlagReturnImmediately = &h10
Const wbemFlagForwardOnly = &h20

arrComputers = Array(".")
For Each strComputer In arrComputers
   WScript.Echo
   WScript.Echo "=========================================="
   WScript.Echo "Computer: " & strComputer
   WScript.Echo "=========================================="

   Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\WMI")
   Set colItems = objWMIService.ExecQuery("SELECT * FROM MSIPNAT_PacketDroppedEvent", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
      WScript.Echo "Active: " & objItem.Active
      WScript.Echo "DestinationAddress: " & objItem.DestinationAddress
      WScript.Echo "DestinationIdentifier: " & objItem.DestinationIdentifier
      WScript.Echo "InstanceName: " & objItem.InstanceName
      WScript.Echo "PacketPath: " & objItem.PacketPath
      WScript.Echo "PacketSize: " & objItem.PacketSize
      WScript.Echo "Protocol: " & objItem.Protocol
      WScript.Echo "ProtocolData1: " & objItem.ProtocolData1
      WScript.Echo "ProtocolData2: " & objItem.ProtocolData2
      WScript.Echo "ProtocolData3: " & objItem.ProtocolData3
      WScript.Echo "ProtocolData4: " & objItem.ProtocolData4
      strSECURITY_DESCRIPTOR = Join(objItem.SECURITY_DESCRIPTOR, ",")
         WScript.Echo "SECURITY_DESCRIPTOR: " & strSECURITY_DESCRIPTOR
      WScript.Echo "SourceAddress: " & objItem.SourceAddress
      WScript.Echo "SourceIdentifier: " & objItem.SourceIdentifier
      WScript.Echo "TIME_CREATED: " & objItem.TIME_CREATED
      WScript.Echo
   Next
Next

