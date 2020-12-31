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
   Set colItems = objWMIService.ExecQuery("SELECT * FROM MS1394_PortDriverInformation", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
      WScript.Echo "Active: " & objItem.Active
      WScript.Echo "Capabilities: " & objItem.Capabilities
      WScript.Echo "ControllerEUI: " & objItem.ControllerEUI
      WScript.Echo "DeciVoltsSupplied: " & objItem.DeciVoltsSupplied
      WScript.Echo "DeciWattsSupplied: " & objItem.DeciWattsSupplied
      WScript.Echo "GeneralAsyncReceiveRequestBufferSize: " & objItem.GeneralAsyncReceiveRequestBufferSize
      WScript.Echo "GeneralAsyncReceiveResponseBufferSize: " & objItem.GeneralAsyncReceiveResponseBufferSize
      WScript.Echo "InstanceName: " & objItem.InstanceName
      WScript.Echo "LinkSpeed: " & objItem.LinkSpeed
      WScript.Echo "MaxAsyncReadPacket: " & objItem.MaxAsyncReadPacket
      WScript.Echo "MaxAsyncWritePacket: " & objItem.MaxAsyncWritePacket
      WScript.Echo "NumberOfIsochRxDmaContexts: " & objItem.NumberOfIsochRxDmaContexts
      WScript.Echo "NumberOfIsochTxDmaContexts: " & objItem.NumberOfIsochTxDmaContexts
      WScript.Echo "NumberOfPhysicalPorts: " & objItem.NumberOfPhysicalPorts
      WScript.Echo "NumberOfResponseWorkers: " & objItem.NumberOfResponseWorkers
      WScript.Echo "NumberOfTransmitWorkers: " & objItem.NumberOfTransmitWorkers
      WScript.Echo "PhySpeed: " & objItem.PhySpeed
      WScript.Echo "Reserved1: " & objItem.Reserved1
      WScript.Echo
   Next
Next

