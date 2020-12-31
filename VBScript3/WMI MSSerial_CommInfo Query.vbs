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
   Set colItems = objWMIService.ExecQuery("SELECT * FROM MSSerial_CommInfo", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
      WScript.Echo "Active: " & objItem.Active
      WScript.Echo "BaudRate: " & objItem.BaudRate
      WScript.Echo "BitsPerByte: " & objItem.BitsPerByte
      WScript.Echo "InstanceName: " & objItem.InstanceName
      WScript.Echo "IsBusy: " & objItem.IsBusy
      WScript.Echo "MaximumBaudRate: " & objItem.MaximumBaudRate
      WScript.Echo "MaximumInputBufferSize: " & objItem.MaximumInputBufferSize
      WScript.Echo "MaximumOutputBufferSize: " & objItem.MaximumOutputBufferSize
      WScript.Echo "Parity: " & objItem.Parity
      WScript.Echo "ParityCheckEnable: " & objItem.ParityCheckEnable
      WScript.Echo "SettableBaudRate: " & objItem.SettableBaudRate
      WScript.Echo "SettableDataBits: " & objItem.SettableDataBits
      WScript.Echo "SettableFlowControl: " & objItem.SettableFlowControl
      WScript.Echo "SettableParity: " & objItem.SettableParity
      WScript.Echo "SettableParityCheck: " & objItem.SettableParityCheck
      WScript.Echo "SettableStopBits: " & objItem.SettableStopBits
      WScript.Echo "StopBits: " & objItem.StopBits
      WScript.Echo "Support16BitMode: " & objItem.Support16BitMode
      WScript.Echo "SupportDTRDSR: " & objItem.SupportDTRDSR
      WScript.Echo "SupportIntervalTimeouts: " & objItem.SupportIntervalTimeouts
      WScript.Echo "SupportParityCheck: " & objItem.SupportParityCheck
      WScript.Echo "SupportRTSCTS: " & objItem.SupportRTSCTS
      WScript.Echo "SupportXonXoff: " & objItem.SupportXonXoff
      WScript.Echo "XoffCharacter: " & objItem.XoffCharacter
      WScript.Echo "XoffXmitThreshold: " & objItem.XoffXmitThreshold
      WScript.Echo "XonCharacter: " & objItem.XonCharacter
      WScript.Echo "XonXmitThreshold: " & objItem.XonXmitThreshold
      WScript.Echo
   Next
Next

