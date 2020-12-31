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
   Set colItems = objWMIService.ExecQuery("SELECT * FROM MSSerial_CommProperties", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
      WScript.Echo "Active: " & objItem.Active
      WScript.Echo "dwCurrentRxQueue: " & objItem.dwCurrentRxQueue
      WScript.Echo "dwCurrentTxQueue: " & objItem.dwCurrentTxQueue
      WScript.Echo "dwMaxBaud: " & objItem.dwMaxBaud
      WScript.Echo "dwMaxRxQueue: " & objItem.dwMaxRxQueue
      WScript.Echo "dwMaxTxQueue: " & objItem.dwMaxTxQueue
      WScript.Echo "dwProvCapabilities: " & objItem.dwProvCapabilities
      WScript.Echo "dwProvCharSize: " & objItem.dwProvCharSize
      WScript.Echo "dwProvSpec1: " & objItem.dwProvSpec1
      WScript.Echo "dwProvSpec2: " & objItem.dwProvSpec2
      WScript.Echo "dwProvSubType: " & objItem.dwProvSubType
      WScript.Echo "dwReserved1: " & objItem.dwReserved1
      WScript.Echo "dwServiceMask: " & objItem.dwServiceMask
      WScript.Echo "dwSettableBaud: " & objItem.dwSettableBaud
      WScript.Echo "dwSettableParams: " & objItem.dwSettableParams
      WScript.Echo "InstanceName: " & objItem.InstanceName
      strwcProvChar = Join(objItem.wcProvChar, ",")
         WScript.Echo "wcProvChar: " & strwcProvChar
      WScript.Echo "wPacketLength: " & objItem.wPacketLength
      WScript.Echo "wPacketVersion: " & objItem.wPacketVersion
      WScript.Echo "wSettableData: " & objItem.wSettableData
      WScript.Echo "wSettableStopParity: " & objItem.wSettableStopParity
      WScript.Echo
   Next
Next

