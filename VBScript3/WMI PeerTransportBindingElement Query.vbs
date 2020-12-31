On Error Resume Next

Const wbemFlagReturnImmediately = &h10
Const wbemFlagForwardOnly = &h20

arrComputers = Array(".")
For Each strComputer In arrComputers
   WScript.Echo
   WScript.Echo "=========================================="
   WScript.Echo "Computer: " & strComputer
   WScript.Echo "=========================================="

   Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\ServiceModel")
   Set colItems = objWMIService.ExecQuery("SELECT * FROM PeerTransportBindingElement", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
      WScript.Echo "ListenIPAddress: " & objItem.ListenIPAddress
      WScript.Echo "ManualAddressing: " & objItem.ManualAddressing
      WScript.Echo "MaxBufferPoolSize: " & objItem.MaxBufferPoolSize
      WScript.Echo "MaxReceivedMessageSize: " & objItem.MaxReceivedMessageSize
      WScript.Echo "Port: " & objItem.Port
      WScript.Echo "Scheme: " & objItem.Scheme
      WScript.Echo "Security: " & objItem.Security
      WScript.Echo
   Next
Next

