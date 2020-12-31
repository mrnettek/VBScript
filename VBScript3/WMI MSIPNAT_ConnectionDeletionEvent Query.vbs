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
   Set colItems = objWMIService.ExecQuery("SELECT * FROM MSIPNAT_ConnectionDeletionEvent", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
      WScript.Echo "Active: " & objItem.Active
      WScript.Echo "InboundConnection: " & objItem.InboundConnection
      WScript.Echo "InstanceName: " & objItem.InstanceName
      WScript.Echo "LocalAddress: " & objItem.LocalAddress
      WScript.Echo "LocalPort: " & objItem.LocalPort
      WScript.Echo "Protocol: " & objItem.Protocol
      WScript.Echo "RemoteAddress: " & objItem.RemoteAddress
      WScript.Echo "RemotePort: " & objItem.RemotePort
      strSECURITY_DESCRIPTOR = Join(objItem.SECURITY_DESCRIPTOR, ",")
         WScript.Echo "SECURITY_DESCRIPTOR: " & strSECURITY_DESCRIPTOR
      WScript.Echo "TIME_CREATED: " & objItem.TIME_CREATED
      WScript.Echo
   Next
Next

