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
   Set colItems = objWMIService.ExecQuery("SELECT * FROM Channel", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
      WScript.Echo "LocalAddress: " & objItem.LocalAddress
      WScript.Echo "RemoteAddress: " & objItem.RemoteAddress
      WScript.Echo "RemoteEndpoint: " & objItem.RemoteEndpoint
      WScript.Echo "SessionId: " & objItem.SessionId
      WScript.Echo "Type: " & objItem.Type
      WScript.Echo
   Next
Next

