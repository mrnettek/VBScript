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
   Set colItems = objWMIService.ExecQuery("SELECT * FROM BinaryMessageEncodingBindingElement", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
      WScript.Echo "MaxReadPoolSize: " & objItem.MaxReadPoolSize
      WScript.Echo "MaxSessionSize: " & objItem.MaxSessionSize
      WScript.Echo "MaxWritePoolSize: " & objItem.MaxWritePoolSize
      WScript.Echo "MessageVersion: " & objItem.MessageVersion
      WScript.Echo "ReaderQuotas: " & objItem.ReaderQuotas
      WScript.Echo
   Next
Next

