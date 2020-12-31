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
   Set colItems = objWMIService.ExecQuery("SELECT * FROM WebScriptEnablingBehavior", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
      WScript.Echo "DefaultBodyStyle: " & objItem.DefaultBodyStyle
      WScript.Echo "DefaultOutgoingRequestFormat: " & objItem.DefaultOutgoingRequestFormat
      WScript.Echo "DefaultOutgoingResponseFormat: " & objItem.DefaultOutgoingResponseFormat
      WScript.Echo "Type: " & objItem.Type
      WScript.Echo
   Next
Next

