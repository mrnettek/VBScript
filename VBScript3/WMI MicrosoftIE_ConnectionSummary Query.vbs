On Error Resume Next

Const wbemFlagReturnImmediately = &h10
Const wbemFlagForwardOnly = &h20

arrComputers = Array(".")
For Each strComputer In arrComputers
   WScript.Echo
   WScript.Echo "=========================================="
   WScript.Echo "Computer: " & strComputer
   WScript.Echo "=========================================="

   Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2\Applications\MicrosoftIE")
   Set colItems = objWMIService.ExecQuery("SELECT * FROM MicrosoftIE_ConnectionSummary", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
      WScript.Echo "Caption: " & objItem.Caption
      WScript.Echo "ConnectionPreference: " & objItem.ConnectionPreference
      WScript.Echo "Description: " & objItem.Description
      WScript.Echo "EnableHttp11: " & objItem.EnableHttp11
      WScript.Echo "ProxyHttp11: " & objItem.ProxyHttp11
      WScript.Echo "SettingID: " & objItem.SettingID
      WScript.Echo
   Next
Next

