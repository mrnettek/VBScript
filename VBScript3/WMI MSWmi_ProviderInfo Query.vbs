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
   Set colItems = objWMIService.ExecQuery("SELECT * FROM MSWmi_ProviderInfo", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
      WScript.Echo "Active: " & objItem.Active
      WScript.Echo "Description: " & objItem.Description
      WScript.Echo "FriendlyName: " & objItem.FriendlyName
      WScript.Echo "InstanceName: " & objItem.InstanceName
      WScript.Echo "Location: " & objItem.Location
      WScript.Echo "Manufacturer: " & objItem.Manufacturer
      WScript.Echo "Service: " & objItem.Service
      WScript.Echo
   Next
Next

