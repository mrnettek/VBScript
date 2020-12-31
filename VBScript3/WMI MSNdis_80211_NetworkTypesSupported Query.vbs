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
   Set colItems = objWMIService.ExecQuery("SELECT * FROM MSNdis_80211_NetworkTypesSupported", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
      WScript.Echo "Active: " & objItem.Active
      WScript.Echo "InstanceName: " & objItem.InstanceName
      strNdis80211NetworkTypes = Join(objItem.Ndis80211NetworkTypes, ",")
         WScript.Echo "Ndis80211NetworkTypes: " & strNdis80211NetworkTypes
      WScript.Echo "NumberOfItems: " & objItem.NumberOfItems
      WScript.Echo
   Next
Next

