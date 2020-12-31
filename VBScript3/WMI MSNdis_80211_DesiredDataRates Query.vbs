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
   Set colItems = objWMIService.ExecQuery("SELECT * FROM MSNdis_80211_DesiredDataRates", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
      WScript.Echo "Active: " & objItem.Active
      WScript.Echo "InstanceName: " & objItem.InstanceName
      strNdis80211DesiredRate = Join(objItem.Ndis80211DesiredRate, ",")
         WScript.Echo "Ndis80211DesiredRate: " & strNdis80211DesiredRate
      WScript.Echo
   Next
Next

