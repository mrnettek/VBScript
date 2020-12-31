On Error Resume Next

Const wbemFlagReturnImmediately = &h10
Const wbemFlagForwardOnly = &h20

arrComputers = Array(".")
For Each strComputer In arrComputers
   WScript.Echo
   WScript.Echo "=========================================="
   WScript.Echo "Computer: " & strComputer
   WScript.Echo "=========================================="

   Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2")
   Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_DiskQuota", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
      WScript.Echo "DiskSpaceUsed: " & objItem.DiskSpaceUsed
      WScript.Echo "Limit: " & objItem.Limit
      WScript.Echo "QuotaVolume: " & objItem.QuotaVolume
      WScript.Echo "Status: " & objItem.Status
      WScript.Echo "User: " & objItem.User
      WScript.Echo "WarningLimit: " & objItem.WarningLimit
      WScript.Echo
   Next
Next

