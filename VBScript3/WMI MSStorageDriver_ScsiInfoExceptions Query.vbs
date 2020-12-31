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
   Set colItems = objWMIService.ExecQuery("SELECT * FROM MSStorageDriver_ScsiInfoExceptions", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
      WScript.Echo "Active: " & objItem.Active
      WScript.Echo "Flags: " & objItem.Flags
      WScript.Echo "InstanceName: " & objItem.InstanceName
      WScript.Echo "IntervalTimer: " & objItem.IntervalTimer
      WScript.Echo "MRIE: " & objItem.MRIE
      WScript.Echo "Padding: " & objItem.Padding
      WScript.Echo "PageSavable: " & objItem.PageSavable
      WScript.Echo "ReportCount: " & objItem.ReportCount
      WScript.Echo
   Next
Next

