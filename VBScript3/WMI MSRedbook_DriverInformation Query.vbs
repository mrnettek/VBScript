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
   Set colItems = objWMIService.ExecQuery("SELECT * FROM MSRedbook_DriverInformation", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
      WScript.Echo "Active: " & objItem.Active
      WScript.Echo "CDDAAccurate: " & objItem.CDDAAccurate
      WScript.Echo "CDDASupported: " & objItem.CDDASupported
      WScript.Echo "InstanceName: " & objItem.InstanceName
      WScript.Echo "MaximumSectorsPerRead: " & objItem.MaximumSectorsPerRead
      WScript.Echo "NumberOfBuffers: " & objItem.NumberOfBuffers
      WScript.Echo "PlayEnabled: " & objItem.PlayEnabled
      WScript.Echo "Reserved1: " & objItem.Reserved1
      WScript.Echo "SectorsPerRead: " & objItem.SectorsPerRead
      WScript.Echo "SectorsPerReadMask: " & objItem.SectorsPerReadMask
      WScript.Echo
   Next
Next

