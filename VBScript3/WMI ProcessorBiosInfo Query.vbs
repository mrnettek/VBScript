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
   Set colItems = objWMIService.ExecQuery("SELECT * FROM ProcessorBiosInfo", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
      WScript.Echo "Active: " & objItem.Active
      WScript.Echo "ApicId: " & objItem.ApicId
      WScript.Echo "InstanceName: " & objItem.InstanceName
      WScript.Echo "NtNumber: " & objItem.NtNumber
      WScript.Echo "PBlk: " & objItem.PBlk
      WScript.Echo "PBlkLen: " & objItem.PBlkLen
      WScript.Echo "Pct: " & objItem.Pct
      WScript.Echo "ProcessorId: " & objItem.ProcessorId
      WScript.Echo "Pss: " & objItem.Pss
      WScript.Echo
   Next
Next

