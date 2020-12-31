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
   Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_SCSIControllerDevice", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
      WScript.Echo "AccessState: " & objItem.AccessState
      WScript.Echo "Antecedent: " & objItem.Antecedent
      WScript.Echo "Dependent: " & objItem.Dependent
      WScript.Echo "NegotiatedDataWidth: " & objItem.NegotiatedDataWidth
      WScript.Echo "NegotiatedSpeed: " & objItem.NegotiatedSpeed
      WScript.Echo "NumberOfHardResets: " & objItem.NumberOfHardResets
      WScript.Echo "NumberOfSoftResets: " & objItem.NumberOfSoftResets
      WScript.Echo
   Next
Next

