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
   Set colItems = objWMIService.ExecQuery("SELECT * FROM CIM_ProcessExecutable", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
      WScript.Echo "Antecedent: " & objItem.Antecedent
      WScript.Echo "BaseAddress: " & objItem.BaseAddress
      WScript.Echo "Dependent: " & objItem.Dependent
      WScript.Echo "GlobalProcessCount: " & objItem.GlobalProcessCount
      WScript.Echo "ModuleInstance: " & objItem.ModuleInstance
      WScript.Echo "ProcessCount: " & objItem.ProcessCount
      WScript.Echo
   Next
Next

