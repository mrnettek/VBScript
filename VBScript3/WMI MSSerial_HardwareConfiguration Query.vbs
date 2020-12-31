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
   Set colItems = objWMIService.ExecQuery("SELECT * FROM MSSerial_HardwareConfiguration", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
      WScript.Echo "Active: " & objItem.Active
      WScript.Echo "BaseIOAddress: " & objItem.BaseIOAddress
      WScript.Echo "InstanceName: " & objItem.InstanceName
      WScript.Echo "InterruptType: " & objItem.InterruptType
      WScript.Echo "IrqAffinityMask: " & objItem.IrqAffinityMask
      WScript.Echo "IrqLevel: " & objItem.IrqLevel
      WScript.Echo "IrqNumber: " & objItem.IrqNumber
      WScript.Echo "IrqVector: " & objItem.IrqVector
      WScript.Echo
   Next
Next

