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
   Set colItems = objWMIService.ExecQuery("SELECT * FROM MSWmi_MofData", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
      WScript.Echo "Active: " & objItem.Active
      strBinaryMofData = Join(objItem.BinaryMofData, ",")
         WScript.Echo "BinaryMofData: " & strBinaryMofData
      WScript.Echo "InstanceName: " & objItem.InstanceName
      WScript.Echo "Size: " & objItem.Size
      WScript.Echo "Unused1: " & objItem.Unused1
      WScript.Echo "Unused2: " & objItem.Unused2
      WScript.Echo "Unused4: " & objItem.Unused4
      WScript.Echo
   Next
Next

