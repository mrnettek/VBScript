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
   Set colItems = objWMIService.ExecQuery("SELECT * FROM MSTapeDriveProblemEvent", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
      WScript.Echo "Active: " & objItem.Active
      WScript.Echo "DriveProblemType: " & objItem.DriveProblemType
      WScript.Echo "InstanceName: " & objItem.InstanceName
      strSECURITY_DESCRIPTOR = Join(objItem.SECURITY_DESCRIPTOR, ",")
         WScript.Echo "SECURITY_DESCRIPTOR: " & strSECURITY_DESCRIPTOR
      strTapeData = Join(objItem.TapeData, ",")
         WScript.Echo "TapeData: " & strTapeData
      WScript.Echo "TIME_CREATED: " & objItem.TIME_CREATED
      WScript.Echo
   Next
Next

