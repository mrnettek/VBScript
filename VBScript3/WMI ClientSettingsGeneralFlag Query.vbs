On Error Resume Next

Const wbemFlagReturnImmediately = &h10
Const wbemFlagForwardOnly = &h20

arrComputers = Array(".")
For Each strComputer In arrComputers
   WScript.Echo
   WScript.Echo "=========================================="
   WScript.Echo "Computer: " & strComputer
   WScript.Echo "=========================================="

   Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\Microsoft\SqlServer\ComputerManagement10")
   Set colItems = objWMIService.ExecQuery("SELECT * FROM ClientSettingsGeneralFlag", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
      WScript.Echo "FlagIndex: " & objItem.FlagIndex
      WScript.Echo "FlagName: " & objItem.FlagName
      WScript.Echo "FlagValue: " & objItem.FlagValue
      WScript.Echo "InstanceName: " & objItem.InstanceName
      WScript.Echo
   Next
Next

