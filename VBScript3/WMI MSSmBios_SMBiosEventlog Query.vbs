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
   Set colItems = objWMIService.ExecQuery("SELECT * FROM MSSmBios_SMBiosEventlog", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
      WScript.Echo "AccessMethod: " & objItem.AccessMethod
      WScript.Echo "AccessMethodAddress: " & objItem.AccessMethodAddress
      WScript.Echo "Active: " & objItem.Active
      WScript.Echo "InstanceName: " & objItem.InstanceName
      WScript.Echo "LengthEachLogTypeDesc: " & objItem.LengthEachLogTypeDesc
      strListLogTypeDesc = Join(objItem.ListLogTypeDesc, ",")
         WScript.Echo "ListLogTypeDesc: " & strListLogTypeDesc
      strLogArea = Join(objItem.LogArea, ",")
         WScript.Echo "LogArea: " & strLogArea
      WScript.Echo "LogAreaLength: " & objItem.LogAreaLength
      WScript.Echo "LogChangeToken: " & objItem.LogChangeToken
      WScript.Echo "LogDataStart: " & objItem.LogDataStart
      WScript.Echo "LogHeaderDescExists: " & objItem.LogHeaderDescExists
      WScript.Echo "LogHeaderFormat: " & objItem.LogHeaderFormat
      WScript.Echo "LogHeaderStart: " & objItem.LogHeaderStart
      WScript.Echo "LogStatus: " & objItem.LogStatus
      WScript.Echo "LogTypeDescLength: " & objItem.LogTypeDescLength
      WScript.Echo "NumberLogTypeDesc: " & objItem.NumberLogTypeDesc
      WScript.Echo "Reserved: " & objItem.Reserved
      WScript.Echo
   Next
Next

