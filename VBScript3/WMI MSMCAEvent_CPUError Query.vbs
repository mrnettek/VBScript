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
   Set colItems = objWMIService.ExecQuery("SELECT * FROM MSMCAEvent_CPUError", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
      WScript.Echo "Active: " & objItem.Active
      WScript.Echo "AdditionalErrors: " & objItem.AdditionalErrors
      WScript.Echo "Cpu: " & objItem.Cpu
      WScript.Echo "ErrorSeverity: " & objItem.ErrorSeverity
      WScript.Echo "InstanceName: " & objItem.InstanceName
      WScript.Echo "Level: " & objItem.Level
      strRawRecord = Join(objItem.RawRecord, ",")
         WScript.Echo "RawRecord: " & strRawRecord
      WScript.Echo "RecordId: " & objItem.RecordId
      strSECURITY_DESCRIPTOR = Join(objItem.SECURITY_DESCRIPTOR, ",")
         WScript.Echo "SECURITY_DESCRIPTOR: " & strSECURITY_DESCRIPTOR
      WScript.Echo "Size: " & objItem.Size
      WScript.Echo "TIME_CREATED: " & objItem.TIME_CREATED
      WScript.Echo "Type: " & objItem.Type
      WScript.Echo
   Next
Next

