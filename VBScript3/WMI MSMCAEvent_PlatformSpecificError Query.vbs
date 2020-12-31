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
   Set colItems = objWMIService.ExecQuery("SELECT * FROM MSMCAEvent_PlatformSpecificError", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
      WScript.Echo "Active: " & objItem.Active
      WScript.Echo "AdditionalErrors: " & objItem.AdditionalErrors
      WScript.Echo "Cpu: " & objItem.Cpu
      WScript.Echo "ErrorSeverity: " & objItem.ErrorSeverity
      WScript.Echo "InstanceName: " & objItem.InstanceName
      strOEM_COMPONENT_ID = Join(objItem.OEM_COMPONENT_ID, ",")
         WScript.Echo "OEM_COMPONENT_ID: " & strOEM_COMPONENT_ID
      WScript.Echo "PLATFORM_BUS_SPECIFIC_DATA: " & objItem.PLATFORM_BUS_SPECIFIC_DATA
      WScript.Echo "PLATFORM_ERROR_STATUS: " & objItem.PLATFORM_ERROR_STATUS
      WScript.Echo "PLATFORM_REQUESTOR_ID: " & objItem.PLATFORM_REQUESTOR_ID
      WScript.Echo "PLATFORM_RESPONDER_ID: " & objItem.PLATFORM_RESPONDER_ID
      WScript.Echo "PLATFORM_TARGET_ID: " & objItem.PLATFORM_TARGET_ID
      strRawRecord = Join(objItem.RawRecord, ",")
         WScript.Echo "RawRecord: " & strRawRecord
      WScript.Echo "RecordId: " & objItem.RecordId
      strSECURITY_DESCRIPTOR = Join(objItem.SECURITY_DESCRIPTOR, ",")
         WScript.Echo "SECURITY_DESCRIPTOR: " & strSECURITY_DESCRIPTOR
      WScript.Echo "Size: " & objItem.Size
      WScript.Echo "TIME_CREATED: " & objItem.TIME_CREATED
      WScript.Echo "Type: " & objItem.Type
      WScript.Echo "VALIDATION_BITS: " & objItem.VALIDATION_BITS
      WScript.Echo
   Next
Next

