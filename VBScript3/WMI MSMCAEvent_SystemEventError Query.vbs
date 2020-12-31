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
   Set colItems = objWMIService.ExecQuery("SELECT * FROM MSMCAEvent_SystemEventError", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
      WScript.Echo "Active: " & objItem.Active
      WScript.Echo "AdditionalErrors: " & objItem.AdditionalErrors
      WScript.Echo "Cpu: " & objItem.Cpu
      WScript.Echo "ErrorSeverity: " & objItem.ErrorSeverity
      WScript.Echo "InstanceName: " & objItem.InstanceName
      strRawRecord = Join(objItem.RawRecord, ",")
         WScript.Echo "RawRecord: " & strRawRecord
      WScript.Echo "RecordId: " & objItem.RecordId
      strSECURITY_DESCRIPTOR = Join(objItem.SECURITY_DESCRIPTOR, ",")
         WScript.Echo "SECURITY_DESCRIPTOR: " & strSECURITY_DESCRIPTOR
      WScript.Echo "SEL_DATA1: " & objItem.SEL_DATA1
      WScript.Echo "SEL_DATA2: " & objItem.SEL_DATA2
      WScript.Echo "SEL_DATA3: " & objItem.SEL_DATA3
      WScript.Echo "SEL_EVENT_DIR_TYPE: " & objItem.SEL_EVENT_DIR_TYPE
      WScript.Echo "SEL_EVM_REV: " & objItem.SEL_EVM_REV
      WScript.Echo "SEL_GENERATOR_ID: " & objItem.SEL_GENERATOR_ID
      WScript.Echo "SEL_RECORD_ID: " & objItem.SEL_RECORD_ID
      WScript.Echo "SEL_RECORD_TYPE: " & objItem.SEL_RECORD_TYPE
      WScript.Echo "SEL_SENSOR_NUM: " & objItem.SEL_SENSOR_NUM
      WScript.Echo "SEL_SENSOR_TYPE: " & objItem.SEL_SENSOR_TYPE
      WScript.Echo "SEL_TIME_STAMP: " & objItem.SEL_TIME_STAMP
      WScript.Echo "Size: " & objItem.Size
      WScript.Echo "TIME_CREATED: " & objItem.TIME_CREATED
      WScript.Echo "Type: " & objItem.Type
      WScript.Echo "VALIDATION_BITS: " & objItem.VALIDATION_BITS
      WScript.Echo
   Next
Next

