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
   Set colItems = objWMIService.ExecQuery("SELECT * FROM MSMCAEvent_MemoryError", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
      WScript.Echo "Active: " & objItem.Active
      WScript.Echo "AdditionalErrors: " & objItem.AdditionalErrors
      WScript.Echo "BUS_SPECIFIC_DATA: " & objItem.BUS_SPECIFIC_DATA
      WScript.Echo "Cpu: " & objItem.Cpu
      WScript.Echo "ErrorSeverity: " & objItem.ErrorSeverity
      WScript.Echo "InstanceName: " & objItem.InstanceName
      WScript.Echo "MEM_BANK: " & objItem.MEM_BANK
      WScript.Echo "MEM_BIT_POSITION: " & objItem.MEM_BIT_POSITION
      WScript.Echo "MEM_CARD: " & objItem.MEM_CARD
      WScript.Echo "MEM_COLUMN: " & objItem.MEM_COLUMN
      WScript.Echo "MEM_ERROR_STATUS: " & objItem.MEM_ERROR_STATUS
      WScript.Echo "MEM_MODULE: " & objItem.MEM_MODULE
      WScript.Echo "MEM_NODE: " & objItem.MEM_NODE
      WScript.Echo "MEM_PHYSICAL_ADDR: " & objItem.MEM_PHYSICAL_ADDR
      WScript.Echo "MEM_PHYSICAL_MASK: " & objItem.MEM_PHYSICAL_MASK
      WScript.Echo "MEM_ROW: " & objItem.MEM_ROW
      strRawRecord = Join(objItem.RawRecord, ",")
         WScript.Echo "RawRecord: " & strRawRecord
      WScript.Echo "RecordId: " & objItem.RecordId
      WScript.Echo "REQUESTOR_ID: " & objItem.REQUESTOR_ID
      WScript.Echo "RESPONDER_ID: " & objItem.RESPONDER_ID
      strSECURITY_DESCRIPTOR = Join(objItem.SECURITY_DESCRIPTOR, ",")
         WScript.Echo "SECURITY_DESCRIPTOR: " & strSECURITY_DESCRIPTOR
      WScript.Echo "Size: " & objItem.Size
      WScript.Echo "TARGET_ID: " & objItem.TARGET_ID
      WScript.Echo "TIME_CREATED: " & objItem.TIME_CREATED
      WScript.Echo "Type: " & objItem.Type
      WScript.Echo "VALIDATION_BITS: " & objItem.VALIDATION_BITS
      WScript.Echo "xMEM_DEVICE: " & objItem.xMEM_DEVICE
      WScript.Echo
   Next
Next

