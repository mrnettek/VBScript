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
   Set colItems = objWMIService.ExecQuery("SELECT * FROM MSMCAEvent_PCIBusError", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
      WScript.Echo "Active: " & objItem.Active
      WScript.Echo "AdditionalErrors: " & objItem.AdditionalErrors
      WScript.Echo "Cpu: " & objItem.Cpu
      WScript.Echo "ErrorSeverity: " & objItem.ErrorSeverity
      WScript.Echo "InstanceName: " & objItem.InstanceName
      WScript.Echo "PCI_BUS_ADDRESS: " & objItem.PCI_BUS_ADDRESS
      WScript.Echo "PCI_BUS_CMD: " & objItem.PCI_BUS_CMD
      WScript.Echo "PCI_BUS_DATA: " & objItem.PCI_BUS_DATA
      WScript.Echo "PCI_BUS_ERROR_STATUS: " & objItem.PCI_BUS_ERROR_STATUS
      WScript.Echo "PCI_BUS_ERROR_TYPE: " & objItem.PCI_BUS_ERROR_TYPE
      WScript.Echo "PCI_BUS_ID_BusNumber: " & objItem.PCI_BUS_ID_BusNumber
      WScript.Echo "PCI_BUS_ID_SegmentNumber: " & objItem.PCI_BUS_ID_SegmentNumber
      WScript.Echo "PCI_BUS_REQUESTOR_ID: " & objItem.PCI_BUS_REQUESTOR_ID
      WScript.Echo "PCI_BUS_RESPONDER_ID: " & objItem.PCI_BUS_RESPONDER_ID
      WScript.Echo "PCI_BUS_TARGET_ID: " & objItem.PCI_BUS_TARGET_ID
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

