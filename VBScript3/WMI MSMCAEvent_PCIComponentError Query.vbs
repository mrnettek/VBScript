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
   Set colItems = objWMIService.ExecQuery("SELECT * FROM MSMCAEvent_PCIComponentError", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
      WScript.Echo "Active: " & objItem.Active
      WScript.Echo "AdditionalErrors: " & objItem.AdditionalErrors
      WScript.Echo "Cpu: " & objItem.Cpu
      WScript.Echo "ErrorSeverity: " & objItem.ErrorSeverity
      WScript.Echo "InstanceName: " & objItem.InstanceName
      WScript.Echo "PCI_COMP_ERROR_STATUS: " & objItem.PCI_COMP_ERROR_STATUS
      WScript.Echo "PCI_COMP_INFO_BusNumber: " & objItem.PCI_COMP_INFO_BusNumber
      WScript.Echo "PCI_COMP_INFO_ClassCodeBaseClass: " & objItem.PCI_COMP_INFO_ClassCodeBaseClass
      WScript.Echo "PCI_COMP_INFO_ClassCodeInterface: " & objItem.PCI_COMP_INFO_ClassCodeInterface
      WScript.Echo "PCI_COMP_INFO_ClassCodeSubClass: " & objItem.PCI_COMP_INFO_ClassCodeSubClass
      WScript.Echo "PCI_COMP_INFO_DeviceId: " & objItem.PCI_COMP_INFO_DeviceId
      WScript.Echo "PCI_COMP_INFO_DeviceNumber: " & objItem.PCI_COMP_INFO_DeviceNumber
      WScript.Echo "PCI_COMP_INFO_FunctionNumber: " & objItem.PCI_COMP_INFO_FunctionNumber
      WScript.Echo "PCI_COMP_INFO_SegmentNumber: " & objItem.PCI_COMP_INFO_SegmentNumber
      WScript.Echo "PCI_COMP_INFO_VendorId: " & objItem.PCI_COMP_INFO_VendorId
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

