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
   Set colItems = objWMIService.ExecQuery("SELECT * FROM HWConfig_PhyDisk", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
      WScript.Echo "BytesPerSector: " & objItem.BytesPerSector
      WScript.Echo "Cylinders: " & objItem.Cylinders
      WScript.Echo "DiskNumber: " & objItem.DiskNumber
      strEventGuid = Join(objItem.EventGuid, ",")
         WScript.Echo "EventGuid: " & strEventGuid
      WScript.Echo "EventSize: " & objItem.EventSize
      WScript.Echo "EventType: " & objItem.EventType
      WScript.Echo "Flags: " & objItem.Flags
      WScript.Echo "InstanceId: " & objItem.InstanceId
      WScript.Echo "KernelTime: " & objItem.KernelTime
      WScript.Echo "Manufacturer: " & objItem.Manufacturer
      WScript.Echo "MofData: " & objItem.MofData
      WScript.Echo "MofLength: " & objItem.MofLength
      strParentGuid = Join(objItem.ParentGuid, ",")
         WScript.Echo "ParentGuid: " & strParentGuid
      WScript.Echo "ParentInstanceId: " & objItem.ParentInstanceId
      WScript.Echo "ReservedHeaderField: " & objItem.ReservedHeaderField
      WScript.Echo "SCSILun: " & objItem.SCSILun
      WScript.Echo "SCSIPath: " & objItem.SCSIPath
      WScript.Echo "SCSIPort: " & objItem.SCSIPort
      WScript.Echo "SCSITarget: " & objItem.SCSITarget
      WScript.Echo "SectorsPerTrack: " & objItem.SectorsPerTrack
      WScript.Echo "ThreadId: " & objItem.ThreadId
      WScript.Echo "TimeStamp: " & objItem.TimeStamp
      WScript.Echo "TraceLevel: " & objItem.TraceLevel
      WScript.Echo "TraceVersion: " & objItem.TraceVersion
      WScript.Echo "TracksPerCylinder: " & objItem.TracksPerCylinder
      WScript.Echo "UserTime: " & objItem.UserTime
      WScript.Echo
   Next
Next

