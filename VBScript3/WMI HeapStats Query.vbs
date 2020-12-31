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
   Set colItems = objWMIService.ExecQuery("SELECT * FROM HeapStats", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
      strEventGuid = Join(objItem.EventGuid, ",")
         WScript.Echo "EventGuid: " & strEventGuid
      WScript.Echo "EventSize: " & objItem.EventSize
      WScript.Echo "EventType: " & objItem.EventType
      WScript.Echo "FinalizationPromotedCount: " & objItem.FinalizationPromotedCount
      WScript.Echo "FinalizationPromotedSize: " & objItem.FinalizationPromotedSize
      WScript.Echo "GCHandleCount: " & objItem.GCHandleCount
      WScript.Echo "GenerationSize0: " & objItem.GenerationSize0
      WScript.Echo "GenerationSize1: " & objItem.GenerationSize1
      WScript.Echo "GenerationSize2: " & objItem.GenerationSize2
      WScript.Echo "GenerationSize3: " & objItem.GenerationSize3
      WScript.Echo "InstanceId: " & objItem.InstanceId
      WScript.Echo "KernelTime: " & objItem.KernelTime
      WScript.Echo "MofData: " & objItem.MofData
      WScript.Echo "MofLength: " & objItem.MofLength
      strParentGuid = Join(objItem.ParentGuid, ",")
         WScript.Echo "ParentGuid: " & strParentGuid
      WScript.Echo "ParentInstanceId: " & objItem.ParentInstanceId
      WScript.Echo "PinnedObjectCount: " & objItem.PinnedObjectCount
      WScript.Echo "ReservedHeaderField: " & objItem.ReservedHeaderField
      WScript.Echo "SinkBlockCount: " & objItem.SinkBlockCount
      WScript.Echo "ThreadId: " & objItem.ThreadId
      WScript.Echo "TimeStamp: " & objItem.TimeStamp
      WScript.Echo "TotalPromotedSize0: " & objItem.TotalPromotedSize0
      WScript.Echo "TotalPromotedSize1: " & objItem.TotalPromotedSize1
      WScript.Echo "TotalPromotedSize2: " & objItem.TotalPromotedSize2
      WScript.Echo "TotalPromotedSize3: " & objItem.TotalPromotedSize3
      WScript.Echo "TraceLevel: " & objItem.TraceLevel
      WScript.Echo "TraceVersion: " & objItem.TraceVersion
      WScript.Echo "UserTime: " & objItem.UserTime
      WScript.Echo
   Next
Next

