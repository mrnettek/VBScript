On Error Resume Next

Const wbemFlagReturnImmediately = &h10
Const wbemFlagForwardOnly = &h20

arrComputers = Array(".")
For Each strComputer In arrComputers
   WScript.Echo
   WScript.Echo "=========================================="
   WScript.Echo "Computer: " & strComputer
   WScript.Echo "=========================================="

   Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2")
   Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_PerfRawData_WindowsWorkflowFoundation3000_WindowsWorkflowFoundation", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
      WScript.Echo "Caption: " & objItem.Caption
      WScript.Echo "Description: " & objItem.Description
      WScript.Echo "Frequency_Object: " & objItem.Frequency_Object
      WScript.Echo "Frequency_PerfTime: " & objItem.Frequency_PerfTime
      WScript.Echo "Frequency_Sys100NS: " & objItem.Frequency_Sys100NS
      WScript.Echo "Name: " & objItem.Name
      WScript.Echo "Timestamp_Object: " & objItem.Timestamp_Object
      WScript.Echo "Timestamp_PerfTime: " & objItem.Timestamp_PerfTime
      WScript.Echo "Timestamp_Sys100NS: " & objItem.Timestamp_Sys100NS
      WScript.Echo "WorkflowsAborted: " & objItem.WorkflowsAborted
      WScript.Echo "WorkflowsAbortedPersec: " & objItem.WorkflowsAbortedPersec
      WScript.Echo "WorkflowsCompleted: " & objItem.WorkflowsCompleted
      WScript.Echo "WorkflowsCompletedPersec: " & objItem.WorkflowsCompletedPersec
      WScript.Echo "WorkflowsCreated: " & objItem.WorkflowsCreated
      WScript.Echo "WorkflowsCreatedPersec: " & objItem.WorkflowsCreatedPersec
      WScript.Echo "WorkflowsExecuting: " & objItem.WorkflowsExecuting
      WScript.Echo "WorkflowsIdlePersec: " & objItem.WorkflowsIdlePersec
      WScript.Echo "WorkflowsInMemory: " & objItem.WorkflowsInMemory
      WScript.Echo "WorkflowsLoaded: " & objItem.WorkflowsLoaded
      WScript.Echo "WorkflowsLoadedPersec: " & objItem.WorkflowsLoadedPersec
      WScript.Echo "WorkflowsPending: " & objItem.WorkflowsPending
      WScript.Echo "WorkflowsPersisted: " & objItem.WorkflowsPersisted
      WScript.Echo "WorkflowsPersistedPersec: " & objItem.WorkflowsPersistedPersec
      WScript.Echo "WorkflowsRunnable: " & objItem.WorkflowsRunnable
      WScript.Echo "WorkflowsSuspended: " & objItem.WorkflowsSuspended
      WScript.Echo "WorkflowsSuspendedPersec: " & objItem.WorkflowsSuspendedPersec
      WScript.Echo "WorkflowsTerminated: " & objItem.WorkflowsTerminated
      WScript.Echo "WorkflowsTerminatedPersec: " & objItem.WorkflowsTerminatedPersec
      WScript.Echo "WorkflowsUnloaded: " & objItem.WorkflowsUnloaded
      WScript.Echo "WorkflowsUnloadedPersec: " & objItem.WorkflowsUnloadedPersec
      WScript.Echo
   Next
Next

