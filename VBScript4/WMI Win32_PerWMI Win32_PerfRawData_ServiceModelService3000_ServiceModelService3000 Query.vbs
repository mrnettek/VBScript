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
   Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_PerfRawData_ServiceModelService3000_ServiceModelService3000", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
      WScript.Echo "Calls: " & objItem.Calls
      WScript.Echo "CallsDuration: " & objItem.CallsDuration
      WScript.Echo "CallsDuration_Base: " & objItem.CallsDuration_Base
      WScript.Echo "CallsFailed: " & objItem.CallsFailed
      WScript.Echo "CallsFailedPerSecond: " & objItem.CallsFailedPerSecond
      WScript.Echo "CallsFaulted: " & objItem.CallsFaulted
      WScript.Echo "CallsFaultedPerSecond: " & objItem.CallsFaultedPerSecond
      WScript.Echo "CallsOutstanding: " & objItem.CallsOutstanding
      WScript.Echo "CallsPerSecond: " & objItem.CallsPerSecond
      WScript.Echo "Caption: " & objItem.Caption
      WScript.Echo "Description: " & objItem.Description
      WScript.Echo "Frequency_Object: " & objItem.Frequency_Object
      WScript.Echo "Frequency_PerfTime: " & objItem.Frequency_PerfTime
      WScript.Echo "Frequency_Sys100NS: " & objItem.Frequency_Sys100NS
      WScript.Echo "Instances: " & objItem.Instances
      WScript.Echo "InstancesCreatedPerSecond: " & objItem.InstancesCreatedPerSecond
      WScript.Echo "Name: " & objItem.Name
      WScript.Echo "QueuedMessagesDropped: " & objItem.QueuedMessagesDropped
      WScript.Echo "QueuedMessagesDroppedPerSecond: " & objItem.QueuedMessagesDroppedPerSecond
      WScript.Echo "QueuedMessagesRejected: " & objItem.QueuedMessagesRejected
      WScript.Echo "QueuedMessagesRejectedPerSecond: " & objItem.QueuedMessagesRejectedPerSecond
      WScript.Echo "QueuedPoisonMessages: " & objItem.QueuedPoisonMessages
      WScript.Echo "QueuedPoisonMessagesPerSecond: " & objItem.QueuedPoisonMessagesPerSecond
      WScript.Echo "ReliableMessagingMessagesDropped: " & objItem.ReliableMessagingMessagesDropped
      WScript.Echo "ReliableMessagingMessagesDroppedPerSecond: " & objItem.ReliableMessagingMessagesDroppedPerSecond
      WScript.Echo "ReliableMessagingSessionsFaulted: " & objItem.ReliableMessagingSessionsFaulted
      WScript.Echo "ReliableMessagingSessionsFaultedPerSecond: " & objItem.ReliableMessagingSessionsFaultedPerSecond
      WScript.Echo "SecurityCallsNotAuthorized: " & objItem.SecurityCallsNotAuthorized
      WScript.Echo "SecurityCallsNotAuthorizedPerSecond: " & objItem.SecurityCallsNotAuthorizedPerSecond
      WScript.Echo "SecurityValidationandAuthenticationFailures: " & objItem.SecurityValidationandAuthenticationFailures
      WScript.Echo "SecurityValidationandAuthenticationFailuresPerSecond: " & objItem.SecurityValidationandAuthenticationFailuresPerSecond
      WScript.Echo "Timestamp_Object: " & objItem.Timestamp_Object
      WScript.Echo "Timestamp_PerfTime: " & objItem.Timestamp_PerfTime
      WScript.Echo "Timestamp_Sys100NS: " & objItem.Timestamp_Sys100NS
      WScript.Echo "TransactedOperationsAborted: " & objItem.TransactedOperationsAborted
      WScript.Echo "TransactedOperationsAbortedPerSecond: " & objItem.TransactedOperationsAbortedPerSecond
      WScript.Echo "TransactedOperationsCommitted: " & objItem.TransactedOperationsCommitted
      WScript.Echo "TransactedOperationsCommittedPerSecond: " & objItem.TransactedOperationsCommittedPerSecond
      WScript.Echo "TransactedOperationsInDoubt: " & objItem.TransactedOperationsInDoubt
      WScript.Echo "TransactedOperationsInDoubtPerSecond: " & objItem.TransactedOperationsInDoubtPerSecond
      WScript.Echo "TransactionsFlowed: " & objItem.TransactionsFlowed
      WScript.Echo "TransactionsFlowedPerSecond: " & objItem.TransactionsFlowedPerSecond
      WScript.Echo
   Next
Next

