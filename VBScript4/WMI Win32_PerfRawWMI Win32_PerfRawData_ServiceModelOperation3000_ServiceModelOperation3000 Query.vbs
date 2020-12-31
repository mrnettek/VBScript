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
   Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_PerfRawData_ServiceModelOperation3000_ServiceModelOperation3000", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
      WScript.Echo "CallFailedPerSecond: " & objItem.CallFailedPerSecond
      WScript.Echo "Calls: " & objItem.Calls
      WScript.Echo "CallsDuration: " & objItem.CallsDuration
      WScript.Echo "CallsDuration_Base: " & objItem.CallsDuration_Base
      WScript.Echo "CallsFailed: " & objItem.CallsFailed
      WScript.Echo "CallsFaulted: " & objItem.CallsFaulted
      WScript.Echo "CallsFaultedPerSecond: " & objItem.CallsFaultedPerSecond
      WScript.Echo "CallsOutstanding: " & objItem.CallsOutstanding
      WScript.Echo "CallsPerSecond: " & objItem.CallsPerSecond
      WScript.Echo "Caption: " & objItem.Caption
      WScript.Echo "Description: " & objItem.Description
      WScript.Echo "Frequency_Object: " & objItem.Frequency_Object
      WScript.Echo "Frequency_PerfTime: " & objItem.Frequency_PerfTime
      WScript.Echo "Frequency_Sys100NS: " & objItem.Frequency_Sys100NS
      WScript.Echo "Name: " & objItem.Name
      WScript.Echo "SecurityCallsNotAuthorized: " & objItem.SecurityCallsNotAuthorized
      WScript.Echo "SecurityCallsNotAuthorizedPerSecond: " & objItem.SecurityCallsNotAuthorizedPerSecond
      WScript.Echo "SecurityValidationandAuthenticationFailures: " & objItem.SecurityValidationandAuthenticationFailures
      WScript.Echo "SecurityValidationandAuthenticationFailuresPerSecond: " & objItem.SecurityValidationandAuthenticationFailuresPerSecond
      WScript.Echo "Timestamp_Object: " & objItem.Timestamp_Object
      WScript.Echo "Timestamp_PerfTime: " & objItem.Timestamp_PerfTime
      WScript.Echo "Timestamp_Sys100NS: " & objItem.Timestamp_Sys100NS
      WScript.Echo "TransactionsFlowed: " & objItem.TransactionsFlowed
      WScript.Echo "TransactionsFlowedPerSecond: " & objItem.TransactionsFlowedPerSecond
      WScript.Echo
   Next
Next

