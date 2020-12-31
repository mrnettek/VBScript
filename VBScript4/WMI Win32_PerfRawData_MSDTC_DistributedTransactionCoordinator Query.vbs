On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_PerfRawData_MSDTC_DistributedTransactionCoordinator",,48)
For Each objItem in colItems
    Wscript.Echo "AbortedTransactions: " & objItem.AbortedTransactions
    Wscript.Echo "AbortedTransactionsPersec: " & objItem.AbortedTransactionsPersec
    Wscript.Echo "ActiveTransactions: " & objItem.ActiveTransactions
    Wscript.Echo "ActiveTransactionsMaximum: " & objItem.ActiveTransactionsMaximum
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "CommittedTransactions: " & objItem.CommittedTransactions
    Wscript.Echo "CommittedTransactionsPersec: " & objItem.CommittedTransactionsPersec
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "ForceAbortedTransactions: " & objItem.ForceAbortedTransactions
    Wscript.Echo "ForceCommittedTransactions: " & objItem.ForceCommittedTransactions
    Wscript.Echo "Frequency_Object: " & objItem.Frequency_Object
    Wscript.Echo "Frequency_PerfTime: " & objItem.Frequency_PerfTime
    Wscript.Echo "Frequency_Sys100NS: " & objItem.Frequency_Sys100NS
    Wscript.Echo "InDoubtTransactions: " & objItem.InDoubtTransactions
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "ResponseTimeAverage: " & objItem.ResponseTimeAverage
    Wscript.Echo "ResponseTimeMaximum: " & objItem.ResponseTimeMaximum
    Wscript.Echo "ResponseTimeMinimum: " & objItem.ResponseTimeMinimum
    Wscript.Echo "Timestamp_Object: " & objItem.Timestamp_Object
    Wscript.Echo "Timestamp_PerfTime: " & objItem.Timestamp_PerfTime
    Wscript.Echo "Timestamp_Sys100NS: " & objItem.Timestamp_Sys100NS
    Wscript.Echo "TransactionsPersec: " & objItem.TransactionsPersec
Next

