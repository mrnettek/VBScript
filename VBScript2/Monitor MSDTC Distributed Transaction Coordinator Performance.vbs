' Description: Uses cooked performance counters to measure Microsoft Distributed Transaction Coordinator performance.


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
set objRefresher = CreateObject("WbemScripting.SWbemRefresher")
Set colItems = objRefresher.AddEnum(objWMIService," & _
    "Win32_PerfFormattedData_MSDTC_DistributedTransactionCoordinator"). _
        objectSet
objRefresher.Refresh

For i = 1 to 5
    For Each objItem in colItems
        Wscript.Echo "Aborted Transactions: " & objItem.AbortedTransactions
        Wscript.Echo "Aborted Transactions Per Second: " & _
            objItem.AbortedTransactionsPersec
        Wscript.Echo "Active Transactions: " & objItem.ActiveTransactions
        Wscript.Echo "Active Transactions Maximum: " & _
            objItem.ActiveTransactionsMaximum
        Wscript.Echo "Caption: " & objItem.Caption
        Wscript.Echo "Committed Transactions: " & objItem.CommittedTransactions
        Wscript.Echo "Committed Transactions Per Second: " & _
            objItem.CommittedTransactionsPersec
        Wscript.Echo "Description: " & objItem.Description
        Wscript.Echo "Force Aborted Transactions: " & _
            objItem.ForceAbortedTransactions
        Wscript.Echo "Force Committed Transactions: " & _
            objItem.ForceCommittedTransactions
        Wscript.Echo "In-Doubt Transactions: " & objItem.InDoubtTransactions
        Wscript.Echo "Name: " & objItem.Name
        Wscript.Echo "Response Time Average: " & objItem.ResponseTimeAverage
        Wscript.Echo "Response Time Maximum: " & objItem.ResponseTimeMaximum
        Wscript.Echo "Response Time Minimum: " & objItem.ResponseTimeMinimum
        Wscript.Echo "Transactions Per Second: " & objItem.TransactionsPersec
        Wscript.Sleep 2000
        objRefresher.Refresh
    Next
Next

