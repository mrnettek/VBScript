' Description: Uses cooked performance counters to return information about print jobs spooled on a print server.


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
set objRefresher = CreateObject("WbemScripting.SWbemRefresher")
Set colItems = objRefresher.AddEnum _
    (objWMIService, "Win32_PerfFormattedData_Spooler_PrintQueue").objectSet
objRefresher.Refresh

For i = 1 to 5
    For Each objItem in colItems
        Wscript.Echo "Add Network Printer Calls: " & _
            objItem.AddNetworkPrinterCalls
        Wscript.Echo "Bytes Printed Per Second: " & objItem.BytesPrintedPersec
        Wscript.Echo "Caption: " & objItem.Caption
        Wscript.Echo "Description: " & objItem.Description
        Wscript.Echo "Enumerate Network Printer Calls: " & _     
        objItem.EnumerateNetworkPrinterCalls
        Wscript.Echo "Job Errors: " & objItem.JobErrors
        Wscript.Echo "Jobs: " & objItem.Jobs
        Wscript.Echo "Jobs Spooling: " & objItem.JobsSpooling
        Wscript.Echo "Maximum Jobs Spooling: " & objItem.MaxJobsSpooling
        Wscript.Echo "Maximum References: " & objItem.MaxReferences
        Wscript.Echo "Name: " & objItem.Name
        Wscript.Echo "Not-Ready Errors: " & objItem.NotReadyErrors
        Wscript.Echo "Out-of-Paper Errors: " & objItem.OutofPaperErrors
        Wscript.Echo "References: " & objItem.References
        Wscript.Echo "Total Jobs Printed: " & objItem.TotalJobsPrinted
        Wscript.Echo "Total Pages Printed: " & objItem.TotalPagesPrinted
        Wscript.Sleep 2000
        objRefresher.Refresh
    Next
Next

