' Description: Uses cooked performance counters to return the count of the objects maintained by the operating system, including events, mutexes, processes, sections, semaphores, and threads.


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
set objRefresher = CreateObject("WbemScripting.SWbemRefresher")
Set colItems = objRefresher.AddEnum _
    (objWMIService, "Win32_PerfFormattedData_PerfOS_Objects").objectSet
objRefresher.Refresh

For i = 1 to 5
    For Each objItem in colItems
        Wscript.Echo "Caption: " & objItem.Caption
        Wscript.Echo "Description: " & objItem.Description
        Wscript.Echo "Events: " & objItem.Events
        Wscript.Echo "Mutexes: " & objItem.Mutexes
        Wscript.Echo "Name: " & objItem.Name
        Wscript.Echo "Processes: " & objItem.Processes
        Wscript.Echo "Sections: " & objItem.Sections
        Wscript.Echo "Semaphores: " & objItem.Semaphores
        Wscript.Echo "Threads: " & objItem.Threads
        objRefresher.Refresh
    Next
Next

