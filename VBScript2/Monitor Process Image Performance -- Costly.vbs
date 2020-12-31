' Description: Uses cooked performance counters to return information about the virtual address usage of images executed by computer processes.


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
set objRefresher = CreateObject("WbemScripting.SWbemRefresher")
Set colItems = objRefresher.AddEnum _
    (objWMIService, "Win32_PerfFormattedData_PerfProc_Image_Costly").objectSet
objRefresher.Refresh

For i = 1 to 5
    For Each objItem in colItems
        Wscript.Echo "Caption: " & objItem.Caption
        Wscript.Echo "Description: " & objItem.Description
        Wscript.Echo "Exec Read-Only: " & objItem.ExecReadOnly
        Wscript.Echo "Exec Read Per Write: " & objItem.ExecReadPerWrite
        Wscript.Echo "Executable: " & objItem.Executable
        Wscript.Echo "Exec Write Copy: " & objItem.ExecWriteCopy
        Wscript.Echo "Name: " & objItem.Name
        Wscript.Echo "No Access: " & objItem.NoAccess
        Wscript.Echo "Read-Only: " & objItem.ReadOnly
        Wscript.Echo "Read Per Write: " & objItem.ReadPerWrite
        Wscript.Echo "Write Copy: " & objItem.WriteCopy
        Wscript.Sleep 2000
        objRefresher.Refresh
    Next
Next

