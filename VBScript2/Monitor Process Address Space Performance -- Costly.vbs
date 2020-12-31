' Description: Uses cooked performance counters to return information about memory allocation and use for a selected process


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
set objRefresher = CreateObject("WbemScripting.SWbemRefresher")
Set colItems = objRefresher.AddEnum(objWMIService, _
    "Win32_PerfFormattedData_PerfProc_ProcessAddressSpace_Costly").objectSet
objRefresher.Refresh

For i = 1 to 5
    For Each objItem in colItems
        Wscript.Echo "Bytes Free: " & objItem.BytesFree
        Wscript.Echo "Bytes Image Free: " & objItem.BytesImageFree
        Wscript.Echo "Bytes Image Reserved: " & objItem.BytesImageReserved
        Wscript.Echo "Bytes Reserved: " & objItem.BytesReserved
        Wscript.Echo "Caption: " & objItem.Caption
        Wscript.Echo "Description: " & objItem.Description
        Wscript.Echo "ID Process: " & objItem.IDProcess
        Wscript.Echo "Image Space Exec Read-Only: " & _
            objItem.ImageSpaceExecReadOnly
        Wscript.Echo "Image Space Exec Read Per Write: " & _
            objItem.ImageSpaceExecReadPerWrite
        Wscript.Echo "Image Space Executable: " & objItem.ImageSpaceExecutable
        Wscript.Echo "Image Space Exec Write Copy: " & _
            objItem.ImageSpaceExecWriteCopy
        Wscript.Echo "Image Space NoA ccess: " & objItem.ImageSpaceNoAccess
        Wscript.Echo "Image Space Read-Only: " & objItem.ImageSpaceReadOnly
        Wscript.Echo "Image Space Read Per Write: " & _
            objItem.ImageSpaceReadPerWrite
        Wscript.Echo "Image Space Write Copy: " & objItem.ImageSpaceWriteCopy
        Wscript.Echo "Mapped Space Exec Read-Only: " & _
            objItem.MappedSpaceExecReadOnly
        Wscript.Echo "Mapped Space Exec Read Per Write: " & _
            objItem.MappedSpaceExecReadPerWrite
        Wscript.Echo "Mapped Space Executable: " & _
            objItem.MappedSpaceExecutable
        Wscript.Echo "Mapped Space Exec Write Copy: " & _
            objItem.MappedSpaceExecWriteCopy
        Wscript.Echo "Mapped Space No Access: " & objItem.MappedSpaceNoAccess
        Wscript.Echo "Mapped Space Read Only: " & objItem.MappedSpaceReadOnly
        Wscript.Echo "Mapped Space Read Per Write: " & _
            objItem.MappedSpaceReadPerWrite
        Wscript.Echo "Mapped Space Write Copy: " & objItem.MappedSpaceWriteCopy
        Wscript.Echo "Name: " & objItem.Name
        Wscript.Echo "Reserved Space Exec Read-Only: " & _
            objItem.ReservedSpaceExecReadOnly
        Wscript.Echo "Reserved Space Exec Read Per Write: " & _
            objItem.ReservedSpaceExecReadPerWrite
        Wscript.Echo "Reserved Space Executable: " & _
            objItem.ReservedSpaceExecutable
        Wscript.Echo "Reserved Space Exec Write Copy: " & _
            objItem.ReservedSpaceExecWriteCopy
        Wscript.Echo "Reserved Space No Access: " & _
            objItem.ReservedSpaceNoAccess
        Wscript.Echo "Reserved Space Read-Only: " & _
            objItem.ReservedSpaceReadOnly
        Wscript.Echo "Reserved Space Read Per Write: " & _
            objItem.ReservedSpaceReadPerWrite
        Wscript.Echo "Reserved Space Write Copy: " & _
            objItem.ReservedSpaceWriteCopy
        Wscript.Echo "Unassigned Space Exec Read-Only: " & _
            objItem.UnassignedSpaceExecReadOnly
        Wscript.Echo "Unassigned Space Exec Read Per Write: " & _
            objItem.UnassignedSpaceExecReadPerWrite
        Wscript.Echo "Unassigned Space Executable: " & _
            objItem.UnassignedSpaceExecutable
        Wscript.Echo "Unassigned Space Exec Write Copy: " & _
            objItem.UnassignedSpaceExecWriteCopy
        Wscript.Echo "Unassigned Space No Access: " & _
            objItem.UnassignedSpaceNoAccess
        Wscript.Echo "Unassigned Space Read-Only: " & _
            objItem.UnassignedSpaceReadOnly
        Wscript.Echo "Unassigned Space Read Per Write: " & _
            objItem.UnassignedSpaceReadPerWrite
        Wscript.Echo "Unassigned Space Write Copy: " & _
        objItem.UnassignedSpaceWriteCopy
        Wscript.Sleep 2000
        objRefresher.Refresh
    Next
Next

