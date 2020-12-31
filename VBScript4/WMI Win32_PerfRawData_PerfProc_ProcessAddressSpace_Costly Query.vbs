On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_PerfRawData_PerfProc_ProcessAddressSpace_Costly",,48)
For Each objItem in colItems
    Wscript.Echo "BytesFree: " & objItem.BytesFree
    Wscript.Echo "BytesImageFree: " & objItem.BytesImageFree
    Wscript.Echo "BytesImageReserved: " & objItem.BytesImageReserved
    Wscript.Echo "BytesReserved: " & objItem.BytesReserved
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "Frequency_Object: " & objItem.Frequency_Object
    Wscript.Echo "Frequency_PerfTime: " & objItem.Frequency_PerfTime
    Wscript.Echo "Frequency_Sys100NS: " & objItem.Frequency_Sys100NS
    Wscript.Echo "IDProcess: " & objItem.IDProcess
    Wscript.Echo "ImageSpaceExecReadOnly: " & objItem.ImageSpaceExecReadOnly
    Wscript.Echo "ImageSpaceExecReadPerWrite: " & objItem.ImageSpaceExecReadPerWrite
    Wscript.Echo "ImageSpaceExecutable: " & objItem.ImageSpaceExecutable
    Wscript.Echo "ImageSpaceExecWriteCopy: " & objItem.ImageSpaceExecWriteCopy
    Wscript.Echo "ImageSpaceNoAccess: " & objItem.ImageSpaceNoAccess
    Wscript.Echo "ImageSpaceReadOnly: " & objItem.ImageSpaceReadOnly
    Wscript.Echo "ImageSpaceReadPerWrite: " & objItem.ImageSpaceReadPerWrite
    Wscript.Echo "ImageSpaceWriteCopy: " & objItem.ImageSpaceWriteCopy
    Wscript.Echo "MappedSpaceExecReadOnly: " & objItem.MappedSpaceExecReadOnly
    Wscript.Echo "MappedSpaceExecReadPerWrite: " & objItem.MappedSpaceExecReadPerWrite
    Wscript.Echo "MappedSpaceExecutable: " & objItem.MappedSpaceExecutable
    Wscript.Echo "MappedSpaceExecWriteCopy: " & objItem.MappedSpaceExecWriteCopy
    Wscript.Echo "MappedSpaceNoAccess: " & objItem.MappedSpaceNoAccess
    Wscript.Echo "MappedSpaceReadOnly: " & objItem.MappedSpaceReadOnly
    Wscript.Echo "MappedSpaceReadPerWrite: " & objItem.MappedSpaceReadPerWrite
    Wscript.Echo "MappedSpaceWriteCopy: " & objItem.MappedSpaceWriteCopy
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "ReservedSpaceExecReadOnly: " & objItem.ReservedSpaceExecReadOnly
    Wscript.Echo "ReservedSpaceExecReadPerWrite: " & objItem.ReservedSpaceExecReadPerWrite
    Wscript.Echo "ReservedSpaceExecutable: " & objItem.ReservedSpaceExecutable
    Wscript.Echo "ReservedSpaceExecWriteCopy: " & objItem.ReservedSpaceExecWriteCopy
    Wscript.Echo "ReservedSpaceNoAccess: " & objItem.ReservedSpaceNoAccess
    Wscript.Echo "ReservedSpaceReadOnly: " & objItem.ReservedSpaceReadOnly
    Wscript.Echo "ReservedSpaceReadPerWrite: " & objItem.ReservedSpaceReadPerWrite
    Wscript.Echo "ReservedSpaceWriteCopy: " & objItem.ReservedSpaceWriteCopy
    Wscript.Echo "Timestamp_Object: " & objItem.Timestamp_Object
    Wscript.Echo "Timestamp_PerfTime: " & objItem.Timestamp_PerfTime
    Wscript.Echo "Timestamp_Sys100NS: " & objItem.Timestamp_Sys100NS
    Wscript.Echo "UnassignedSpaceExecReadOnly: " & objItem.UnassignedSpaceExecReadOnly
    Wscript.Echo "UnassignedSpaceExecReadPerWrite: " & objItem.UnassignedSpaceExecReadPerWrite
    Wscript.Echo "UnassignedSpaceExecutable: " & objItem.UnassignedSpaceExecutable
    Wscript.Echo "UnassignedSpaceExecWriteCopy: " & objItem.UnassignedSpaceExecWriteCopy
    Wscript.Echo "UnassignedSpaceNoAccess: " & objItem.UnassignedSpaceNoAccess
    Wscript.Echo "UnassignedSpaceReadOnly: " & objItem.UnassignedSpaceReadOnly
    Wscript.Echo "UnassignedSpaceReadPerWrite: " & objItem.UnassignedSpaceReadPerWrite
    Wscript.Echo "UnassignedSpaceWriteCopy: " & objItem.UnassignedSpaceWriteCopy
Next

