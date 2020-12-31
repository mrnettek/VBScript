On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_PerfRawData_PerfProc_FullImage_Costly",,48)
For Each objItem in colItems
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "ExecReadOnly: " & objItem.ExecReadOnly
    Wscript.Echo "ExecReadPerWrite: " & objItem.ExecReadPerWrite
    Wscript.Echo "Executable: " & objItem.Executable
    Wscript.Echo "ExecWriteCopy: " & objItem.ExecWriteCopy
    Wscript.Echo "Frequency_Object: " & objItem.Frequency_Object
    Wscript.Echo "Frequency_PerfTime: " & objItem.Frequency_PerfTime
    Wscript.Echo "Frequency_Sys100NS: " & objItem.Frequency_Sys100NS
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "NoAccess: " & objItem.NoAccess
    Wscript.Echo "ReadOnly: " & objItem.ReadOnly
    Wscript.Echo "ReadPerWrite: " & objItem.ReadPerWrite
    Wscript.Echo "Timestamp_Object: " & objItem.Timestamp_Object
    Wscript.Echo "Timestamp_PerfTime: " & objItem.Timestamp_PerfTime
    Wscript.Echo "Timestamp_Sys100NS: " & objItem.Timestamp_Sys100NS
    Wscript.Echo "WriteCopy: " & objItem.WriteCopy
Next

