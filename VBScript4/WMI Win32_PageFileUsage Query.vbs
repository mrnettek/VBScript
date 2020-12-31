On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_PageFileUsage",,48)
For Each objItem in colItems
    Wscript.Echo "AllocatedBaseSize: " & objItem.AllocatedBaseSize
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "CurrentUsage: " & objItem.CurrentUsage
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "InstallDate: " & objItem.InstallDate
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "PeakUsage: " & objItem.PeakUsage
    Wscript.Echo "Status: " & objItem.Status
    Wscript.Echo "TempPageFile: " & objItem.TempPageFile
Next

