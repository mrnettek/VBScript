' Description: Retrieves page file usage statistics.


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colPageFiles = objWMIService.ExecQuery("Select * from Win32_PageFileUsage")

For Each objPageFile in colPageFiles
    Wscript.Echo "Allocated Base Size: " & objPageFile.AllocatedBaseSize
    Wscript.Echo "Current Usage: " & objPageFile.CurrentUsage
    Wscript.Echo "Description: " & objPageFile.Description
    Wscript.Echo "Install Date: " & objPageFile.InstallDate
    Wscript.Echo "Name: " & objPageFile.Name   
    Wscript.Echo "Peak Usage: " & objPageFile.PeakUsage 
Next

