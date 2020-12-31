' Description: Returns information about all the I/O ports found on a computer.


On Error Resume Next

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery("Select * from Win32_PortResource")

For Each objItem in colItems
    Wscript.Echo "Alias: " & objItem.Alias
    Wscript.Echo "Ending Address: " & objItem.EndingAddress
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "Starting Address: " & objItem.StartingAddress
    Wscript.Echo
Next

