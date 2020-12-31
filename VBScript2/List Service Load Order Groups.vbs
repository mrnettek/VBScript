' Description: Returns a list of all the service load order groups found on a computer, and well as their load order.


On Error Resume Next

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery("Select * from Win32_LoadOrderGroup")

For Each objItem in colItems
    Wscript.Echo "Driver Enabled: " & objItem.DriverEnabled
    Wscript.Echo "Group Order: " & objItem.GroupOrder
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo
Next

