' Description: Returns starting and ending addresses for all memory devices installed on a computer.


On Error Resume Next

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery("Select * from Win32_MemoryDevice")

For Each objItem in colItems
    Wscript.Echo "Device ID: " & objItem.DeviceID
    Wscript.Echo "Ending Address: " & objItem.EndingAddress
    Wscript.Echo "Starting Address: " & objItem.StartingAddress
    Wscript.Echo
Next

