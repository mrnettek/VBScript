' Description: Retrieves information about all the Start menu program groups currently in use on a computer.


On Error Resume Next

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery("Select * from Win32_ProgramGroup")

For Each objItem in colItems
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "Group Name: " & objItem.GroupName
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "User Name: " & objItem.UserName
    Wscript.Echo
Next

