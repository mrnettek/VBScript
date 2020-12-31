' Description: Returns network information for a cluster server.


On Error Resume Next

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\mscluster")

Set colItems = objWMIService.ExecQuery("Select * from MSCluster_Network")

For Each objItem in colItems
    Wscript.Echo "Address: " & objItem.Address
    Wscript.Echo "Address mask: " & objItem.AddressMask
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "Characteristics: " & objItem.Characteristics
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "Flags: " & objItem.Flags
    Wscript.Echo "Installation date: " & objItem.InstallDate
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "Role: " & objItem.Role
    Wscript.Echo "State: " & objItem.State
    Wscript.Echo "Status: " & objItem.Status
    Wscript.Echo
Next

