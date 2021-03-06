' Description: Enumerates all the Dfs targets on a computer.


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colDfsTargets = objWMIService.ExecQuery _
    ("Select * from Win32_DFSTarget")

For each objDfsTarget in colDfsTargets
    Wscript.Echo "Caption: " & objDfsTarget.Caption   
    Wscript.Echo "Description: " & objDfsTarget.Description
    Wscript.Echo "Install Date: " & objDfsTarget.InstallDate
    Wscript.Echo "Link Name: " & objDfsTarget.LinkName       
    Wscript.Echo "Name: " & objDfsTarget.Name 
    Wscript.Echo "Server Name: " & objDfsTarget.ServerName
    Wscript.Echo "Share Name: " & objDfsTarget.ShareName     
    Wscript.Echo "State: " & objDfsTarget.State       
    Wscript.Echo "Status: " & objDfsTarget.Status     
Next

