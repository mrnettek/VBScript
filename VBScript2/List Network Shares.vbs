' Description: Lists all the shared folders on a computer.


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colShares = objWMIService.ExecQuery("Select * from Win32_Share")

For each objShare in colShares
    Wscript.Echo "Allow Maximum: " & objShare.AllowMaximum   
    Wscript.Echo "Caption: " & objShare.Caption   
    Wscript.Echo "Maximum Allowed: " & objShare.MaximumAllowed
    Wscript.Echo "Name: " & objShare.Name   
    Wscript.Echo "Path: " & objShare.Path   
    Wscript.Echo "Type: " & objShare.Type   
Next

