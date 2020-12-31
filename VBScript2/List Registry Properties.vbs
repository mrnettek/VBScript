' Description: Returns information about the computer registry.


On Error Resume Next

strComputer = "."
Set objWMIService=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _ 
    strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery("Select * from Win32_Registry")

For Each objItem in colItems
    Wscript.Echo "Current Size: " & objItem.CurrentSize
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "Install Date: " & objItem.InstallDate
    Wscript.Echo "Maximum Size: " & objItem.MaximumSize
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "Proposed Size: " & objItem.ProposedSize
Next

