On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_Registry",,48)
For Each objItem in colItems
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "CurrentSize: " & objItem.CurrentSize
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "InstallDate: " & objItem.InstallDate
    Wscript.Echo "MaximumSize: " & objItem.MaximumSize
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "ProposedSize: " & objItem.ProposedSize
    Wscript.Echo "Status: " & objItem.Status
Next

