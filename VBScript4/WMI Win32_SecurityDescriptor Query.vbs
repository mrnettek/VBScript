On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_SecurityDescriptor",,48)
For Each objItem in colItems
    Wscript.Echo "ControlFlags: " & objItem.ControlFlags
    Wscript.Echo "DACL: " & objItem.DACL
    Wscript.Echo "Group: " & objItem.Group
    Wscript.Echo "Owner: " & objItem.Owner
    Wscript.Echo "SACL: " & objItem.SACL
Next

