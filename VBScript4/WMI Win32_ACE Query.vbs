On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_ACE",,48)
For Each objItem in colItems
    Wscript.Echo "AccessMask: " & objItem.AccessMask
    Wscript.Echo "AceFlags: " & objItem.AceFlags
    Wscript.Echo "AceType: " & objItem.AceType
    Wscript.Echo "GuidInheritedObjectType: " & objItem.GuidInheritedObjectType
    Wscript.Echo "GuidObjectType: " & objItem.GuidObjectType
    Wscript.Echo "Trustee: " & objItem.Trustee
Next

