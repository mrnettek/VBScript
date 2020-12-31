On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_TokenGroups",,48)
For Each objItem in colItems
    Wscript.Echo "GroupCount: " & objItem.GroupCount
    Wscript.Echo "Groups: " & objItem.Groups
Next

