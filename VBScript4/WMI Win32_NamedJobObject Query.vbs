On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_NamedJobObject",,48)
For Each objItem in colItems
    Wscript.Echo "BasicUIRestrictions: " & objItem.BasicUIRestrictions
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "CollectionID: " & objItem.CollectionID
    Wscript.Echo "Description: " & objItem.Description
Next

