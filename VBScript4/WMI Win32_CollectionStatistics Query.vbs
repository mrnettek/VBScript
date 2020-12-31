On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_CollectionStatistics",,48)
For Each objItem in colItems
    Wscript.Echo "Collection: " & objItem.Collection
    Wscript.Echo "Stats: " & objItem.Stats
Next

