On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from CIM_MemoryCapacity",,48)
For Each objItem in colItems
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "MaximumMemoryCapacity: " & objItem.MaximumMemoryCapacity
    Wscript.Echo "MemoryType: " & objItem.MemoryType
    Wscript.Echo "MinimumMemoryCapacity: " & objItem.MinimumMemoryCapacity
    Wscript.Echo "Name: " & objItem.Name
Next

