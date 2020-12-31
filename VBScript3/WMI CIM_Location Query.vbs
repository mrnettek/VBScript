On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from CIM_Location",,48)
For Each objItem in colItems
    Wscript.Echo "Address: " & objItem.Address
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "PhysicalPosition: " & objItem.PhysicalPosition
Next

