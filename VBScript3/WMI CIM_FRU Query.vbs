On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from CIM_FRU",,48)
For Each objItem in colItems
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "FRUNumber: " & objItem.FRUNumber
    Wscript.Echo "IdentifyingNumber: " & objItem.IdentifyingNumber
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "RevisionLevel: " & objItem.RevisionLevel
    Wscript.Echo "Vendor: " & objItem.Vendor
Next

