' Description: Returns a list of ProgIDs (Programmatic Identifiers) found on a computer.


On Error Resume Next

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * from Win32_ProgIDSpecification")

For Each objItem in colItems
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "Check ID: " & objItem.CheckID
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "Parent: " & objItem.Parent
    Wscript.Echo "ProgID: " & objItem.ProgID
Next

