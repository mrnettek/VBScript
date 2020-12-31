' Description: Returns information about the local groups found on a computer.


On Error Resume Next

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery _
    ("Select * from Win32_Group  Where LocalAccount = True")

For Each objItem in colItems
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "Domain: " & objItem.Domain
    Wscript.Echo "Local Account: " & objItem.LocalAccount
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "SID: " & objItem.SID
    Wscript.Echo "SID Type: " & objItem.SIDType
    Wscript.Echo "Status: " & objItem.Status
    Wscript.Echo
Next

