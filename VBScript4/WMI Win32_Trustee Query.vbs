On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_Trustee",,48)
For Each objItem in colItems
    Wscript.Echo "Domain: " & objItem.Domain
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "SID: " & objItem.SID
    Wscript.Echo "SidLength: " & objItem.SidLength
    Wscript.Echo "SIDString: " & objItem.SIDString
Next

