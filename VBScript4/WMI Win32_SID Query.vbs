On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_SID",,48)
For Each objItem in colItems
    Wscript.Echo "AccountName: " & objItem.AccountName
    Wscript.Echo "BinaryRepresentation: " & objItem.BinaryRepresentation
    Wscript.Echo "ReferencedDomainName: " & objItem.ReferencedDomainName
    Wscript.Echo "SID: " & objItem.SID
    Wscript.Echo "SidLength: " & objItem.SidLength
Next

