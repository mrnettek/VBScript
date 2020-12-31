On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_SIDandAttributes",,48)
For Each objItem in colItems
    Wscript.Echo "Attributes: " & objItem.Attributes
    Wscript.Echo "SID: " & objItem.SID
Next

