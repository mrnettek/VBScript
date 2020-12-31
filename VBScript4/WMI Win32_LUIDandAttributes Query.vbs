On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_LUIDandAttributes",,48)
For Each objItem in colItems
    Wscript.Echo "Attributes: " & objItem.Attributes
    Wscript.Echo "LUID: " & objItem.LUID
Next

