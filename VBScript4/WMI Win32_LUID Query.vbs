On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_LUID",,48)
For Each objItem in colItems
    Wscript.Echo "HighPart: " & objItem.HighPart
    Wscript.Echo "LowPart: " & objItem.LowPart
Next

