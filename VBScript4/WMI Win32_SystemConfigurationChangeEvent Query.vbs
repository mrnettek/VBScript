On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_SystemConfigurationChangeEvent",,48)
For Each objItem in colItems
    Wscript.Echo "EventType: " & objItem.EventType
    Wscript.Echo "SECURITY_DESCRIPTOR: " & objItem.SECURITY_DESCRIPTOR
    Wscript.Echo "TIME_CREATED: " & objItem.TIME_CREATED
Next

