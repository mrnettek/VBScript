On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_ThreadStopTrace",,48)
For Each objItem in colItems
    Wscript.Echo "ProcessID: " & objItem.ProcessID
    Wscript.Echo "SECURITY_DESCRIPTOR: " & objItem.SECURITY_DESCRIPTOR
    Wscript.Echo "ThreadID: " & objItem.ThreadID
    Wscript.Echo "TIME_CREATED: " & objItem.TIME_CREATED
Next

